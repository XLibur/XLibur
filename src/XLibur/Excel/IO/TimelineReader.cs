using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static XLibur.Excel.IO.ControlPartReading;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads timelines and their caches out of a package and binds them to the pivot tables they filter.
/// </summary>
/// <remarks>
/// <para>
/// <b>Nothing here attaches a DOM to a part.</b> Timeline parts survive a round trip today because
/// <c>xl/timelines/*.xml</c> and <c>xl/timelineCaches/*.xml</c> are never opened, so they are copied
/// through byte for byte with every attribute XLibur has no model for. Reaching a part through
/// <c>TimeLinePart.Timelines</c> or <c>TimeLineCachePart.TimelineCacheDefinition</c> would create an
/// attached DOM that the SDK tracks and writes back on save, replacing those bytes with its own
/// serialisation. Every read below therefore goes through an <see cref="OpenXmlPartReader"/>, which
/// streams the part and hands back a detached tree the part knows nothing about.
/// </para>
/// <para>
/// Reading runs after pivot tables have been loaded, because binding needs them.
/// </para>
/// </remarks>
internal static class TimelineReader
{
    internal static void LoadTimelines(WorkbookPart workbookPart, Sheets sheets, XLWorksheets worksheets)
    {
        var caches = ReadCaches(workbookPart);
        if (caches.Count == 0)
            return;

        BindCaches(caches.Values, worksheets);
        ReadTimelines(workbookPart, sheets, worksheets, caches);
    }

    // ── Caches ──────────────────────────────────────────────────────────

    private static Dictionary<string, XLTimelineCache> ReadCaches(WorkbookPart workbookPart)
    {
        var caches = new Dictionary<string, XLTimelineCache>(XLHelper.NameComparer);

        foreach (var cachePart in workbookPart.TimeLineCacheParts)
        {
            var definition = ReadDetached<X15.TimelineCacheDefinition>(cachePart);
            var name = definition?.Name?.Value;
            if (name is null)
                continue;

            caches[name] = ReadCache(definition!, name, workbookPart.GetIdOfPart(cachePart));
        }

        return caches;
    }

    private static XLTimelineCache ReadCache(
        X15.TimelineCacheDefinition definition, string name, string relId)
    {
        var cache = new XLTimelineCache(name, definition.SourceName?.Value ?? string.Empty)
        {
            WorkbookRelId = relId,
        };

        foreach (var pivotTable in definition.TimelineCachePivotTables?
                     .Elements<X15.TimelineCachePivotTable>() ?? [])
        {
            if (pivotTable.Name?.Value is { } pivotTableName)
                cache.PivotTableNames.Add(pivotTableName);
        }

        if (definition.TimelineState is not { } state)
            return cache;

        cache.PivotCacheId = state.PivotCacheId?.Value;

        // The SDK types filterType as an enumeration, but InnerText is the raw token, which is what
        // carries a value written by a newer Excel through a save unchanged.
        if (state.FilterType?.InnerText is { Length: > 0 } filterType)
            cache.FilterType = filterType;

        if (state.MinimalRefreshVersion?.Value is { } minimal)
            cache.MinimalRefreshVersion = minimal;

        if (state.LastRefreshVersion?.Value is { } last)
            cache.LastRefreshVersion = last;

        if (state.BoundsTimelineRange is { } bounds)
        {
            cache.BoundsStart = bounds.StartDate?.Value;
            cache.BoundsEnd = bounds.EndDate?.Value;
        }

        if (state.SelectionTimelineRange is { } selection)
        {
            cache.SelectionStart = selection.StartDate?.Value;
            cache.SelectionEnd = selection.EndDate?.Value;
        }

        return cache;
    }

    // ── Binding ─────────────────────────────────────────────────────────

    private static void BindCaches(IEnumerable<XLTimelineCache> caches, XLWorksheets worksheets)
    {
        var pivotTables = PivotTablesByName(worksheets);

        foreach (var cache in caches)
            BindPivotTables(cache, pivotTables);
    }

    // ── Timelines ───────────────────────────────────────────────────────

    private static void ReadTimelines(
        WorkbookPart workbookPart,
        Sheets sheets,
        XLWorksheets worksheets,
        Dictionary<string, XLTimelineCache> caches)
    {
        foreach (var (worksheetPart, worksheet) in WorksheetParts(workbookPart, sheets, worksheets))
        {
            foreach (var timelinePart in worksheetPart.TimeLineParts)
                ReadTimelinePart(worksheetPart, timelinePart, worksheet, caches);

            // Where each timeline sits is in the drawing part, not the timeline part. Read after the
            // timelines exist, because the frames are matched to them by name.
            if (worksheet.TimelinesInternal.Count > 0)
                TimelineAnchorXml.ReadPositions(worksheetPart.DrawingsPart, worksheet.TimelinesInternal);
        }
    }

    private static void ReadTimelinePart(
        WorksheetPart worksheetPart,
        TimeLinePart timelinePart,
        XLWorksheet worksheet,
        Dictionary<string, XLTimelineCache> caches)
    {
        var timelines = ReadDetached<X15.Timelines>(timelinePart);
        if (timelines is null)
            return;

        var relId = worksheetPart.GetIdOfPart(timelinePart);
        foreach (var timeline in timelines.Elements<X15.Timeline>())
            AddTimeline(timeline, relId, worksheet, caches);
    }

    private static void AddTimeline(
        X15.Timeline timeline,
        string relId,
        XLWorksheet worksheet,
        Dictionary<string, XLTimelineCache> caches)
    {
        var name = timeline.Name?.Value;
        var cacheName = timeline.Cache?.Value;
        if (name is null || cacheName is null || !caches.TryGetValue(cacheName, out var cache))
            return;

        var xlTimeline = new XLTimeline(worksheet, cache, name)
        {
            PartRelId = relId,
            SelectionLevelRaw = timeline.SelectionLevel?.Value,
            ScrollPosition = timeline.ScrollPosition?.Value,
        };

        // Seeded rather than assigned: going through the properties would mark every loaded timeline
        // as edited and bring parts nobody touched in for patching. The four booleans default to
        // true, which is what Excel means by omitting them.
        xlTimeline.SeedLoadedFormat(
            timeline.Caption?.Value ?? name,
            timeline.ShowHeader?.Value ?? true,
            timeline.ShowSelectionLabel?.Value ?? true,
            timeline.ShowTimeLevel?.Value ?? true,
            timeline.ShowHorizontalScrollbar?.Value ?? true,
            timeline.Style?.Value,
            timeline.Level?.Value ?? 0);

        worksheet.TimelinesInternal.Add(xlTimeline);
    }
}
