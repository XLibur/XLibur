using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using XLibur.Excel.Drawings;

namespace XLibur.Excel;

internal sealed class XLTimelines : IXLTimelines
{
    private readonly List<XLTimeline> _timelines = [];
    private readonly XLWorksheet _worksheet;

    internal XLTimelines(XLWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public int Count => _timelines.Count;

    internal IReadOnlyList<XLTimeline> Items => _timelines;

    public IEnumerator<IXLTimeline> GetEnumerator() => _timelines.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IXLTimeline Timeline(string name)
    {
        if (!TryGetTimeline(name, out var timeline))
            throw new KeyNotFoundException($"The worksheet has no timeline named '{name}'.");

        return timeline;
    }

    public bool TryGetTimeline(string name, [NotNullWhen(true)] out IXLTimeline? timeline)
    {
        foreach (var candidate in _timelines)
        {
            if (XLHelper.NameComparer.Equals(candidate.Name, name))
            {
                timeline = candidate;
                return true;
            }
        }

        timeline = null;
        return false;
    }

    internal void Add(XLTimeline timeline) => _timelines.Add(timeline);

    public IXLTimeline Add(IXLPivotTable pivotTable, string dateFieldName) =>
        AddTimeline((XLPivotTable)pivotTable, dateFieldName);

    internal XLTimeline AddTimeline(XLPivotTable pivotTable, string dateFieldName)
    {
        var pivotCache = (XLPivotCache)pivotTable.PivotCache;
        if (!pivotCache.TryGetFieldIndex(dateFieldName, out var fieldIndex))
        {
            throw new ArgumentException(
                $"The pivot cache of '{pivotTable.Name}' has no field named '{dateFieldName}'.",
                nameof(dateFieldName));
        }

        // Excel decides a field is timeline-able from the date statistics on its shared items. A
        // timeline over a field that holds no dates is a repair prompt, not a degraded timeline, so
        // it is refused here rather than written and discovered in Excel.
        var stats = pivotCache.GetFieldValues(fieldIndex).Stats;
        if (!stats.ContainsDate || stats.MinDate is not { } minDate || stats.MaxDate is not { } maxDate)
        {
            throw new ArgumentException(
                $"The field '{dateFieldName}' holds no dates, so it cannot carry a timeline.",
                nameof(dateFieldName));
        }

        var cache = new XLTimelineCache(NextCacheName(dateFieldName), dateFieldName)
        {
            IsNew = true,
            PivotCache = pivotCache,

            // Excel rounds the field's range outward to whole years — the round-trip fixture's field
            // runs 1998-05-19 to 2004-02-06 and its bounds read 1998-01-01 to 2005-01-01.
            BoundsStart = new DateTime(minDate.Year, 1, 1),
            BoundsEnd = new DateTime(maxDate.Year + 1, 1, 1),
        };
        cache.PivotTables.Add(pivotTable);
        cache.PivotTableNames.Add(pivotTable.Name);

        var area = pivotTable.Area;
        return AddNew(cache, dateFieldName, DefaultPositionBeside(area.FirstPoint.Row, area.LastPoint.Column));
    }

    private XLTimeline AddNew(XLTimelineCache cache, string sourceName, IXLCell position)
    {
        var name = NextTimelineName(sourceName);
        var timeline = new XLTimeline(_worksheet, cache, name)
        {
            IsNew = true,
            FromMarker = new XLMarker(position),
        };

        // Seeded rather than assigned, so a timeline created and left alone carries no pending
        // edits. Months is the level Excel starts a new timeline at.
        timeline.SeedLoadedFormat(
            name,
            showHeader: true,
            showSelectionLabel: true,
            showTimeLevel: true,
            showHorizontalScrollbar: true,
            style: null,
            level: (uint)XLTimelineLevel.Months);

        _timelines.Add(timeline);
        return timeline;
    }

    /// <summary>
    /// Where a new timeline goes when the caller has not said: two columns to the right of the pivot
    /// table it filters, at that table's top row.
    /// </summary>
    /// <remarks>
    /// A default is not optional here. <c>DrawingAnchorFactory</c> documents that a drawing handed no
    /// marker gets one at A1 — silently, with no exception and no missing element. For a picture
    /// that is reasonable. For a timeline it would drop the band over the top-left of the sheet,
    /// covering the very data it filters, and the caller would have no idea why. So every timeline
    /// XLibur creates is given a marker before the factory sees it, and that fallback stays
    /// unreachable from here.
    /// </remarks>
    private IXLCell DefaultPositionBeside(int topRow, int rightmostColumn) =>
        _worksheet.Cell(
            Math.Max(1, topRow),
            Math.Min(XLHelper.MaxColumnNumber, rightmostColumn + 2));

    /// <summary>
    /// A cache name not already taken, in the shape Excel uses: <c>NativeTimeline_Date</c>, then
    /// <c>NativeTimeline_Date1</c>.
    /// </summary>
    /// <remarks>
    /// The name is not decoration. A timeline refers to its cache by it, and a <c>#N/A</c> defined
    /// name is written under the same name, so it has to be a legal defined name: no spaces, and
    /// nothing that would parse as a cell reference.
    /// </remarks>
    private string NextCacheName(string sourceName)
    {
        var stem = "NativeTimeline_" + Sanitise(sourceName);
        var taken = WorkbookCacheNames();

        if (!taken.Contains(stem))
            return stem;

        for (var suffix = 1; ; suffix++)
        {
            var candidate = stem + suffix.ToString(CultureInfo.InvariantCulture);
            if (!taken.Contains(candidate))
                return candidate;
        }
    }

    /// <summary>
    /// A timeline name not already taken, in the shape Excel uses: <c>Date</c>, then <c>Date 1</c>.
    /// Timeline names are unique across the workbook, not just the sheet.
    /// </summary>
    private string NextTimelineName(string sourceName)
    {
        var taken = new HashSet<string>(XLHelper.NameComparer);
        foreach (var worksheet in _worksheet.Workbook.WorksheetsInternal)
        {
            foreach (var timeline in worksheet.TimelinesInternal.Items)
                taken.Add(timeline.Name);
        }

        if (!taken.Contains(sourceName))
            return sourceName;

        for (var suffix = 1; ; suffix++)
        {
            var candidate = sourceName + " " + suffix.ToString(CultureInfo.InvariantCulture);
            if (!taken.Contains(candidate))
                return candidate;
        }
    }

    private HashSet<string> WorkbookCacheNames()
    {
        var taken = new HashSet<string>(XLHelper.NameComparer);
        foreach (var worksheet in _worksheet.Workbook.WorksheetsInternal)
        {
            foreach (var timeline in worksheet.TimelinesInternal.Items)
                taken.Add(timeline.Cache.Name);
        }

        // A defined name already using the stem would collide with the one written for the cache.
        foreach (var definedName in _worksheet.Workbook.DefinedNamesInternal)
            taken.Add(definedName.Name);

        return taken;
    }

    private static string Sanitise(string sourceName)
    {
        var builder = new StringBuilder(sourceName.Length);
        foreach (var c in sourceName)
            builder.Append(char.IsLetterOrDigit(c) || c == '_' ? c : '_');

        return builder.Length > 0 ? builder.ToString() : "Field";
    }

    /// <summary>
    /// Drops a timeline from the worksheet and records what the save path has to unpick.
    /// </summary>
    /// <remarks>
    /// Removing a timeline is not a matter of dropping one element. Its cache part, the workbook's
    /// registration of that cache, the <c>#N/A</c> defined name written for it, the worksheet's
    /// <c>extLst</c> reference and the drawing anchor all have to go with it, or the saved file has
    /// an orphan Excel will offer to repair.
    /// </remarks>
    internal void Remove(XLTimeline timeline)
    {
        if (_timelines.Remove(timeline) && !timeline.IsNew)
            Removed.Add(timeline);
    }

    /// <summary>
    /// Timelines removed since the workbook was loaded, still holding the relationship ids and cache
    /// names the save path needs to clean up after them. Cleared once a save has consumed it.
    /// </summary>
    internal List<XLTimeline> Removed { get; } = [];
}
