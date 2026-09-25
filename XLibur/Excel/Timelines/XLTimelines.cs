using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
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
        var pivotCache = pivotTable.PivotCache;
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
            //
            // Unspecified, stated rather than defaulted: a serial date in a workbook is wall-clock
            // with no zone, so converting one is always wrong. DateTime.MaxValue below is
            // Unspecified for the same reason, which keeps both bounds on the same footing.
            //
            // S125 reads the DateTime.MaxValue sentence below as commented-out code.
#pragma warning disable S125
            // A field whose last date falls in year 9999 has no next-year boundary to round up to;
            // DateTime.MaxValue is the outermost bound there is, so it stands in rather than letting
            // the constructor throw on year 10000.
            BoundsStart = new DateTime(minDate.Year, 1, 1, 0, 0, 0, DateTimeKind.Unspecified),
#pragma warning restore S125
            BoundsEnd = maxDate.Year < 9999
                ? new DateTime(maxDate.Year + 1, 1, 1, 0, 0, 0, DateTimeKind.Unspecified)
                : DateTime.MaxValue,
        };
        cache.PivotTables.Add(pivotTable);
        cache.PivotTableNames.Add(pivotTable.Name);

        var area = pivotTable.Area;
        return AddNew(cache, dateFieldName, XLSheetControlNames.DefaultPositionBeside(
            _worksheet, area.FirstPoint.Row, area.LastPoint.Column));
    }

    private XLTimeline AddNew(XLTimelineCache cache, string sourceName, IXLCell position)
    {
        var name = XLSheetControlNames.NextControlName(_worksheet.Workbook, sourceName);
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
    /// A timeline cache name not already taken: <c>NativeTimeline_Date</c>, then
    /// <c>NativeTimeline_Date1</c>. Only timeline caches and defined names count as taken — see
    /// <see cref="XLSheetControlNames.NextCacheName"/>.
    /// </summary>
    private string NextCacheName(string sourceName) =>
        XLSheetControlNames.NextCacheName(
            _worksheet.Workbook,
            "NativeTimeline_",
            sourceName,
            static worksheet => worksheet.TimelinesInternal.Items.Select(timeline => timeline.Cache.Name));

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
