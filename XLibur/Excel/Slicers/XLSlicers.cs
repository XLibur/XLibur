using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using XLibur.Excel.Drawings;
using XLibur.Excel.Tables;

namespace XLibur.Excel;

internal sealed class XLSlicers : IXLSlicers
{
    private readonly List<XLSlicer> _slicers = [];
    private readonly XLWorksheet _worksheet;

    internal XLSlicers(XLWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public int Count => _slicers.Count;

    internal IReadOnlyList<XLSlicer> Items => _slicers;

    public IEnumerator<IXLSlicer> GetEnumerator() => _slicers.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IXLSlicer Slicer(string name)
    {
        if (!TryGetSlicer(name, out var slicer))
            throw new KeyNotFoundException($"The worksheet has no slicer named '{name}'.");

        return slicer;
    }

    public bool TryGetSlicer(string name, [NotNullWhen(true)] out IXLSlicer? slicer)
    {
        foreach (var candidate in _slicers)
        {
            if (XLHelper.NameComparer.Equals(candidate.Name, name))
            {
                slicer = candidate;
                return true;
            }
        }

        slicer = null;
        return false;
    }

    internal void Add(XLSlicer slicer) => _slicers.Add(slicer);

    public IXLSlicer Add(IXLPivotTable pivotTable, string fieldName) =>
        AddPivotSlicer((XLPivotTable)pivotTable, fieldName);

    public IXLSlicer Add(IXLTable table, string columnName) =>
        AddTableSlicer((XLTable)table, columnName);

    /// <summary>
    /// Creates a slicer that filters a pivot table on one of its cache fields.
    /// </summary>
    internal XLSlicer AddPivotSlicer(XLPivotTable pivotTable, string fieldName)
    {
        var cache = pivotTable.PivotCache;
        if (!cache.TryGetFieldIndex(fieldName, out var fieldIndex))
        {
            throw new ArgumentException(
                $"The pivot cache of '{pivotTable.Name}' has no field named '{fieldName}'.", nameof(fieldName));
        }

        var slicerCache = new XLSlicerCache(NextCacheName(fieldName), fieldName, XLSlicerSourceKind.PivotTable)
        {
            IsNew = true,
            PivotCache = cache,
        };
        slicerCache.PivotTables.Add(pivotTable);
        slicerCache.PivotTableNames.Add(pivotTable.Name);

        // A slicer nobody has clicked filters nothing, and in the file that is every item selected
        // rather than an absent item list. Populating it here rather than at save time keeps the
        // model agreeing with itself before and after a round trip.
        var sharedItems = cache.GetFieldSharedItems(fieldIndex);
        for (var i = 0; i < sharedItems.Count; i++)
            slicerCache.Items.Add(new XLSlicerCacheItem((uint)i, Selected: true));

        var area = pivotTable.Area;
        return AddNew(slicerCache, fieldName, XLSheetControlNames.DefaultPositionBeside(
            _worksheet, area.FirstPoint.Row, area.LastPoint.Column));
    }

    /// <summary>
    /// Creates a slicer that filters a table on one of its columns.
    /// </summary>
    /// <remarks>
    /// A table slicer holds no item list of its own — it drives the bound column of the table's auto
    /// filter — so a new one has nothing to seed.
    /// </remarks>
    internal XLSlicer AddTableSlicer(XLTable table, string columnName)
    {
        var position = 0;
        var found = 0;
        foreach (var field in table.Fields)
        {
            position++;
            if (XLHelper.NameComparer.Equals(field.Name, columnName))
            {
                found = position;
                break;
            }
        }

        if (found == 0)
            throw new ArgumentException($"The table '{table.Name}' has no column named '{columnName}'.", nameof(columnName));

        // Both ids are assigned while the parts are written, not modelled, so only the position is
        // knowable now. The writer fills in TableId and TableColumnId from what it actually emits.
        var slicerCache = new XLSlicerCache(NextCacheName(columnName), columnName, XLSlicerSourceKind.Table)
        {
            IsNew = true,
            Table = table,
            TableColumnPosition = found,
        };

        return AddNew(slicerCache, columnName, XLSheetControlNames.DefaultPositionBeside(
            _worksheet,
            table.RangeAddress.FirstAddress.RowNumber,
            table.RangeAddress.LastAddress.ColumnNumber));
    }

    private XLSlicer AddNew(XLSlicerCache cache, string sourceName, IXLCell position)
    {
        var name = XLSheetControlNames.NextControlName(_worksheet.Workbook, sourceName);
        var slicer = new XLSlicer(_worksheet, cache, name) { IsNew = true, FromMarker = new XLMarker(position) };

        // Seeded rather than assigned, so a slicer created and left alone carries no pending edits.
        // rowHeight is required by the schema, so a new slicer starts at Excel's own default rather
        // than with nothing.
        slicer.SeedLoadedFormat(name, showCaption: true, style: null, columnCount: 1, XLSlicer.DefaultRowHeightPt);

        _slicers.Add(slicer);
        return slicer;
    }

    /// <summary>
    /// A slicer cache name not already taken: <c>Slicer_Region</c>, then <c>Slicer_Region1</c>.
    /// Only slicer caches and defined names count as taken — see
    /// <see cref="XLSheetControlNames.NextCacheName"/>.
    /// </summary>
    private string NextCacheName(string sourceName) =>
        XLSheetControlNames.NextCacheName(
            _worksheet.Workbook,
            "Slicer_",
            sourceName,
            static worksheet => worksheet.SlicersInternal.Items.Select(slicer => slicer.Cache.Name));

    /// <summary>
    /// Drops a slicer from the worksheet and records what the save path has to unpick.
    /// </summary>
    /// <remarks>
    /// Removing a slicer is not a matter of dropping one element. Its cache part, the workbook's
    /// registration of that cache, the <c>#N/A</c> defined name Excel writes for it and the
    /// worksheet's own <c>extLst</c> reference all have to go with it, or the saved file has an
    /// orphan Excel will offer to repair. The name is kept here so the writers can find what to
    /// remove once the package is open.
    /// </remarks>
    internal void Remove(XLSlicer slicer)
    {
        if (_slicers.Remove(slicer) && !slicer.IsNew)
            Removed.Add(slicer);
    }

    /// <summary>
    /// Slicers removed since the workbook was loaded, still holding the relationship ids and cache
    /// names the save path needs to clean up after them. Cleared once a save has consumed it.
    /// </summary>
    internal List<XLSlicer> Removed { get; } = [];
}
