using System.Collections;
using System.Collections.Generic;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

internal sealed class XLPivotCaches : IXLPivotCaches, IEnumerable<XLPivotCache>, IWorkbookListener
{
    private readonly XLWorkbook _workbook;
    private readonly List<XLPivotCache> _caches = new();

    public XLPivotCaches(XLWorkbook workbook)
    {
        _workbook = workbook;
    }

    IXLPivotCache IXLPivotCaches.Add(IXLRange range) => Add(SheetArea.From(range));

    IEnumerator<IXLPivotCache> IEnumerable<IXLPivotCache>.GetEnumerator() => GetEnumerator();

    IEnumerator<XLPivotCache> IEnumerable<XLPivotCache>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public List<XLPivotCache>.Enumerator GetEnumerator() => _caches.GetEnumerator();

    internal XLPivotCache Add(SheetArea area)
    {
        var source = _workbook.TryGetTable(area, out var table)
            ? new XLPivotSourceReference(table.Name)
            : new XLPivotSourceReference(area);

        var newPivotCache = new XLPivotCache(source, _workbook);
        newPivotCache.Refresh();
        _caches.Add(newPivotCache);
        return newPivotCache;
    }

    internal XLPivotCache Add(IXLPivotSource source)
    {
        var newPivotCache = new XLPivotCache(source, _workbook);
        _caches.Add(newPivotCache);
        return newPivotCache;
    }

    /// <summary>
    /// Try to find an existing pivot cache for the passed area. The area
    /// is checked against both types of source references (tables and
    /// ranges) and if area matches, the cache is returned.
    /// </summary>
    internal XLPivotCache? Find(SheetArea area)
    {
        // This method mimics behavior of Excel.
        // If there is a table for the area and there is a cache for the table, return cache for the table.
        if (_workbook.TryGetTable(area, out var table))
        {
            // Table exists, so try to find it and match with the source reference.
            var tableSource = new XLPivotSourceReference(table.Name);
            foreach (var cache in _caches)
            {
                if (cache.Source.Equals(tableSource))
                    return cache;
            }
        }

        // Try to find a cache with area source.
        var areaSource = new XLPivotSourceReference(area);
        foreach (var cache in _caches)
        {
            if (cache.Source.Equals(areaSource))
                return cache;
        }

        return null;
    }

    /// <summary>
    /// A cache whose source is a range on the renamed sheet names the sheet by its new name, so the
    /// source resolves again and is saved as Excel saves it: the <c>rename-*</c> fixture writes
    /// <c>sheet="Renamed"</c> (D66). A source given by a table or a defined name follows that table or
    /// name, and a source in another workbook is not this workbook's sheet.
    /// </summary>
    void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
    {
        for (var i = 0; i < _caches.Count; i++)
        {
            var cache = _caches[i];
            if (cache.Source is not XLPivotSourceReference source || source.UsesName)
                continue;

            var area = source.Area.Value;
            if (XLHelper.SheetComparer.Equals(area.Name, oldSheetName))
                cache.Source = new XLPivotSourceReference(new SheetArea(newSheetName, area.Area));
        }
    }

    /// <summary>
    /// Nothing changes. The <c>delete-*</c> fixture shows Excel keeping a cache whose source was on the
    /// deleted sheet, with its records, its source as it was (<c>sheet="Data"</c>) and the pivot table
    /// on another sheet that uses it.
    /// </summary>
    void IWorkbookListener.OnSheetDeleting(string sheetName)
    {
        // Excel keeps the source as it was; see the summary.
    }
}
