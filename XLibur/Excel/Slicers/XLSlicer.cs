using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace XLibur.Excel;

[DebuggerDisplay("{Name} ({SourceKind})")]
internal sealed class XLSlicer : IXLSlicer
{
    private readonly XLWorksheet _worksheet;

    internal XLSlicer(XLWorksheet worksheet, XLSlicerCache cache, string name)
    {
        _worksheet = worksheet;
        Cache = cache;
        Name = name;
    }

    /// <summary>
    /// The cache that binds the slicer to what it filters and holds its selection.
    /// </summary>
    internal XLSlicerCache Cache { get; }

    /// <summary>
    /// The id of the relationship from the worksheet part to the slicers part this slicer was read
    /// from. Together with <see cref="Name"/>, which is unique within a part, this is how the write
    /// path will find the element again to patch it. Null for a slicer not read from a package.
    /// </summary>
    internal string? PartRelId { get; init; }

    public string Name { get; }

    public string Caption { get; internal init; } = string.Empty;

    public bool ShowCaption { get; internal init; } = true;

    public string? Style { get; internal init; }

    public uint ColumnCount { get; internal init; } = 1;

    public double? RowHeightPt { get; internal init; }

    public XLSlicerSourceKind SourceKind => Cache.SourceKind;

    public string SourceFieldName => Cache.SourceName;

    public IXLWorksheet Worksheet => _worksheet;

    public IReadOnlyList<IXLPivotTable> PivotTables => Cache.PivotTables;

    public IXLTable? Table => Cache.Table;

    public bool HasSelection => SourceKind == XLSlicerSourceKind.PivotTable
        ? Cache.Items.Any(i => i.Selected)
        : TableFilterValues() is not null;

    public IReadOnlyList<XLCellValue> SelectedItems => SourceKind == XLSlicerSourceKind.PivotTable
        ? PivotSelectedItems()
        : TableFilterValues() ?? [];

    /// <summary>
    /// Resolves the cache's selected item indices against the pivot cache field's shared items.
    /// </summary>
    /// <remarks>
    /// The indices are into the shared items of the field named by
    /// <see cref="XLSlicerCache.SourceName"/>, which is why a slicer cannot be read without the
    /// pivot cache behind it. A field whose shared items were not stored — Excel omits them for
    /// fields nothing indexes — leaves the selection unresolvable, and this reports nothing rather
    /// than guessing.
    /// </remarks>
    private List<XLCellValue> PivotSelectedItems()
    {
        var selected = new List<XLCellValue>();
        var cache = Cache.PivotCache;
        if (cache is null || !cache.TryGetFieldIndex(Cache.SourceName, out var fieldIndex))
            return selected;

        var sharedItems = cache.GetFieldSharedItems(fieldIndex);
        foreach (var item in Cache.Items)
        {
            if (item.Selected && item.Index < sharedItems.Count)
                selected.Add(sharedItems[item.Index]);
        }

        return selected;
    }

    /// <summary>
    /// The values a table slicer's column is filtered to, or null when it is not filtered by value.
    /// </summary>
    /// <remarks>
    /// A table slicer has no item list of its own: clicking its buttons applies a value filter to
    /// the bound column of the table's auto filter, and that is where the selection lives. A column
    /// carrying some other filter kind — custom, top ten, dynamic, by colour, none of which a
    /// slicer can produce but any of which a user can apply by hand — is reported as no selection
    /// rather than as an empty one.
    /// </remarks>
    private List<XLCellValue>? TableFilterValues()
    {
        if (Cache.Table is not { } table || Cache.TableColumnPosition is not { } position)
            return null;

        if (!table.AutoFilter.Columns.TryGetValue(position, out var filterColumn)
            || filterColumn.FilterType != XLFilterType.Regular)
        {
            return null;
        }

        var values = new List<XLCellValue>();
        foreach (var filter in filterColumn)
        {
            if (filter.Value is string text)
                values.Add(text);
        }

        return values.Count > 0 ? values : null;
    }
}
