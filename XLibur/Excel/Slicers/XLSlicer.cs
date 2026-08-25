using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace XLibur.Excel;

[DebuggerDisplay("{Name} ({SourceKind})")]
internal sealed class XLSlicer : IXLSlicer
{
    /// <summary>
    /// The button row height Excel writes for a new slicer, 247650 EMU.
    /// </summary>
    /// <remarks>
    /// <c>rowHeight</c> is a <em>required</em> attribute of <c>x14:slicer</c>, which is easy to miss
    /// because it reads like styling. A slicer written without it fails schema validation, so there
    /// has to be a value to fall back on rather than an absent attribute.
    /// </remarks>
    internal const double DefaultRowHeightPt = 19.5;

    private readonly XLWorksheet _worksheet;
    private string _caption;
    private bool _showCaption = true;
    private string? _style;
    private uint _columnCount = 1;
    private double? _rowHeightPt;

    internal XLSlicer(XLWorksheet worksheet, XLSlicerCache cache, string name)
    {
        _worksheet = worksheet;
        Cache = cache;
        Name = name;
        _caption = name;
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
    internal string? PartRelId { get; set; }

    /// <summary>
    /// Whether the slicer was created through the API rather than read from a package. A new slicer
    /// is generated on save; a loaded one is only ever patched.
    /// </summary>
    internal bool IsNew { get; set; }

    /// <summary>
    /// Which properties the caller has assigned since the slicer was loaded.
    /// </summary>
    internal XLSlicerFormat AssignedFormat { get; private set; }

    public string Name { get; }

    public string Caption
    {
        get => _caption;
        set
        {
            _caption = value ?? throw new ArgumentNullException(nameof(value));
            AssignedFormat |= XLSlicerFormat.Caption;
        }
    }

    public bool ShowCaption
    {
        get => _showCaption;
        set
        {
            _showCaption = value;
            AssignedFormat |= XLSlicerFormat.ShowCaption;
        }
    }

    public string? Style
    {
        get => _style;
        set
        {
            _style = value;
            AssignedFormat |= XLSlicerFormat.Style;
        }
    }

    public uint ColumnCount
    {
        get => _columnCount;
        set
        {
            if (value == 0)
                throw new ArgumentOutOfRangeException(nameof(value), "A slicer must have at least one column.");

            _columnCount = value;
            AssignedFormat |= XLSlicerFormat.ColumnCount;
        }
    }

    public double? RowHeightPt
    {
        get => _rowHeightPt;
        set
        {
            if (value is <= 0)
                throw new ArgumentOutOfRangeException(nameof(value), "A slicer's row height must be positive.");

            _rowHeightPt = value;
            AssignedFormat |= XLSlicerFormat.RowHeight;
        }
    }

    /// <summary>
    /// Sets the properties read from a package without marking them as assigned.
    /// </summary>
    /// <remarks>
    /// This is what keeps <see cref="AssignedFormat"/> honest. It has to stay the only way the
    /// reader populates a slicer: assigning through the properties instead would mark every loaded
    /// slicer as edited, and the patcher would then rewrite parts nobody touched.
    /// </remarks>
    internal void SeedLoadedFormat(
        string caption, bool showCaption, string? style, uint columnCount, double? rowHeightPt)
    {
        _caption = caption;
        _showCaption = showCaption;
        _style = style;
        _columnCount = columnCount;
        _rowHeightPt = rowHeightPt;
    }

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
