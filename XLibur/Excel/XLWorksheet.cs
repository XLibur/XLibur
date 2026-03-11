using XLibur.Excel.Caching;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.Drawings;
using XLibur.Excel.Ranges.Index;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XLibur.Excel.InsertData;
using static XLibur.Excel.XLProtectionAlgorithm;

namespace XLibur.Excel;

internal sealed class XLWorksheet : XLRangeBase, IXLWorksheet
{
    #region Fields

    private readonly Dictionary<int, int> _columnOutlineCount = new();
    private readonly Dictionary<int, int> _rowOutlineCount = new();
    private readonly XLRangeFactory _rangeFactory;
    private readonly XLRangeRepository _rangeRepository;
    private readonly List<IXLRangeIndex> _rangeIndices;
    private readonly XLRanges _selectedRanges;

    internal int ZOrder = 1;
    private string _name;
    internal int _position;

    private double _rowHeight;
    private bool _tabActive;
    private XLSheetProtection _protection;

    /// <summary>
    /// Fake address to be used everywhere the invalid address is needed.
    /// </summary>
    internal readonly XLAddress InvalidAddress;

    #endregion Fields

    #region Constructor

    public XLWorksheet(string sheetName, XLWorkbook workbook, uint sheetId)
        : base(
            new XLRangeAddress(
                new XLAddress(null, XLHelper.MinRowNumber, XLHelper.MinColumnNumber, false, false),
                new XLAddress(null, XLHelper.MaxRowNumber, XLHelper.MaxColumnNumber, false, false)),
            ((XLStyle)workbook.Style).Value)
    {
        Workbook = workbook;
        SheetId = sheetId;
        InvalidAddress = new XLAddress(this, 0, 0, false, false);

        var firstAddress = new XLAddress(this, RangeAddress.FirstAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber,
            RangeAddress.FirstAddress.FixedRow, RangeAddress.FirstAddress.FixedColumn);
        var lastAddress = new XLAddress(this, RangeAddress.LastAddress.RowNumber, RangeAddress.LastAddress.ColumnNumber,
            RangeAddress.LastAddress.FixedRow, RangeAddress.LastAddress.FixedColumn);
        RangeAddress = new XLRangeAddress(firstAddress, lastAddress);
        _rangeFactory = new XLRangeFactory(this);
        _rangeRepository = new XLRangeRepository(workbook, _rangeFactory.Create);
        _rangeIndices = new List<IXLRangeIndex>();

        Pictures = new XLPictures(this);
        DefinedNames = new XLDefinedNames(this);
        SheetView = new XLSheetView(this);
        Tables = new XLTables();
        Hyperlinks = new XLHyperlinks(this);
        DataValidations = new XLDataValidations(this);
        PivotTables = new XLPivotTables(this);
        _protection = new XLSheetProtection(DefaultProtectionAlgorithm);
        AutoFilter = new XLAutoFilter();
        ConditionalFormats = new XLConditionalFormats();
        SparklineGroupsInternal = new XLSparklineGroups(this);
        Internals = new XLWorksheetInternals(new XLCellsCollection(this), new XLColumnsCollection(),
            new XLRowsCollection(), new XLRanges());
        PageSetup = new XLPageSetup((XLPageSetup)workbook.PageOptions, this);
        Outline = new XLOutline(workbook.Outline);
        _columnWidth = workbook.ColumnWidth;
        _rowHeight = workbook.RowHeight;
        RowHeightChanged = Math.Abs(workbook.RowHeight - XLWorkbook.DefaultRowHeight) > XLHelper.Epsilon;

        XLHelper.ValidateSheetName(sheetName);
        _name = sheetName;
        Charts = new XLCharts();
        ShowFormulas = workbook.ShowFormulas;
        ShowGridLines = workbook.ShowGridLines;
        ShowOutlineSymbols = workbook.ShowOutlineSymbols;
        ShowRowColHeaders = workbook.ShowRowColHeaders;
        ShowRuler = workbook.ShowRuler;
        ShowWhiteSpace = workbook.ShowWhiteSpace;
        ShowZeros = workbook.ShowZeros;
        RightToLeft = workbook.RightToLeft;
        TabColor = XLColor.NoColor;
        _selectedRanges = new XLRanges();

        Author = workbook.Author;
    }

    #endregion Constructor

    [Obsolete($"Use {nameof(DefinedNames)} instead.")]
    IXLDefinedNames IXLWorksheet.NamedRanges => DefinedNames;

    IXLDefinedNames IXLWorksheet.DefinedNames => DefinedNames;

    internal XLDefinedNames DefinedNames { get; }

    public override XLRangeType RangeType
    {
        get { return XLRangeType.Worksheet; }
    }

    /// <summary>
    /// Reference to a VML that contains notes, forms controls and so on. All such things are generally unified into
    /// a single legacy VML file, set during load/save.
    /// </summary>
    public string? LegacyDrawingId;

    private double _columnWidth;

    public XLWorksheetInternals Internals { get; private set; }

    internal XLSparklineGroups SparklineGroupsInternal { get; }

    public XLRangeFactory RangeFactory
    {
        get { return _rangeFactory; }
    }

    protected override IEnumerable<XLStylizedBase> Children
    {
        get
        {
            var columnsUsed = Internals.ColumnsCollection.Keys
                .Union(Internals.CellsCollection.ColumnsUsedKeys)
                .Distinct()
                .OrderBy(c => c)
                .ToList();
            foreach (var col in columnsUsed)
                yield return Column(col);

            var rowsUsed = Internals.RowsCollection.Keys
                .Union(Internals.CellsCollection.RowsUsedKeys)
                .Distinct()
                .OrderBy(r => r)
                .ToList();
            foreach (var row in rowsUsed)
                yield return Row(row);
        }
    }

    internal bool RowHeightChanged { get; set; }

    internal bool ColumnWidthChanged { get; set; }

    /// <summary>
    /// <para>
    /// The id of a sheet that is unique across the workbook, kept across load/save.
    /// The ids of sheets are not reused. That is important for referencing the sheet
    /// range/point through sheetId. If sheetIds were reused, references would refer
    /// to the wrong sheet after the original sheetId was reused. Excel also doesn't
    /// reuse sheetIds.
    /// </para>
    /// <para>
    /// Referencing sheet through non-reused sheetIds means that reference can survive
    /// sheet renaming without any changes. Always &gt; 0 (Excel will crash on 0).
    /// </para>
    /// </summary>
    internal uint SheetId { get; set; }

    /// <summary>
    /// A cached <c>r:id</c> of the sheet from the file. If the sheet is a new one (not
    /// yet saved), the value is null until workbook is saved. Use <see cref="SheetId"/>
    /// instead is possible. Mostly for removing deleted sheet parts during save.
    /// </summary>
    internal string? RelId { get; set; }

    public XLDataValidations DataValidations { get; private set; }

    public IXLCharts Charts { get; private set; }

    public XLSheetProtection Protection
    {
        get => _protection;
        set => _protection = value.Clone().CastTo<XLSheetProtection>();
    }

    public XLAutoFilter AutoFilter { get; private set; }

    public bool IsDeleted { get; private set; }

    #region IXLWorksheet Members

    public XLWorkbook Workbook { get; private set; }

    public double ColumnWidth
    {
        get { return _columnWidth; }
        set
        {
            ColumnWidthChanged = true;
            _columnWidth = value;
        }
    }

    public double RowHeight
    {
        get { return _rowHeight; }
        set
        {
            RowHeightChanged = true;
            _rowHeight = value;
        }
    }

    public string Name
    {
        get { return _name; }
        set
        {
            if (_name == value) return;

            XLHelper.ValidateSheetName(value);

            Workbook.WorksheetsInternal.Rename(_name, value);
            _name = value;
        }
    }

    public int Position
    {
        get { return _position; }
        set
        {
            if (value > Workbook.WorksheetsInternal.Count + Workbook.UnsupportedSheets.Count + 1)
                throw new ArgumentOutOfRangeException(nameof(value), "Index must be equal or less than the number of worksheets + 1.");

            if (value < _position)
            {
                Workbook.WorksheetsInternal
                    .Where<XLWorksheet>(w => w.Position >= value && w.Position < _position)
                    .ForEach(w => w._position += 1);
            }

            if (value > _position)
            {
                Workbook.WorksheetsInternal
                    .Where<XLWorksheet>(w => w.Position <= value && w.Position > _position)
                    .ForEach(w => (w)._position -= 1);
            }

            _position = value;
        }
    }

    public IXLPageSetup PageSetup { get; private set; }

    public IXLOutline Outline { get; private set; }

    IXLRow? IXLWorksheet.FirstRowUsed()
    {
        return FirstRowUsed();
    }

    IXLRow? IXLWorksheet.FirstRowUsed(XLCellsUsedOptions options)
    {
        return FirstRowUsed(options);
    }

    IXLRow? IXLWorksheet.LastRowUsed()
    {
        return LastRowUsed();
    }

    IXLRow? IXLWorksheet.LastRowUsed(XLCellsUsedOptions options)
    {
        return LastRowUsed(options);
    }

    IXLColumn IXLWorksheet.LastColumn()
    {
        return LastColumn();
    }

    IXLColumn IXLWorksheet.FirstColumn()
    {
        return FirstColumn();
    }

    IXLRow IXLWorksheet.FirstRow()
    {
        return FirstRow();
    }

    IXLRow IXLWorksheet.LastRow()
    {
        return LastRow();
    }

    IXLColumn? IXLWorksheet.FirstColumnUsed()
    {
        return FirstColumnUsed();
    }

    IXLColumn? IXLWorksheet.FirstColumnUsed(XLCellsUsedOptions options)
    {
        return FirstColumnUsed(options);
    }

    IXLColumn? IXLWorksheet.LastColumnUsed()
    {
        return LastColumnUsed();
    }

    IXLColumn? IXLWorksheet.LastColumnUsed(XLCellsUsedOptions options)
    {
        return LastColumnUsed(options);
    }

    public IXLColumns Columns()
    {
        var columnMap = new HashSet<int>();

        columnMap.UnionWith(Internals.CellsCollection.ColumnsUsedKeys);
        columnMap.UnionWith(Internals.ColumnsCollection.Keys);

        var retVal = new XLColumns(this, StyleValue, columnMap.Select(Column));
        return retVal;
    }

    public IXLColumns Columns(string columns)
    {
        var retVal = new XLColumns(null, StyleValue);
        var columnPairs = columns.Split(',');
        foreach (var tPair in columnPairs.Select(pair => pair.Trim()))
        {
            string firstColumn;
            string lastColumn;
            if (tPair.Contains(':') || tPair.Contains('-'))
            {
                var columnRange = XLHelper.SplitRange(tPair);
                firstColumn = columnRange[0];
                lastColumn = columnRange[1];
            }
            else
            {
                firstColumn = tPair;
                lastColumn = tPair;
            }

            if (int.TryParse(firstColumn, out _))
            {
                foreach (var col in Columns(int.Parse(firstColumn), int.Parse(lastColumn)))
                    retVal.Add((XLColumn)col);
            }
            else
            {
                foreach (var col in Columns(firstColumn, lastColumn))
                    retVal.Add((XLColumn)col);
            }
        }
        return retVal;
    }

    public IXLColumns Columns(string firstColumn, string lastColumn)
    {
        return Columns(XLHelper.GetColumnNumberFromLetter(firstColumn),
            XLHelper.GetColumnNumberFromLetter(lastColumn));
    }

    public IXLColumns Columns(int firstColumn, int lastColumn)
    {
        var retVal = new XLColumns(null, StyleValue);

        for (var co = firstColumn; co <= lastColumn; co++)
            retVal.Add(Column(co));
        return retVal;
    }

    public IXLRows Rows()
    {
        var rowMap = new HashSet<int>();

        rowMap.UnionWith(Internals.CellsCollection.RowsUsedKeys);
        rowMap.UnionWith(Internals.RowsCollection.Keys);

        var retVal = new XLRows(this, StyleValue, rowMap.Select(Row));
        return retVal;
    }

    public IXLRows Rows(string rows)
    {
        var retVal = new XLRows(null, StyleValue);
        var rowPairs = rows.Split(',');
        foreach (var tPair in rowPairs.Select(pair => pair.Trim()))
        {
            string firstRow;
            string lastRow;
            if (tPair.Contains(':') || tPair.Contains('-'))
            {
                var rowRange = XLHelper.SplitRange(tPair);
                firstRow = rowRange[0];
                lastRow = rowRange[1];
            }
            else
            {
                firstRow = tPair;
                lastRow = tPair;
            }

            Rows(int.Parse(firstRow), int.Parse(lastRow))
                .ForEach(row => retVal.Add((XLRow)row));
        }
        return retVal;
    }

    public IXLRows Rows(int firstRow, int lastRow)
    {
        var retVal = new XLRows(null, StyleValue);

        for (var ro = firstRow; ro <= lastRow; ro++)
            retVal.Add(Row(ro));
        return retVal;
    }

    IXLRow IXLWorksheet.Row(int row)
    {
        return Row(row);
    }

    IXLColumn IXLWorksheet.Column(int column)
    {
        return Column(column);
    }

    IXLColumn IXLWorksheet.Column(string column)
    {
        return Column(column);
    }

    IXLCell IXLWorksheet.Cell(int row, int column)
    {
        return Cell(row, column);
    }

    IXLCell IXLWorksheet.Cell(string cellAddressInRange)
    {
        return Cell(cellAddressInRange) ?? throw new ArgumentException($"'{cellAddressInRange}' is not A1 address or workbook named range.");
    }

    IXLCell IXLWorksheet.Cell(int row, string column)
    {
        return Cell(row, column);
    }

    IXLCell IXLWorksheet.Cell(IXLAddress cellAddressInRange)
    {
        return Cell(cellAddressInRange);
    }

    IXLRange IXLWorksheet.Range(IXLRangeAddress rangeAddress)
    {
        return Range(rangeAddress);
    }

    IXLRange IXLWorksheet.Range(string rangeAddress)
    {
        return Range(rangeAddress) ?? throw new ArgumentException($"'{rangeAddress}' is not A1 address or named range.");
    }

    IXLRange IXLWorksheet.Range(IXLCell firstCell, IXLCell lastCell)
    {
        return Range(firstCell, lastCell);
    }

    IXLRange IXLWorksheet.Range(string firstCellAddress, string lastCellAddress)
    {
        return Range(firstCellAddress, lastCellAddress);
    }

    IXLRange IXLWorksheet.Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
    {
        return Range(firstCellAddress, lastCellAddress);
    }

    IXLRange IXLWorksheet.Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
    {
        return Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn);
    }

    IXLRanges IXLWorksheet.Ranges(string ranges) => Ranges(ranges);

    public IXLWorksheet CollapseRows()
    {
        Enumerable.Range(1, 8).ForEach(i => CollapseRows(i));
        return this;
    }

    public IXLWorksheet CollapseColumns()
    {
        Enumerable.Range(1, 8).ForEach(i => CollapseColumns(i));
        return this;
    }

    public IXLWorksheet ExpandRows()
    {
        Enumerable.Range(1, 8).ForEach(i => ExpandRows(i));
        return this;
    }

    public IXLWorksheet ExpandColumns()
    {
        Enumerable.Range(1, 8).ForEach(i => ExpandColumns(i));
        return this;
    }

    public IXLWorksheet CollapseRows(int outlineLevel)
    {
        if (outlineLevel is < 1 or > 8)
            throw new ArgumentOutOfRangeException(nameof(outlineLevel), "Outline level must be between 1 and 8.");

        Internals.RowsCollection.Values.Where(r => r.OutlineLevel == outlineLevel).ForEach(r => r.Collapse());
        return this;
    }

    public IXLWorksheet CollapseColumns(int outlineLevel)
    {
        if (outlineLevel is < 1 or > 8)
            throw new ArgumentOutOfRangeException(nameof(outlineLevel), "Outline level must be between 1 and 8.");

        Internals.ColumnsCollection.Values.Where(c => c.OutlineLevel == outlineLevel).ForEach(c => c.Collapse());
        return this;
    }

    public IXLWorksheet ExpandRows(int outlineLevel)
    {
        if (outlineLevel < 1 || outlineLevel > 8)
            throw new ArgumentOutOfRangeException(nameof(outlineLevel), "Outline level must be between 1 and 8.");

        Internals.RowsCollection.Values.Where(r => r.OutlineLevel == outlineLevel).ForEach(r => r.Expand());
        return this;
    }

    public IXLWorksheet ExpandColumns(int outlineLevel)
    {
        if (outlineLevel is < 1 or > 8)
            throw new ArgumentOutOfRangeException(nameof(outlineLevel), "Outline level must be between 1 and 8.");

        Internals.ColumnsCollection.Values.Where(c => c.OutlineLevel == outlineLevel).ForEach(c => c.Expand());
        return this;
    }

    public void Delete()
    {
        IsDeleted = true;
        Workbook.DefinedNamesInternal.OnWorksheetDeleted(Name);
        Workbook.NotifyWorksheetDeleting(this);
        Workbook.WorksheetsInternal.Delete(Name);
    }


    [Obsolete($"Used {nameof(DefinedName)} instead.")]
    IXLDefinedName IXLWorksheet.NamedRange(string name) => DefinedName(name);

    IXLDefinedName IXLWorksheet.DefinedName(string name) => DefinedName(name);

    internal XLDefinedName DefinedName(string name)
    {
        return DefinedNames.DefinedName(name);
    }

    IXLSheetView IXLWorksheet.SheetView { get => SheetView; }

    public XLSheetView SheetView { get; private set; }

    IXLTables IXLWorksheet.Tables => Tables;

    internal XLTables Tables { get; }

    public IXLTable Table(int index)
    {
        return Tables.Table(index);
    }

    public IXLTable Table(string name)
    {
        return Tables.Table(name);
    }

    public IXLWorksheet CopyTo(string newSheetName)
    {
        return CopyTo(Workbook, newSheetName, Workbook.WorksheetsInternal.Count + 1);
    }

    public IXLWorksheet CopyTo(string newSheetName, int position)
    {
        return CopyTo(Workbook, newSheetName, position);
    }

    public IXLWorksheet CopyTo(XLWorkbook workbook)
    {
        return CopyTo(workbook, Name, workbook.WorksheetsInternal.Count + 1);
    }

    public IXLWorksheet CopyTo(XLWorkbook workbook, string newSheetName)
    {
        return CopyTo(workbook, newSheetName, workbook.WorksheetsInternal.Count + 1);
    }

    public IXLWorksheet CopyTo(XLWorkbook workbook, string newSheetName, int position)
    {
        if (IsDeleted)
            throw new InvalidOperationException($"`{Name}` has been deleted and cannot be copied.");

        var targetSheet = (XLWorksheet)workbook.WorksheetsInternal.Add(newSheetName, position);
        Internals.ColumnsCollection.ForEach(kp => kp.Value.CopyTo(targetSheet.Column(kp.Key)));
        Internals.RowsCollection.ForEach(kp => kp.Value.CopyTo(targetSheet.Row(kp.Key)));
        Internals.CellsCollection.GetCells().ForEach(c => targetSheet.Cell(c.Address).CopyFrom(c, XLCellCopyOptions.Values | XLCellCopyOptions.Styles));
        DataValidations.ForEach(dv => targetSheet.DataValidations.Add(new XLDataValidation(dv, this)));
        targetSheet.Visibility = Visibility;
        targetSheet.ColumnWidth = ColumnWidth;
        targetSheet.ColumnWidthChanged = ColumnWidthChanged;
        targetSheet.RowHeight = RowHeight;
        targetSheet.RowHeightChanged = RowHeightChanged;
        targetSheet.InnerStyle = InnerStyle;
        targetSheet.PageSetup = new XLPageSetup((XLPageSetup)PageSetup, targetSheet);
        ((XLHeaderFooter)targetSheet.PageSetup.Header).Changed = true;
        ((XLHeaderFooter)targetSheet.PageSetup.Footer).Changed = true;
        targetSheet.Outline = new XLOutline(Outline);
        targetSheet.SheetView = new XLSheetView(targetSheet, SheetView);
        targetSheet.SelectedRanges.RemoveAll();

        Pictures.ForEach(picture => picture.CopyTo(targetSheet));
        Tables.ForEach<XLTable>(t => t.CopyTo(targetSheet, false));
        DefinedNames.ForEach<XLDefinedName>(nr => nr.CopyTo(targetSheet)); // Names must modify table references, so keep the order.
        PivotTables.ForEach<XLPivotTable>(pt => pt.CopyTo(targetSheet.Cell(pt.TargetCell.Address.CastTo<XLAddress>().WithoutWorksheet())));
        ConditionalFormats.ForEach(cf => cf.CopyTo(targetSheet));
        SparklineGroups.CopyTo(targetSheet);
        MergedRanges.ForEach(mr => targetSheet.Range(((XLRangeAddress)mr.RangeAddress).WithoutWorksheet()).Merge());
        SelectedRanges.ForEach(sr => targetSheet.SelectedRanges.Add(targetSheet.Range(((XLRangeAddress)sr.RangeAddress).WithoutWorksheet())));

        if (AutoFilter.IsEnabled)
        {
            var range = targetSheet.Range(((XLRangeAddress)AutoFilter.Range.RangeAddress).WithoutWorksheet());
            range.SetAutoFilter();
        }

        return targetSheet;
    }

    internal XLHyperlinks Hyperlinks { get; }

    IXLHyperlinks IXLWorksheet.Hyperlinks => Hyperlinks;

    IXLDataValidations IXLWorksheet.DataValidations => DataValidations;

    public XLWorksheetVisibility Visibility
    {
        get;
        set
        {
            if (value != XLWorksheetVisibility.Visible)
                TabSelected = false;

            field = value;
        }
    }

    public IXLWorksheet Hide()
    {
        Visibility = XLWorksheetVisibility.Hidden;
        return this;
    }

    public IXLWorksheet Unhide()
    {
        Visibility = XLWorksheetVisibility.Visible;
        return this;
    }

    IXLSheetProtection IXLProtectable<IXLSheetProtection, XLSheetProtectionElements>.Protection
    {
        get => Protection;
        set => Protection = (XLSheetProtection)value;
    }

    public IXLSheetProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm)
    {
        return Protection.Protect(algorithm);
    }

    public IXLSheetProtection Protect(XLSheetProtectionElements allowedElements)
        => Protection.Protect(allowedElements);

    public IXLSheetProtection Protect(Algorithm algorithm, XLSheetProtectionElements allowedElements)
        => Protection.Protect(algorithm, allowedElements);

    public IXLSheetProtection Protect(string password, Algorithm algorithm = DefaultProtectionAlgorithm)
    {
        return Protection.Protect(password, algorithm);
    }

    public IXLSheetProtection Protect(string password, Algorithm algorithm, XLSheetProtectionElements allowedElements)
    {
        return Protection.Protect(password, algorithm, allowedElements);
    }

    IXLElementProtection IXLProtectable.Protect(Algorithm algorithm)
    {
        return Protect(algorithm);
    }

    IXLElementProtection IXLProtectable.Protect(string password, Algorithm algorithm)
    {
        return Protect(password, algorithm);
    }

    public IXLSheetProtection Unprotect()
    {
        return Protection.Unprotect();
    }

    public IXLSheetProtection Unprotect(string password)
    {
        return Protection.Unprotect(password);
    }

    IXLElementProtection IXLProtectable.Unprotect()
    {
        return Unprotect();
    }

    IXLElementProtection IXLProtectable.Unprotect(string password)
    {
        return Unprotect(password);
    }

    public new IXLRange Sort()
    {
        return GetRangeForSort().Sort();
    }

    public new IXLRange Sort(string columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false, bool ignoreBlanks = true)
    {
        return GetRangeForSort().Sort(columnsToSortBy, sortOrder, matchCase, ignoreBlanks);
    }

    public new IXLRange Sort(int columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false, bool ignoreBlanks = true)
    {
        return GetRangeForSort().Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
    }

    public new IXLRange SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false,
        bool ignoreBlanks = true)
    {
        return GetRangeForSort().SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
    }

    public bool ShowFormulas { get; set; }

    public bool ShowGridLines { get; set; }

    public bool ShowOutlineSymbols { get; set; }

    public bool ShowRowColHeaders { get; set; }

    public bool ShowRuler { get; set; }

    public bool ShowWhiteSpace { get; set; }

    public bool ShowZeros { get; set; }

    public IXLWorksheet SetShowFormulas()
    {
        ShowFormulas = true;
        return this;
    }

    public IXLWorksheet SetShowFormulas(bool value)
    {
        ShowFormulas = value;
        return this;
    }

    public IXLWorksheet SetShowGridLines()
    {
        ShowGridLines = true;
        return this;
    }

    public IXLWorksheet SetShowGridLines(bool value)
    {
        ShowGridLines = value;
        return this;
    }

    public IXLWorksheet SetShowOutlineSymbols()
    {
        ShowOutlineSymbols = true;
        return this;
    }

    public IXLWorksheet SetShowOutlineSymbols(bool value)
    {
        ShowOutlineSymbols = value;
        return this;
    }

    public IXLWorksheet SetShowRowColHeaders()
    {
        ShowRowColHeaders = true;
        return this;
    }

    public IXLWorksheet SetShowRowColHeaders(bool value)
    {
        ShowRowColHeaders = value;
        return this;
    }

    public IXLWorksheet SetShowRuler()
    {
        ShowRuler = true;
        return this;
    }

    public IXLWorksheet SetShowRuler(bool value)
    {
        ShowRuler = value;
        return this;
    }

    public IXLWorksheet SetShowWhiteSpace()
    {
        ShowWhiteSpace = true;
        return this;
    }

    public IXLWorksheet SetShowWhiteSpace(bool value)
    {
        ShowWhiteSpace = value;
        return this;
    }

    public IXLWorksheet SetShowZeros()
    {
        ShowZeros = true;
        return this;
    }

    public IXLWorksheet SetShowZeros(bool value)
    {
        ShowZeros = value;
        return this;
    }

    public XLColor TabColor { get; set; }

    public IXLWorksheet SetTabColor(XLColor color)
    {
        TabColor = color;
        return this;
    }

    public bool TabSelected { get; set; }

    public bool TabActive
    {
        get { return _tabActive; }
        set
        {
            if (value && !_tabActive)
            {
                foreach (var ws in Worksheet.Workbook.WorksheetsInternal)
                    ws._tabActive = false;
            }
            _tabActive = value;
        }
    }

    public IXLWorksheet SetTabSelected()
    {
        TabSelected = true;
        return this;
    }

    public IXLWorksheet SetTabSelected(bool value)
    {
        TabSelected = value;
        return this;
    }

    public IXLWorksheet SetTabActive()
    {
        TabActive = true;
        return this;
    }

    public IXLWorksheet SetTabActive(bool value)
    {
        TabActive = value;
        return this;
    }

    IXLPivotTable IXLWorksheet.PivotTable(string name)
    {
        return PivotTable(name);
    }

    IXLPivotTables IXLWorksheet.PivotTables => PivotTables;

    public XLPivotTables PivotTables { get; }

    public bool RightToLeft { get; set; }

    public IXLWorksheet SetRightToLeft()
    {
        RightToLeft = true;
        return this;
    }

    public IXLWorksheet SetRightToLeft(bool value)
    {
        RightToLeft = value;
        return this;
    }

    public override XLRanges Ranges(string ranges)
    {
        var retVal = new XLRanges();
        foreach (var rangeAddressStr in ranges.Split(',').Select(s => s.Trim()))
        {
            if (rangeAddressStr.StartsWith("#REF!"))
            {
                continue;
            }

            if (XLHelper.IsValidRangeAddress(rangeAddressStr))
            {
                retVal.Add(Range(new XLRangeAddress(Worksheet, rangeAddressStr)));
            }
            else if (DefinedNames.TryGetValue(rangeAddressStr, out var worksheetNamedRange))
            {
                worksheetNamedRange.Ranges.ForEach(retVal.Add);
            }
            else if (Workbook.DefinedNames.TryGetValue(rangeAddressStr, out var workbookDefinedName)
                     && workbookDefinedName.Ranges.First().Worksheet == this)
            {
                workbookDefinedName.Ranges.ForEach(retVal.Add);
            }
        }
        return retVal;
    }

    IXLAutoFilter IXLWorksheet.AutoFilter
    {
        get { return AutoFilter; }
    }

    public IXLRows RowsUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents, Func<IXLRow, bool>? predicate = null)
    {
        var rows = new XLRows(worksheet: null, StyleValue);
        var rowsUsed = new HashSet<int>();
        foreach (var rowNum in Internals.RowsCollection.Keys.Concat(Internals.CellsCollection.RowsUsedKeys))
        {
            if (!rowsUsed.Add(rowNum))
            {
                continue;
            }
            var row = Row(rowNum);
            if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
                rows.Add(row);
        }
        return rows;
    }

    public IXLRows RowsUsed(Func<IXLRow, bool>? predicate = null)
    {
        return RowsUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    public IXLColumns ColumnsUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents, Func<IXLColumn, bool>? predicate = null)
    {
        var columns = new XLColumns(worksheet: null, StyleValue);
        var columnsUsed = new HashSet<int>();
        Internals.ColumnsCollection.Keys.ForEach(r => columnsUsed.Add(r));
        Internals.CellsCollection.ColumnsUsedKeys.ForEach(r => columnsUsed.Add(r));
        foreach (var columnNum in columnsUsed)
        {
            var column = Column(columnNum);
            if (!column.IsEmpty(options) && (predicate == null || predicate(column)))
                columns.Add(column);
        }
        return columns;
    }

    public IXLColumns ColumnsUsed(Func<IXLColumn, bool>? predicate = null)
    {
        return ColumnsUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    internal void RegisterRangeIndex(IXLRangeIndex rangeIndex)
    {
        _rangeIndices.Add(rangeIndex);
    }

    internal void Cleanup()
    {
        Internals.Dispose();
        Pictures.ForEach(p => p.Dispose());
        _rangeRepository.Clear();
        _rangeIndices.Clear();
    }

    #endregion IXLWorksheet Members

    #region Outlines

    public void IncrementColumnOutline(int level)
    {
        if (level <= 0) return;
        if (_columnOutlineCount.TryGetValue(level, out var value))
            _columnOutlineCount[level] = value + 1;
        else
            _columnOutlineCount.Add(level, 1);
    }

    public void DecrementColumnOutline(int level)
    {
        if (level <= 0) return;
        if (_columnOutlineCount.TryGetValue(level, out var value))
        {
            if (value > 0)
                _columnOutlineCount[level] = value - 1;
        }
        else
            _columnOutlineCount.Add(level, 0);
    }

    public int GetMaxColumnOutline()
    {
        var list = _columnOutlineCount.Where(kp => kp.Value > 0).ToList();
        return list.Count == 0 ? 0 : list.Max(kp => kp.Key);
    }

    public void IncrementRowOutline(int level)
    {
        if (level <= 0) return;
        if (_rowOutlineCount.TryGetValue(level, out var value))
            _rowOutlineCount[level] = value + 1;
        else
            _rowOutlineCount.Add(level, 0);
    }

    public void DecrementRowOutline(int level)
    {
        if (level <= 0) return;
        if (_rowOutlineCount.TryGetValue(level, out var value))
        {
            if (value > 0)
                _rowOutlineCount[level] = level - 1;
        }
        else
            _rowOutlineCount.Add(level, 0);
    }

    public int GetMaxRowOutline()
    {
        return _rowOutlineCount.Count == 0 ? 0 : _rowOutlineCount.Where(kp => kp.Value > 0).Max(kp => kp.Key);
    }

    #endregion Outlines

    public XLRow? FirstRowUsed()
    {
        return FirstRowUsed(XLCellsUsedOptions.AllContents);
    }

    public XLRow? FirstRowUsed(XLCellsUsedOptions options)
    {
        var rngRow = AsRange().FirstRowUsed(options);
        return rngRow != null ? Row(rngRow.RangeAddress.FirstAddress.RowNumber) : null;
    }

    public XLRow? LastRowUsed()
    {
        return LastRowUsed(XLCellsUsedOptions.AllContents);
    }

    public XLRow? LastRowUsed(XLCellsUsedOptions options)
    {
        var rngRow = AsRange().LastRowUsed(options);
        return rngRow != null ? Row(rngRow.RangeAddress.LastAddress.RowNumber) : null;
    }

    public XLColumn LastColumn()
    {
        return Column(XLHelper.MaxColumnNumber);
    }

    public XLColumn FirstColumn()
    {
        return Column(1);
    }

    public XLRow FirstRow()
    {
        return Row(1);
    }

    public XLRow LastRow()
    {
        return Row(XLHelper.MaxRowNumber);
    }

    public XLColumn? FirstColumnUsed()
    {
        return FirstColumnUsed(XLCellsUsedOptions.AllContents);
    }

    public XLColumn? FirstColumnUsed(XLCellsUsedOptions options)
    {
        var rngColumn = AsRange().FirstColumnUsed(options);
        return rngColumn != null ? Column(rngColumn.RangeAddress.FirstAddress.ColumnNumber) : null;
    }

    public XLColumn? LastColumnUsed()
    {
        return LastColumnUsed(XLCellsUsedOptions.AllContents);
    }

    public XLColumn? LastColumnUsed(XLCellsUsedOptions options)
    {
        var rngColumn = AsRange().LastColumnUsed(options);
        return rngColumn != null ? Column(rngColumn.RangeAddress.LastAddress.ColumnNumber) : null;
    }

    public XLRow Row(int row)
    {
        return Row(row, true);
    }

    public XLColumn Column(int columnNumber)
    {
        if (columnNumber <= 0 || columnNumber > XLHelper.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(columnNumber), $"Column number must be between 1 and {XLHelper.MaxColumnNumber}");

        if (Internals.ColumnsCollection.TryGetValue(columnNumber, out var column))
            return column;
        // This is a new column so we're going to reference all
        // cells in this column to preserve their formatting
        Internals.RowsCollection.Keys.ForEach(r => Cell(r, columnNumber).PingStyle());

        column = RangeFactory.CreateColumn(columnNumber);
        Internals.ColumnsCollection.Add(columnNumber, column);

        return column;
    }

    public IXLColumn Column(string column)
    {
        return Column(XLHelper.GetColumnNumberFromLetter(column));
    }

    public override XLRange AsRange()
    {
        return Range(1, 1, XLHelper.MaxRowNumber, XLHelper.MaxColumnNumber);
    }

    internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
    {
        if (!range.IsEntireColumn())
        {
            var model = new XLRangeAddress(
                range.RangeAddress.FirstAddress,
                new XLAddress(range.RangeAddress.LastAddress.RowNumber, XLHelper.MaxColumnNumber, false, false));
            var rangesToSplit = Worksheet.MergedRanges
                .GetIntersectedRanges(model)
                .Where(r => r.RangeAddress.FirstAddress.RowNumber < range.RangeAddress.FirstAddress.RowNumber ||
                            r.RangeAddress.LastAddress.RowNumber > range.RangeAddress.LastAddress.RowNumber)
                .ToList();
            foreach (var rangeToSplit in rangesToSplit)
            {
                Worksheet.MergedRanges.Remove(rangeToSplit);
            }
        }

        Workbook.WorksheetsInternal.ForEach<XLWorksheet>(ws => MoveDefinedNamesColumns(range, columnsShifted, ws.DefinedNames));
        MoveDefinedNamesColumns(range, columnsShifted, Workbook.DefinedNamesInternal);
        ShiftConditionalFormattingColumns(range, columnsShifted);
        ShiftDataValidationColumns(range, columnsShifted);
        ShiftPageBreaksColumns(range, columnsShifted);
        RemoveInvalidSparklines();

        ISheetListener hyperlinks = Hyperlinks;
        if (columnsShifted > 0)
        {
            var area = XLSheetRange
                .FromRangeAddress(range.RangeAddress)
                .ExtendRight(columnsShifted - 1);
            Workbook.CalcEngine.OnInsertAreaAndShiftRight(range.Worksheet, area);
            hyperlinks.OnInsertAreaAndShiftRight(range.Worksheet, area);
        }
        else if (columnsShifted < 0)
        {
            var area = XLSheetRange.FromRangeAddress(range.RangeAddress);
            Workbook.CalcEngine.OnDeleteAreaAndShiftLeft(range.Worksheet, area);
            hyperlinks.OnDeleteAreaAndShiftLeft(range.Worksheet, area);
        }
    }

    private void ShiftPageBreaksColumns(XLRange range, int columnsShifted)
    {
        for (var i = 0; i < PageSetup.ColumnBreaks.Count; i++)
        {
            var br = PageSetup.ColumnBreaks[i];
            if (range.RangeAddress.FirstAddress.ColumnNumber <= br)
            {
                PageSetup.ColumnBreaks[i] = br + columnsShifted;
            }
        }
    }

    private void ShiftConditionalFormattingColumns(XLRange range, int columnsShifted)
    {
        if (!ConditionalFormats.Any()) return;
        var firstCol = range.RangeAddress.FirstAddress.ColumnNumber;
        if (firstCol == 1) return;

        var colNum = columnsShifted > 0 ? firstCol - 1 : firstCol;
        var model = Column(colNum).AsRange();

        foreach (var cf in ConditionalFormats.ToList())
        {
            var cfRanges = cf.Ranges.ToList();
            cf.Ranges.RemoveAll();

            foreach (var cfRange in cfRanges)
            {
                var cfAddress = cfRange.RangeAddress;
                IXLRange newRange;
                if (cfRange.Intersects(model))
                {
                    newRange = Range(cfAddress.FirstAddress.RowNumber,
                        cfAddress.FirstAddress.ColumnNumber,
                        cfAddress.LastAddress.RowNumber,
                        Math.Min(XLHelper.MaxColumnNumber, cfAddress.LastAddress.ColumnNumber + columnsShifted));
                }
                else if (cfAddress.FirstAddress.ColumnNumber >= firstCol)
                {
                    newRange = Range(cfAddress.FirstAddress.RowNumber,
                        Math.Max(cfAddress.FirstAddress.ColumnNumber + columnsShifted, firstCol),
                        cfAddress.LastAddress.RowNumber,
                        Math.Min(XLHelper.MaxColumnNumber, cfAddress.LastAddress.ColumnNumber + columnsShifted));
                }
                else
                    newRange = cfRange;

                if (newRange.RangeAddress.IsValid &&
                    newRange.RangeAddress.FirstAddress.ColumnNumber <=
                    newRange.RangeAddress.LastAddress.ColumnNumber)
                    cf.Ranges.Add(newRange);
            }

            if (!cf.Ranges.Any())
                ConditionalFormats.Remove(f => f == cf);
        }
    }

    private void ShiftDataValidationColumns(XLRange range, int columnsShifted)
    {
        if (!DataValidations.Any()) return;
        var firstCol = range.RangeAddress.FirstAddress.ColumnNumber;
        if (firstCol == 1) return;

        var colNum = columnsShifted > 0 ? firstCol - 1 : firstCol;
        var model = Column(colNum).AsRange();

        foreach (var dv in DataValidations.ToList())
        {
            var dvRanges = dv.Ranges.ToList();
            dv.ClearRanges();

            foreach (var dvRange in dvRanges)
            {
                var dvAddress = dvRange.RangeAddress;
                IXLRange newRange;
                if (dvRange.Intersects(model))
                {
                    newRange = Range(dvAddress.FirstAddress.RowNumber,
                        dvAddress.FirstAddress.ColumnNumber,
                        dvAddress.LastAddress.RowNumber,
                        Math.Min(XLHelper.MaxColumnNumber, dvAddress.LastAddress.ColumnNumber + columnsShifted));
                }
                else if (dvAddress.FirstAddress.ColumnNumber >= firstCol)
                {
                    newRange = Range(dvAddress.FirstAddress.RowNumber,
                        Math.Max(dvAddress.FirstAddress.ColumnNumber + columnsShifted, firstCol),
                        dvAddress.LastAddress.RowNumber,
                        Math.Min(XLHelper.MaxColumnNumber, dvAddress.LastAddress.ColumnNumber + columnsShifted));
                }
                else
                    newRange = dvRange;

                if (newRange.RangeAddress.IsValid &&
                    newRange.RangeAddress.FirstAddress.ColumnNumber <=
                    newRange.RangeAddress.LastAddress.ColumnNumber)
                    dv.AddRange(newRange);
            }

            if (!dv.Ranges.Any())
                DataValidations.Delete(v => v == dv);
        }
    }

    internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
    {
        if (!range.IsEntireRow())
        {
            var model = new XLRangeAddress(
                range.RangeAddress.FirstAddress,
                new XLAddress(XLHelper.MaxRowNumber, range.RangeAddress.LastAddress.ColumnNumber, false, false));
            var rangesToSplit = Worksheet.MergedRanges
                .GetIntersectedRanges(model)
                .Where(r => r.RangeAddress.FirstAddress.ColumnNumber < range.RangeAddress.FirstAddress.ColumnNumber ||
                            r.RangeAddress.LastAddress.ColumnNumber > range.RangeAddress.LastAddress.ColumnNumber)
                .ToList();
            foreach (var rangeToSplit in rangesToSplit)
            {
                Worksheet.MergedRanges.Remove(rangeToSplit);
            }
        }

        Workbook.WorksheetsInternal.ForEach<XLWorksheet>(ws => MoveDefinedNamesRows(range, rowsShifted, ws.DefinedNames));
        MoveDefinedNamesRows(range, rowsShifted, Workbook.DefinedNamesInternal);
        ShiftConditionalFormattingRows(range, rowsShifted);
        ShiftDataValidationRows(range, rowsShifted);
        RemoveInvalidSparklines();
        ShiftPageBreaksRows(range, rowsShifted);

        ISheetListener hyperlinks = Hyperlinks;
        if (rowsShifted > 0)
        {
            var area = XLSheetRange
                .FromRangeAddress(range.RangeAddress)
                .ExtendBelow(rowsShifted - 1);
            Workbook.CalcEngine.OnInsertAreaAndShiftDown(range.Worksheet, area);
            hyperlinks.OnInsertAreaAndShiftDown(range.Worksheet, area);
        }
        else if (rowsShifted < 0)
        {
            var area = XLSheetRange.FromRangeAddress(range.RangeAddress);
            Workbook.CalcEngine.OnDeleteAreaAndShiftUp(range.Worksheet, area);
            hyperlinks.OnDeleteAreaAndShiftUp(range.Worksheet, area);
        }
    }

    private void ShiftPageBreaksRows(XLRange range, int rowsShifted)
    {
        for (var i = 0; i < PageSetup.RowBreaks.Count; i++)
        {
            var br = PageSetup.RowBreaks[i];
            if (range.RangeAddress.FirstAddress.RowNumber <= br)
            {
                PageSetup.RowBreaks[i] = br + rowsShifted;
            }
        }
    }

    private void ShiftConditionalFormattingRows(XLRange range, int rowsShifted)
    {
        if (!ConditionalFormats.Any()) return;
        var firstRow = range.RangeAddress.FirstAddress.RowNumber;
        if (firstRow == 1) return;

        var rowNum = rowsShifted > 0 ? firstRow - 1 : firstRow;
        var model = Row(rowNum).AsRange();

        foreach (var cf in ConditionalFormats.ToList())
        {
            var cfRanges = cf.Ranges.ToList();
            cf.Ranges.RemoveAll();

            foreach (var cfRange in cfRanges)
            {
                var cfAddress = cfRange.RangeAddress;
                IXLRange newRange;
                if (cfRange.Intersects(model))
                {
                    newRange = Range(cfAddress.FirstAddress.RowNumber,
                        cfAddress.FirstAddress.ColumnNumber,
                        Math.Min(XLHelper.MaxRowNumber, cfAddress.LastAddress.RowNumber + rowsShifted),
                        cfAddress.LastAddress.ColumnNumber);
                }
                else if (cfAddress.FirstAddress.RowNumber >= firstRow)
                {
                    newRange = Range(Math.Max(cfAddress.FirstAddress.RowNumber + rowsShifted, firstRow),
                        cfAddress.FirstAddress.ColumnNumber,
                        Math.Min(XLHelper.MaxRowNumber, cfAddress.LastAddress.RowNumber + rowsShifted),
                        cfAddress.LastAddress.ColumnNumber);
                }
                else
                    newRange = cfRange;

                if (newRange.RangeAddress.IsValid &&
                    newRange.RangeAddress.FirstAddress.RowNumber <= newRange.RangeAddress.LastAddress.RowNumber)
                    cf.Ranges.Add(newRange);
            }

            if (!cf.Ranges.Any())
                ConditionalFormats.Remove(f => f == cf);
        }
    }

    private void ShiftDataValidationRows(XLRange range, int rowsShifted)
    {
        if (!DataValidations.Any()) return;
        var firstRow = range.RangeAddress.FirstAddress.RowNumber;
        if (firstRow == 1) return;

        var rowNum = rowsShifted > 0 ? firstRow - 1 : firstRow;
        var model = Row(rowNum).AsRange();

        foreach (var dv in DataValidations.ToList())
        {
            var dvRanges = dv.Ranges.ToList();
            dv.ClearRanges();

            foreach (var dvRange in dvRanges)
            {
                var dvAddress = dvRange.RangeAddress;
                IXLRange newRange;
                if (dvRange.Intersects(model))
                {
                    newRange = Range(dvAddress.FirstAddress.RowNumber,
                        dvAddress.FirstAddress.ColumnNumber,
                        Math.Min(XLHelper.MaxRowNumber, dvAddress.LastAddress.RowNumber + rowsShifted),
                        dvAddress.LastAddress.ColumnNumber);
                }
                else if (dvAddress.FirstAddress.RowNumber >= firstRow)
                {
                    newRange = Range(Math.Max(dvAddress.FirstAddress.RowNumber + rowsShifted, firstRow),
                        dvAddress.FirstAddress.ColumnNumber,
                        Math.Min(XLHelper.MaxRowNumber, dvAddress.LastAddress.RowNumber + rowsShifted),
                        dvAddress.LastAddress.ColumnNumber);
                }
                else
                    newRange = dvRange;

                if (newRange.RangeAddress.IsValid &&
                    newRange.RangeAddress.FirstAddress.RowNumber <= newRange.RangeAddress.LastAddress.RowNumber)
                    dv.AddRange(newRange);
            }

            if (!dv.Ranges.Any())
                DataValidations.Delete(v => v == dv);
        }
    }

    private void RemoveInvalidSparklines()
    {
        var invalidSparklines = SparklineGroups.SelectMany(g => g)
            .Where(sl => !((XLAddress)sl.Location.Address).IsValid)
            .ToList();

        foreach (var sparkline in invalidSparklines)
        {
            Worksheet.SparklineGroups.Remove(sparkline.Location);
        }
    }

    private void MoveDefinedNamesRows(XLRange range, int rowsShifted, XLDefinedNames definedNames)
    {
        foreach (var definedName in definedNames)
        {
            if (definedName.SheetReferencesList.Any())
            {
                var newRangeList =
                    definedName.SheetReferencesList.Select(r => XLCell.ShiftFormulaRows(r, this, range, rowsShifted)).Where(
                        newReference => newReference.Length > 0).ToList();
                var unionFormula = string.Join(",", newRangeList);
                definedName.SetRefersTo(unionFormula);
            }
        }
    }

    private void MoveDefinedNamesColumns(XLRange range, int columnsShifted, XLDefinedNames definedNames)
    {
        foreach (var definedName in definedNames)
        {
            var newRangeList =
                definedName.SheetReferencesList.Select(r => XLCell.ShiftFormulaColumns(r, this, range, columnsShifted)).Where(
                    newReference => newReference.Length > 0).ToList();
            var unionFormula = string.Join(",", newRangeList);
            definedName.SetRefersTo(unionFormula);
        }
    }

    public void NotifyRangeShiftedRows(XLRange range, int rowsShifted)
    {
        var rangesToShift = _rangeRepository
            .Where(r => r.RangeAddress.IsValid)
            .OrderBy(r => r.RangeAddress.FirstAddress.RowNumber * -Math.Sign(rowsShifted))
            .ToList();

        WorksheetRangeShiftedRows(range, rowsShifted);

        var collapsed = false;
        foreach (var storedRange in rangesToShift)
        {
            if (storedRange.IsEntireColumn())
                continue;

            if (ReferenceEquals(range, storedRange))
                continue;

            storedRange.WorksheetRangeShiftedRows(range, rowsShifted);
            if (range.RangeAddress == storedRange.RangeAddress)
            {
                collapsed = true;
            }
        }
        if (!collapsed)
        {
            range.WorksheetRangeShiftedRows(range, rowsShifted);
        }
    }

    public void NotifyRangeShiftedColumns(XLRange range, int columnsShifted)
    {
        var rangesToShift = _rangeRepository
            .Where(r => r.RangeAddress.IsValid)
            .OrderBy(r => r.RangeAddress.FirstAddress.ColumnNumber * -Math.Sign(columnsShifted))
            .ToList();

        WorksheetRangeShiftedColumns(range, columnsShifted);

        var collapsed = false;
        foreach (var storedRange in rangesToShift)
        {
            if (storedRange.IsEntireRow())
                continue;

            if (ReferenceEquals(range, storedRange))
                continue;

            storedRange.WorksheetRangeShiftedColumns(range, columnsShifted);
            if (range.RangeAddress == storedRange.RangeAddress)
            {
                collapsed = true;
            }
        }
        if (!collapsed)
        {
            range.WorksheetRangeShiftedColumns(range, columnsShifted);
        }
    }

    public XLRow Row(int rowNumber, bool pingCells)
    {
        if (rowNumber <= 0 || rowNumber > XLHelper.MaxRowNumber)
            throw new ArgumentOutOfRangeException(nameof(rowNumber), $"Row number must be between 1 and {XLHelper.MaxRowNumber}");

        if (Internals.RowsCollection.TryGetValue(rowNumber, out var row))
            return row;
        if (pingCells)
        {
            // This is a new row so we're going to reference all
            // cells in columns of this row to preserve their formatting
            Internals.ColumnsCollection.Keys.ForEach(c => Cell(rowNumber, c).PingStyle());
        }

        row = RangeFactory.CreateRow(rowNumber);
        Internals.RowsCollection.Add(rowNumber, row);

        return row;
    }

    public IXLTable Table(XLRange range, bool addToTables, bool setAutofilter = true)
    {
        return Table(range, TableNameGenerator.GetNewTableName(Workbook), addToTables, setAutofilter);
    }

    public IXLTable Table(XLRange range, string name, bool addToTables, bool setAutofilter = true)
    {
        CheckRangeNotOverlappingOtherEntities(range);
        XLRangeAddress rangeAddress;
        if (range.Rows().Count() == 1)
        {
            rangeAddress = new XLRangeAddress(range.FirstCell().Address, range.LastCell().CellBelow().Address);
            range.InsertRowsBelow(1);
        }
        else
            rangeAddress = range.RangeAddress;

        var rangeKey = new XLRangeKey(XLRangeType.Table, rangeAddress);
        var table = (XLTable)_rangeRepository.GetOrCreate(ref rangeKey);

        if (table.Name != name)
            table.Name = name;

        if (addToTables && !Tables.Contains(table))
        {
            Tables.Add(table);
        }

        if (setAutofilter && !table.ShowAutoFilter)
            table.InitializeAutoFilter();

        return table;
    }

    private void CheckRangeNotOverlappingOtherEntities(XLRange range)
    {
        // Check that the range doesn't overlap with any existing tables
        var firstOverlappingTable = Tables.FirstOrDefault<XLTable>(t => t.RangeUsed()!.Intersects(range));
        if (firstOverlappingTable != null)
            throw new InvalidOperationException($"The range {range.RangeAddress.ToStringRelative(includeSheet: true)} is already part of table '{firstOverlappingTable.Name}'");

        // Check that the range doesn't overlap with any filters
        if (AutoFilter.IsEnabled && AutoFilter.Range.Intersects(range))
            throw new InvalidOperationException($"The range {range.RangeAddress.ToStringRelative(includeSheet: true)} overlaps with the worksheet's autofilter.");
    }

    private IXLRange GetRangeForSort()
    {
        var range = RangeUsed()!;
        SortColumns.ForEach(e => range.SortColumns.Add(e.ElementNumber, e.SortOrder, e.IgnoreBlanks, e.MatchCase));
        SortRows.ForEach(e => range.SortRows.Add(e.ElementNumber, e.SortOrder, e.IgnoreBlanks, e.MatchCase));
        return range;
    }

    public XLPivotTable PivotTable(string name)
    {
        return PivotTables.PivotTable(name);
    }

    public override IXLCells Cells()
    {
        return Cells(true, XLCellsUsedOptions.All);
    }

    public override XLCells Cells(bool usedCellsOnly)
    {
        if (usedCellsOnly)
            return Cells(true, XLCellsUsedOptions.AllContents);
        return Range((this as IXLRangeBase).FirstCellUsed(XLCellsUsedOptions.All)!,
                (this as IXLRangeBase).LastCellUsed(XLCellsUsedOptions.All)!)
            .Cells(false, XLCellsUsedOptions.All);
    }

    public override XLCell? Cell(string cellAddressInRange)
    {
        var cell = base.Cell(cellAddressInRange);
        if (cell is not null)
            return cell;

        if (Workbook.DefinedNames.TryGetValue(cellAddressInRange, out var definedName))
        {
            if (!definedName.Ranges.Any())
                return null;

            return definedName.Ranges.First().FirstCell().CastTo<XLCell>();
        }

        return null;
    }

    public override XLRange? Range(string rangeAddressStr)
    {
        if (XLHelper.IsValidRangeAddress(rangeAddressStr))
            return Range(new XLRangeAddress(Worksheet, rangeAddressStr));

        if (rangeAddressStr.Contains('['))
            return Table(rangeAddressStr[..rangeAddressStr.IndexOf("[", StringComparison.Ordinal)]) as XLRange;

        if (DefinedNames.TryGetValue(rangeAddressStr, out var sheetDefinedName))
            return sheetDefinedName.Ranges.First().CastTo<XLRange>();

        if (Workbook.DefinedNamesInternal.TryGetValue(rangeAddressStr, out var workbookDefinedName))
        {
            if (!workbookDefinedName.Ranges.Any())
                return null;

            return workbookDefinedName.Ranges.First().CastTo<XLRange>();
        }

        return null;
    }

    public IXLRanges MergedRanges => Internals.MergedRanges;

    IXLConditionalFormats IXLWorksheet.ConditionalFormats => ConditionalFormats;

    internal XLConditionalFormats ConditionalFormats { get; }

    public IXLSparklineGroups SparklineGroups => SparklineGroupsInternal;

    public IXLRanges SelectedRanges
    {
        get
        {
            _selectedRanges.RemoveAll(r => !r.RangeAddress.IsValid);
            return _selectedRanges;
        }
    }

    IXLCell? IXLWorksheet.ActiveCell
    {
        get => ActiveCell is not null ? new XLCell(this, ActiveCell.Value) : null;
        set => ActiveCell = value is not null ? XLSheetPoint.FromAddress(value.Address) : null;
    }

    /// <summary>
    /// Address of the active cell / cursor in the worksheet.
    /// </summary>
    internal XLSheetPoint? ActiveCell { get; set; }

    private XLCalcEngine CalcEngine => Workbook.CalcEngine;

    public XLCellValue Evaluate(string expression, string? formulaAddress = null)
    {
        IXLAddress? address = formulaAddress is not null ? XLAddress.Create(formulaAddress) : null;
        return CalcEngine.EvaluateFormula(expression, Workbook, this, address, true).ToCellValue();
    }

    public void RecalculateAllFormulas()
    {
        Internals.CellsCollection.FormulaSlice.MarkDirty(XLSheetRange.Full);
        Workbook.CalcEngine.Recalculate(Workbook, SheetId);
    }

    public string Author { get; set; }

    public override string ToString()
    {
        return Name;
    }

    public IXLPictures Pictures { get; private set; }

    public bool IsPasswordProtected => Protection.IsPasswordProtected;

    public bool IsProtected => Protection.IsProtected;

    public IXLPicture Picture(string pictureName)
    {
        return Pictures.Picture(pictureName);
    }

    public IXLPicture AddPicture(Stream stream)
    {
        return Pictures.Add(stream);
    }

    public IXLPicture AddPicture(Stream stream, string name)
    {
        return Pictures.Add(stream, name);
    }

    internal IXLPicture AddPicture(Stream stream, string name, int id)
    {
        return ((XLPictures)Pictures).Add(stream, name, id);
    }

    public IXLPicture AddPicture(Stream stream, XLPictureFormat format)
    {
        return Pictures.Add(stream, format);
    }

    public IXLPicture AddPicture(Stream stream, XLPictureFormat format, string name)
    {
        return Pictures.Add(stream, format, name);
    }

    public IXLPicture AddPicture(string imageFile)
    {
        return Pictures.Add(imageFile);
    }

    public IXLPicture AddPicture(string imageFile, string name)
    {
        return Pictures.Add(imageFile, name);
    }

    public override bool IsEntireRow()
    {
        return true;
    }

    public override bool IsEntireColumn()
    {
        return true;
    }

    internal IXLTable InsertTable(XLSheetPoint origin, IInsertDataReader reader, string? tableName, bool createTable, bool addHeadings, bool transpose)
    {
        if (createTable && Tables.Any<XLTable>(t => t.Area.Contains(origin)))
            throw new InvalidOperationException($"This cell '{origin}' is already part of a table.");

        var range = InsertData(origin, reader, addHeadings, transpose);

        if (createTable)
            // Create a table and save it in the file
            return tableName == null ? range.CreateTable() : range.CreateTable(tableName);
        // Create a table but keep it in memory. Saved file will contain only "raw" data and column headers
        return tableName == null ? range.AsTable() : range.AsTable(tableName);
    }

    internal XLRange InsertData(XLSheetPoint origin, IInsertDataReader reader, bool addHeadings, bool transpose)
    {
        // Prepare data. Heading is basically just another row of data, so unify it.
        var rows = reader.GetRecords();
        var propCount = reader.GetPropertiesCount();
        if (addHeadings)
        {
            var headings = new XLCellValue[propCount];
            for (var i = 0; i < propCount; i++)
                headings[i] = reader.GetPropertyName(i);

            rows = new[] { headings }.Concat(rows);
        }

        if (transpose)
        {
            rows = TransposeJaggedArray(rows);
        }

        var valueSlice = Internals.CellsCollection.ValueSlice;
        var styleSlice = Internals.CellsCollection.StyleSlice;

        // A buffer to avoid multiple enumerations of the source.
        var rowBuffer = new List<XLCellValue>();
        var maximumColumn = origin.Column;
        var rowNumber = origin.Row;
        foreach (var row in rows)
        {
            rowBuffer.AddRange(row);

            // InsertData should also clear data and if row doesn't have enough data,
            // fill in the rest. Only fill up to the props to be consistent. We can't
            // know how long any next row will be, so props are used as a source of truth
            // for which columns should be cleared.
            for (var i = rowBuffer.Count; i < propCount; ++i)
                rowBuffer.Add(Blank.Value);

            // Each row can have different number of values, so we have to check every row.
            maximumColumn = Math.Max(origin.Column + rowBuffer.Count - 1, maximumColumn);
            if (maximumColumn > XLHelper.MaxColumnNumber || rowNumber > XLHelper.MaxRowNumber)
                throw new ArgumentException("Data would write out of the sheet.");

            var column = origin.Column;
            foreach (var t in rowBuffer)
            {
                var value = t;
                var point = new XLSheetPoint(rowNumber, column);
                var modifiedStyle = GetStyleForValue(value, point);
                if (modifiedStyle is not null)
                {
                    if (value.IsText && value.GetText()[0] == '\'')
                        value = value.GetText().Substring(1);

                    styleSlice.Set(point, modifiedStyle);
                }

                valueSlice.SetCellValue(point, value);
                column++;
            }

            rowBuffer.Clear();
            rowNumber++;
        }

        // If there is no row, rowNumber is kept at origin instead of last row + 1 .
        var lastRow = Math.Max(rowNumber - 1, origin.Row);
        var insertedArea = new XLSheetRange(origin, new XLSheetPoint(lastRow, maximumColumn));

        // If inserted area affected a table, we must fix headings and totals, because these values
        // are duplicated. Basically the table values are the truth and cells are a reflection of the
        // truth, but here we inserted shadow first.
        foreach (var table in Tables)
            table.RefreshFieldsFromCells(insertedArea);

        // Invalidate only once, not for every cell.
        Workbook.CalcEngine.MarkDirty(Worksheet, insertedArea);

        // Return area that contains all inserted cells, no matter how jagged were data.
        return Range(
            insertedArea.FirstPoint.Row,
            insertedArea.FirstPoint.Column,
            insertedArea.LastPoint.Row,
            insertedArea.LastPoint.Column);

        // Rather memory inefficient, but the original code also materialized
        // data through Linq/required multiple enumerations.
        static List<List<XLCellValue>> TransposeJaggedArray(IEnumerable<IEnumerable<XLCellValue>> enumerable)
        {
            var destination = new List<List<XLCellValue>>();

            var sourceRow = 1;
            foreach (var row in enumerable)
            {
                var sourceColumn = 1;
                foreach (var sourceValue in row)
                {
                    // The original has `sourceValue` at [sourceRow, sourceColumn]
                    var destinationRowCount = destination.Count;
                    if (sourceColumn > destinationRowCount)
                        destination.Add([]);

                    // There can be jagged arrays and the destination can have spaces between columns.
                    var destinationRow = destination[sourceColumn - 1];
                    while (destinationRow.Count < sourceRow - 1)
                        destinationRow.Add(Blank.Value);

                    destinationRow.Add(sourceValue);
                    sourceColumn++;
                }

                sourceRow++;
            }

            return destination;
        }
    }

    /// <summary>
    /// Get cell or null, if cell doesn't exist.
    /// </summary>
    internal XLCell? GetCell(XLSheetPoint point)
    {
        return Worksheet.Internals.CellsCollection.GetUsedCell(point);
    }

    public XLRange GetOrCreateRange(XLRangeParameters xlRangeParameters)
    {
        var rangeKey = new XLRangeKey(XLRangeType.Range, xlRangeParameters.RangeAddress);
        var range = _rangeRepository.GetOrCreate(ref rangeKey);
        if (xlRangeParameters.DefaultStyle != null && range.StyleValue == StyleValue)
            range.InnerStyle = xlRangeParameters.DefaultStyle;

        return (XLRange)range;
    }

    /// <summary>
    /// Get a range row from the shared repository or create a new one.
    /// </summary>
    /// <param name="address">Address of range row.</param>
    /// <param name="defaultStyle">Style to apply. If null the worksheet's style is applied.</param>
    /// <returns>Range row with the specified address.</returns>
    public XLRangeRow RangeRow(XLRangeAddress address, IXLStyle? defaultStyle = null)
    {
        var rangeKey = new XLRangeKey(XLRangeType.RangeRow, address);
        var rangeRow = (XLRangeRow)_rangeRepository.GetOrCreate(ref rangeKey);

        if (defaultStyle != null && rangeRow.StyleValue == StyleValue)
            rangeRow.InnerStyle = defaultStyle;

        return rangeRow;
    }

    /// <summary>
    /// Get a range column from the shared repository or create a new one.
    /// </summary>
    /// <param name="address">Address of range column.</param>
    /// <param name="defaultStyle">Style to apply. If null the worksheet's style is applied.</param>
    /// <returns>Range column with the specified address.</returns>
    public XLRangeColumn RangeColumn(XLRangeAddress address, IXLStyle? defaultStyle = null)
    {
        var rangeKey = new XLRangeKey(XLRangeType.RangeColumn, address);
        var rangeColumn = (XLRangeColumn)_rangeRepository.GetOrCreate(ref rangeKey);

        if (defaultStyle != null && rangeColumn.StyleValue == StyleValue)
            rangeColumn.InnerStyle = defaultStyle;

        return rangeColumn;
    }

    protected override void OnRangeAddressChanged(XLRangeAddress oldAddress, XLRangeAddress newAddress)
    {
    }

    public void RelocateRange(XLRangeType rangeType, XLRangeAddress oldAddress, XLRangeAddress newAddress)
    {
        if (_rangeRepository == null)
            return;

        var oldKey = new XLRangeKey(rangeType, oldAddress);
        var newKey = new XLRangeKey(rangeType, newAddress);
        var range = _rangeRepository.Replace(ref oldKey, ref newKey);

        foreach (var rangeIndex in _rangeIndices)
        {
            if (!rangeIndex.MatchesType(rangeType))
                continue;

            if (rangeIndex.Remove(oldAddress) &&
                newAddress.IsValid &&
                range != null)
            {
                rangeIndex.Add(range);
            }
        }
    }

    internal void DeleteColumn(int columnNumber)
    {
        Internals.ColumnsCollection.Remove(columnNumber);

        var columnsToMove = new List<int>(Internals.ColumnsCollection.Where(c => c.Key > columnNumber).Select(c => c.Key).OrderBy(c => c));
        foreach (var column in columnsToMove)
        {
            Internals.ColumnsCollection.Add(column - 1, Internals.ColumnsCollection[column]);
            Internals.ColumnsCollection.Remove(column);

            Internals.ColumnsCollection[column - 1].SetColumnNumber(column - 1);
        }
    }

    internal void DeleteRow(int rowNumber)
    {
        Internals.RowsCollection.Remove(rowNumber);

        var rowsToMove = new List<int>(Internals.RowsCollection.Where(c => c.Key > rowNumber).Select(c => c.Key).OrderBy(r => r));
        foreach (var row in rowsToMove)
        {
            Internals.RowsCollection.Add(row - 1, Worksheet.Internals.RowsCollection[row]);
            Internals.RowsCollection.Remove(row);

            Internals.RowsCollection[row - 1].SetRowNumber(row - 1);
        }
    }

    internal void DeleteRange(XLRangeAddress rangeAddress)
    {
        var rangeKey = new XLRangeKey(XLRangeType.Range, rangeAddress);
        _rangeRepository.Remove(ref rangeKey);
    }

    /// <summary>
    /// Get the actual style for a point in the sheet.
    /// </summary>
    internal XLStyleValue GetStyleValue(XLSheetPoint point)
    {
        var styleValue = Internals.CellsCollection.StyleSlice[point];
        if (styleValue is not null)
            return styleValue;

        // If the slice doesn't contain any value, determine values by inheriting.
        // Cells that lie on an intersection of a XLColumn and a XLRow have their
        // style set when column/row is created to avoid problems with correct which
        // style has precedence. I.e. set column blue, set row red => cell is red.
        // Swap order the the cell is blue.
        var sheetStyle = StyleValue;
        var rowStyle = Internals.RowsCollection.TryGetValue(point.Row, out var row)
            ? row.StyleValue
            : sheetStyle;
        var colStyle = Internals.ColumnsCollection.TryGetValue(point.Column, out var column)
            ? column.StyleValue
            : sheetStyle;

        return XLStyleValue.Combine(sheetStyle, rowStyle, colStyle);
    }

    /// <summary>
    /// Get a style that should be used for a <paramref name="value"/>,
    /// if the value is set to the <paramref name="point"/>.
    /// </summary>
    internal XLStyleValue? GetStyleForValue(XLCellValue value, XLSheetPoint point)
    {
        // Because StyleValue property retrieves value from a slice,
        // access it only if necessary. This happens during every cell
        // of modification and thus is performance-critical.
        switch (value.Type)
        {
            case XLDataType.DateTime:
                {
                    var onlyDatePart = value.GetUnifiedNumber() % 1 == 0;
                    var styleValue = GetStyleValue(point);
                    if (styleValue.NumberFormat.Format.Length == 0 &&
                        styleValue.NumberFormat.NumberFormatId == 0)
                    {
                        var dateTimeNumberFormat = styleValue.NumberFormat.WithNumberFormatId(onlyDatePart ? 14 : 22);
                        return styleValue.WithNumberFormat(dateTimeNumberFormat);
                    }
                }
                break;

            case XLDataType.TimeSpan:
                {
                    var styleValue = GetStyleValue(point);
                    if (styleValue.NumberFormat.Format.Length == 0 && styleValue.NumberFormat.NumberFormatId == 0)
                    {
                        var durationNumberFormat = styleValue.NumberFormat.WithNumberFormatId(46);
                        return styleValue.WithNumberFormat(durationNumberFormat);
                    }
                }
                break;

            case XLDataType.Text:
                {
                    var text = value.GetText();
                    XLStyleValue? styleValue = null;
                    if (text.Length > 0 && text[0] == '\'')
                    {
                        styleValue = GetStyleValue(point);
                        styleValue = styleValue.WithIncludeQuotePrefix(true);
                    }

                    var containsNewLine = text.AsSpan()
                        .Contains(Environment.NewLine.AsSpan(), StringComparison.Ordinal);
                    if (containsNewLine)
                    {
                        styleValue ??= GetStyleValue(point);
                        if (!styleValue.Alignment.WrapText)
                        {
                            styleValue = styleValue.WithAlignment(static alignment => alignment.WithWrapText(true));
                        }
                    }

                    return styleValue;
                }
        }

        return null;
    }
}
