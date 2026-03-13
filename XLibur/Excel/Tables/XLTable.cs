
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Dynamic;
using System.Globalization;
using System.Linq;
using System.Text;
using XLibur.Extensions;

namespace XLibur.Excel;

[DebuggerDisplay("{Name}")]
internal sealed class XLTable : XLRange, IXLTable
{
    internal bool _showTotalsRow;
    internal HashSet<string> _uniqueNames = null!;

    /// <summary>
    /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
    /// </summary>
    public XLTable(XLRangeParameters xlRangeParameters)
        : base(xlRangeParameters)
    {
        InitializeValues(false);
    }
    public override XLRangeType RangeType
    {
        get { return XLRangeType.Table; }
    }

    private IXLRangeAddress? _lastRangeAddress;
    private Dictionary<string, IXLTableField>? _fieldNames;

    public Dictionary<string, IXLTableField> FieldNames
    {
        get
        {
            if (_fieldNames != null && _lastRangeAddress != null && _lastRangeAddress.Equals(RangeAddress))
                return _fieldNames;

            _lastRangeAddress = RangeAddress;

            RescanFieldNames();

            return _fieldNames!;
        }
    }

    /// <summary>
    /// Area of the range, including headings and totals, if table has them.
    /// </summary>
    internal XLSheetRange Area => XLSheetRange.FromRangeAddress(RangeAddress);

    private void RescanFieldNames()
    {
        if (ShowHeaderRow)
        {
            var oldFieldNames = _fieldNames ?? CreateFieldNames();
            _fieldNames = CreateFieldNames();
            var headersRow = HeadersRow(false)!;
            var cellPos = 0;
            foreach (XLCell cell in headersRow.Cells())
            {
                var cellValue = cell.CachedValue;
                var name = cellValue.ToString(CultureInfo.CurrentCulture);

                if (oldFieldNames.TryGetValue(name, out var tableField))// && tableField.Column.ColumnNumber() == cell.Address.ColumnNumber)
                {
                    ((XLTableField)tableField).Index = cellPos;
                    _fieldNames.Add(name, tableField);
                    cellPos++;
                    continue;
                }

                // Be careful here. Fields names may actually be whitespace, but not empty
                if (string.IsNullOrEmpty(name))
                {
                    name = GetUniqueName("Column", cellPos + 1, true);
                }
                if (_fieldNames.ContainsKey(name))
                    throw new ArgumentException("The header row contains more than one field name '" + name + "'.");

                _fieldNames.Add(name, new XLTableField(this, name) { Index = cellPos++ });

                // Field names are the source of the truth that is projected
                // to the cells and field names can be only text. Fix the cell,
                // so cell fulfills its job of being dependent on the field name.
                if (!cellValue.Equals(name))
                {
                    cell.SetValue(name, false, false);
                }
            }
        }
        else
        {
            _fieldNames ??= CreateFieldNames();
            var colCount = ColumnCount();
            for (var i = 1; i <= colCount; i++)
            {
                if (_fieldNames.Values.All(f => f.Index != i - 1))
                {
                    var name = "Column" + i;

                    _fieldNames.Add(name, new XLTableField(this, name) { Index = i - 1 });
                }
            }
        }
    }

    internal void AddFields(IEnumerable<string> fieldNames)
    {
        _fieldNames = CreateFieldNames();

        var cellPos = 0;
        foreach (var name in fieldNames)
        {
            _fieldNames.Add(name, new XLTableField(this, name) { Index = cellPos++ });
        }
    }

    internal void RenameField(string oldName, string newName)
    {
        if (!_fieldNames!.Remove(oldName, out var field))
            throw new ArgumentException("The field does not exist in this table", nameof(oldName));

        _fieldNames.Add(newName, field);
    }

    internal string? RelId { get; set; }

    public IXLTableRange? DataRange
    {
        get
        {
            var firstDataRowNumber = 1;
            var lastDataRowNumber = RowCount();

            if (_showHeaderRow)
                firstDataRowNumber++;

            if (_showTotalsRow)
                lastDataRowNumber--;

            if (firstDataRowNumber > lastDataRowNumber)
                return null;

            var range = Range(firstDataRowNumber, 1, lastDataRowNumber, ColumnCount());

            return new XLTableRange(range, this);
        }
    }

    public XLAutoFilter AutoFilter
    {
        get
        {
            field ??= new XLAutoFilter { IsTableAutoFilter = true };

            field.Range = ShowTotalsRow ? Range(1, 1, RowCount() - 1, ColumnCount()) : AsRange();
            return field;
        }
    }

    public override IXLAutoFilter SetAutoFilter()
    {
        return AutoFilter;
    }

    protected override void OnRangeAddressChanged(XLRangeAddress oldAddress, XLRangeAddress newAddress)
    {
        //Do nothing for table
    }

    #region IXLTable Members

    public bool EmphasizeFirstColumn { get; set; }
    public bool EmphasizeLastColumn { get; set; }
    public bool ShowRowStripes { get; set; }
    public bool ShowColumnStripes { get; set; }

    public bool ShowAutoFilter
    {
        get => _showHeaderRow && field;
        set;
    }

    public XLTableTheme Theme { get; set; } = null!;

    public string Name
    {
        get;
        set
        {
            if (field == value) return;

            // Allow casing-only changes (e.g. "Table1" -> "TABLE1") without conflict check
            var oldname = field ?? string.Empty;
            var casingOnlyChange = !string.IsNullOrWhiteSpace(oldname) &&
                                   string.Equals(oldname, value, StringComparison.OrdinalIgnoreCase);

            if (!casingOnlyChange && !TableNameValidator.IsValidTableName(value, Worksheet, out var message))
                throw new ArgumentException(message);

            field = value;

            // Some totals row formula depends on the table name. Update them.
            if (_fieldNames?.Any() ?? false)
                Fields.ForEach(f => ((XLTableField)f).UpdateTableFieldTotalsRowFormula());

            if (!string.IsNullOrWhiteSpace(oldname) && !string.Equals(oldname, field, StringComparison.OrdinalIgnoreCase))
            {
                Worksheet.Tables.Add(this);
                if (Worksheet.Tables.Contains(oldname))
                    Worksheet.Tables.Remove(oldname);
            }
        }
    } = null!;

    public bool ShowTotalsRow
    {
        get => _showTotalsRow;
        set
        {
            if (value && !_showTotalsRow)
                InsertRowsBelow(1);
            else if (!value && _showTotalsRow)
                TotalsRow()!.Delete();

            _showTotalsRow = value;

            // Invalidate fields' columns
            Fields.Cast<XLTableField>().ForEach(f => f.Column = null!);

            if (_showTotalsRow)
            {
                AutoFilter.Range = Worksheet.Range(
                    RangeAddress.FirstAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber - 1, RangeAddress.LastAddress.ColumnNumber);
            }
            else
                AutoFilter.Range = Worksheet.Range(RangeAddress);
        }
    }

    public IXLRangeRow? HeadersRow()
    {
        return HeadersRow(true);
    }

    internal XLRangeRow? HeadersRow(bool scanForNewFieldsNames)
    {
        if (!ShowHeaderRow) return null;

        if (scanForNewFieldsNames)
        {
            var tempResult = FieldNames;
        }

        return FirstRow();
    }

    public IXLRangeRow? TotalsRow()
    {
        return ShowTotalsRow ? LastRow() : null;
    }

    public IXLTableField Field(string fieldName)
    {
        return Field(GetFieldIndex(fieldName));
    }

    public IXLTableField Field(int fieldIndex)
    {
        return FieldNames.Values.First(f => f.Index == fieldIndex);
    }

    public IEnumerable<IXLTableField> Fields
    {
        get
        {
            var columnCount = ColumnCount();
            for (var co = 0; co < columnCount; co++)
                yield return Field(co);
        }
    }

    public IXLTable Resize(IXLRangeAddress rangeAddress)
    {
        return Resize(Worksheet.Range(RangeAddress));
    }

    public IXLTable Resize(string rangeAddress)
    {
        return Resize(Worksheet.Range(rangeAddress)!);
    }

    public IXLTable Resize(IXLCell firstCell, IXLCell lastCell)
    {
        return Resize(Worksheet.Range(firstCell, lastCell));
    }

    public IXLTable Resize(string firstCellAddress, string lastCellAddress)
    {
        return Resize(Worksheet.Range(firstCellAddress, lastCellAddress));
    }

    public IXLTable Resize(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
    {
        return Resize(Worksheet.Range(firstCellAddress, lastCellAddress));
    }

    public IXLTable Resize(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
    {
        return Resize(Worksheet.Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn));
    }

    public IXLTable Resize(IXLRange range)
    {
        if (!ShowHeaderRow)
            throw new NotImplementedException("Resizing of tables with no headers not supported yet.");

        if (Worksheet != range.Worksheet)
            throw new InvalidOperationException("You cannot resize a table to a range on a different sheet.");

        var totalsRowChanged = ShowTotalsRow ? range.LastRow()!.RowNumber() - TotalsRow()!.RowNumber() : 0;
        var oldTotalsRowNumber = ShowTotalsRow ? TotalsRow()!.RowNumber() : -1;

        var existingHeaders = FieldNames.Keys;
        var newHeaders = new HashSet<string>();

        // Force evaluation of f.Column field
        var tempArray = Fields.Select(f => f.Column).ToArray();

        var firstRow = range.Row(1);
        if (!firstRow.FirstCell().Address.Equals(HeadersRow()!.FirstCell().Address)
            || !firstRow.LastCell().Address.Equals(HeadersRow()!.LastCell().Address))
        {
            _uniqueNames.Clear();
            var co = 1;
            foreach (var c in firstRow.Cells())
            {
                if (c.IsEmpty(XLCellsUsedOptions.Contents))
                    c.Value = GetUniqueName("Column", co, true);

                var header = c.GetString();
                _uniqueNames.Add(header);

                if (!existingHeaders.Contains(header))
                    newHeaders.Add(header);

                co++;
            }
        }

        if (totalsRowChanged < 0)
        {
            range.Rows(r => r.RowNumber().Equals(TotalsRow()!.RowNumber() + totalsRowChanged)).Single().InsertRowsAbove(1);
            range = Worksheet.Range(range.FirstCell(), range.LastCell().CellAbove());
            oldTotalsRowNumber++;
        }
        else if (totalsRowChanged > 0)
        {
            TotalsRow()!.RowBelow(totalsRowChanged + 1).InsertRowsAbove(1);
            TotalsRow()!.AsRange().Delete(XLShiftDeletedCells.ShiftCellsUp);
        }

        RangeAddress = (XLRangeAddress)range.RangeAddress;
        RescanFieldNames();

        if (ShowTotalsRow)
        {
            foreach (var f in _fieldNames!.Values)
            {
                var fieldColumn = f.Index + 1;
                var c = TotalsRow()!.Cell(fieldColumn);
                if (!c.IsEmpty() && newHeaders.Contains(f.Name))
                {
                    f.TotalsRowLabel = c.GetFormattedString();
                }
            }

            if (totalsRowChanged != 0)
            {
                foreach (var f in _fieldNames!.Values.Cast<XLTableField>())
                {
                    f.UpdateTableFieldTotalsRowFormula();
                    var fieldColumn = f.Index + 1;
                    var c = TotalsRow()!.Cell(fieldColumn);
                    if (!string.IsNullOrWhiteSpace(f.TotalsRowLabel))
                    {
                        //Remove previous row's label
                        var oldTotalsCell = Worksheet.Cell(oldTotalsRowNumber, f.Column.ColumnNumber());
                        if (oldTotalsCell.Value.Equals(f.TotalsRowLabel))
                            oldTotalsCell.Value = Blank.Value;
                    }

                    if (!string.IsNullOrEmpty(f.TotalsRowLabel))
                        c.SetValue(f.TotalsRowLabel);
                }
            }
        }

        return this;
    }

    public IXLTable SetEmphasizeFirstColumn()
    {
        EmphasizeFirstColumn = true;
        return this;
    }

    public IXLTable SetEmphasizeFirstColumn(bool value)
    {
        EmphasizeFirstColumn = value;
        return this;
    }

    public IXLTable SetEmphasizeLastColumn()
    {
        EmphasizeLastColumn = true;
        return this;
    }

    public IXLTable SetEmphasizeLastColumn(bool value)
    {
        EmphasizeLastColumn = value;
        return this;
    }

    public IXLTable SetShowRowStripes()
    {
        ShowRowStripes = true;
        return this;
    }

    public IXLTable SetShowRowStripes(bool value)
    {
        ShowRowStripes = value;
        return this;
    }

    public IXLTable SetShowColumnStripes()
    {
        ShowColumnStripes = true;
        return this;
    }

    public IXLTable SetShowColumnStripes(bool value)
    {
        ShowColumnStripes = value;
        return this;
    }

    public IXLTable SetShowTotalsRow()
    {
        ShowTotalsRow = true;
        return this;
    }

    public IXLTable SetShowTotalsRow(bool value)
    {
        ShowTotalsRow = value;
        return this;
    }

    public IXLTable SetShowAutoFilter()
    {
        ShowAutoFilter = true;
        return this;
    }

    public IXLTable SetShowAutoFilter(bool value)
    {
        ShowAutoFilter = value;
        return this;
    }

    public new IXLRange Sort(string columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false, bool ignoreBlanks = true)
    {
        var toSortBy = new StringBuilder();
        foreach (var coPairTrimmed in columnsToSortBy.Split(',').Select(coPair => coPair.Trim()))
        {
            string coString;
            string order;
            if (coPairTrimmed.Contains(' '))
            {
                var pair = coPairTrimmed.Split(' ');
                coString = pair[0];
                order = pair[1];
            }
            else
            {
                coString = coPairTrimmed;
                order = sortOrder == XLSortOrder.Ascending ? "ASC" : "DESC";
            }

            if (!int.TryParse(coString, out var co))
                co = Field(coString).Index + 1;

            if (toSortBy.Length > 0)
                toSortBy.Append(',');

            toSortBy.Append(co);
            toSortBy.Append(' ');
            toSortBy.Append(order);
        }
        return DataRange!.Sort(toSortBy.ToString(), sortOrder, matchCase, ignoreBlanks);
    }

    public new IXLTable Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    IXLAutoFilter IXLTable.AutoFilter
    {
        get { return AutoFilter; }
    }

    #endregion IXLTable Members

    private void InitializeValues(bool setAutofilter)
    {
        ShowRowStripes = true;
        _showHeaderRow = true;
        Theme = XLTableTheme.TableStyleMedium2;
        if (setAutofilter)
            InitializeAutoFilter();

        if (RowCount() == 1)
            InsertRowsBelow(1);
    }

    public void InitializeAutoFilter()
    {
        ShowAutoFilter = true;
    }

    internal void OnAddedToTables()
    {
        _uniqueNames = [];
        var co = 1;
        foreach (var c in Row(1).Cells())
        {
            // Be careful here. Fields names may actually be whitespace, but not empty
            if (c.IsEmpty(XLCellsUsedOptions.Contents))
                ((XLCell)c).SetValue(GetUniqueName("Column", co, true), false, false);
            _uniqueNames.Add(c.GetString());
            co++;
        }
    }

    private static Dictionary<string, IXLTableField> CreateFieldNames()
    {
        return new Dictionary<string, IXLTableField>(StringComparer.CurrentCultureIgnoreCase);
    }

    private string GetUniqueName(string originalName, int initialOffset, bool enforceOffset)
    {
        var name = string.Concat(originalName, enforceOffset ? initialOffset.ToInvariantString() : string.Empty);
        if (_uniqueNames?.Contains(name) ?? false)
        {
            var i = initialOffset;
            name = originalName + i.ToInvariantString();
            while (_uniqueNames.Contains(name))
            {
                i++;
                name = originalName + i.ToInvariantString();
            }
        }

        return name;
    }

    public int GetFieldIndex(string name)
    {
        // There is a discrepancy in the way headers with line breaks are stored.
        // The entry in the table definition will contain \r\n
        // but the shared string value of the actual cell will contain only \n
        name = name.Replace("\r\n", "\n");
        if (FieldNames.TryGetValue(name, out var tableField))
            return tableField.Index;

        throw new ArgumentOutOfRangeException("The header row doesn't contain field name '" + name + "'.");
    }

    internal bool _showHeaderRow;

    public bool ShowHeaderRow
    {
        get => _showHeaderRow;
        set
        {
            if (_showHeaderRow == value) return;

            if (_showHeaderRow)
            {
                var headersRow = HeadersRow();
                _uniqueNames = new HashSet<string>();
                var co = 1;
                foreach (var c in headersRow!.Cells())
                {
                    if (string.IsNullOrWhiteSpace(c.GetString()))
                        c.Value = GetUniqueName("Column", co, true);
                    _uniqueNames.Add(c.GetString());
                    co++;
                }

                headersRow.Clear();
                RangeAddress = new XLRangeAddress(
                    new XLAddress(Worksheet, RangeAddress.FirstAddress.RowNumber + 1,
                        RangeAddress.FirstAddress.ColumnNumber,
                        RangeAddress.FirstAddress.FixedRow,
                        RangeAddress.FirstAddress.FixedColumn),
                    RangeAddress.LastAddress);
            }
            else
            {
                var asRange = Worksheet.Range(
                    RangeAddress.FirstAddress.RowNumber - 1,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber);
                var firstRow = asRange.FirstRow()!;
                IXLRangeRow rangeRow;
                if (firstRow.IsEmpty(XLCellsUsedOptions.All))
                {
                    rangeRow = firstRow;
                    RangeAddress = new XLRangeAddress(
                        new XLAddress(Worksheet,
                            RangeAddress.FirstAddress.RowNumber - 1,
                            RangeAddress.FirstAddress.ColumnNumber,
                            RangeAddress.FirstAddress.FixedRow,
                            RangeAddress.FirstAddress.FixedColumn),
                        RangeAddress.LastAddress);
                }
                else
                {
                    var fAddress = RangeAddress.FirstAddress;
                    //var lAddress = RangeAddress.LastAddress;

                    rangeRow = firstRow.InsertRowsBelow(1, false).First();

                    RangeAddress = new XLRangeAddress(
                        fAddress,
                        RangeAddress.LastAddress);
                }

                var co = 1;
                foreach (var name in FieldNames.Values.Select(f => f.Name))
                {
                    rangeRow.Cell(co).SetValue(name);
                    co++;
                }
            }

            _showHeaderRow = value;

            // Invalidate fields' columns
            Fields.Cast<XLTableField>().ForEach(f => f.Column = null!);
        }
    }

    public IXLTable SetShowHeaderRow()
    {
        return SetShowHeaderRow(true);
    }

    public IXLTable SetShowHeaderRow(bool value)
    {
        ShowHeaderRow = value;
        return this;
    }

    public void ExpandTableRows(int rows)
    {
        RangeAddress = new XLRangeAddress(
            RangeAddress.FirstAddress,
            new XLAddress(Worksheet, RangeAddress.LastAddress.RowNumber + rows,
                RangeAddress.LastAddress.ColumnNumber,
                RangeAddress.LastAddress.FixedRow,
                RangeAddress.LastAddress.FixedColumn));
    }

    public override XLRangeColumn Column(int columnNumber)
    {
        var column = base.Column(columnNumber);
        column.Table = this;
        return column;
    }

    public override XLRangeColumn Column(string columnName)
    {
        var column = base.Column(columnName);
        column.Table = this;
        return column;
    }

    public override IXLRangeColumns Columns(int firstColumn, int lastColumn)
    {
        var columns = base.Columns(firstColumn, lastColumn);
        columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    public override IXLRangeColumns Columns(Func<IXLRangeColumn, bool>? predicate = null)
    {
        var columns = base.Columns(predicate);
        columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    public override IXLRangeColumns Columns(string columns)
    {
        var cols = base.Columns(columns);
        cols.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
        return cols;
    }

    public override IXLRangeColumns Columns(string firstColumn, string lastColumn)
    {
        var columns = base.Columns(firstColumn, lastColumn);
        columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    internal override XLRangeColumns ColumnsUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, bool>? predicate = null)
    {
        var columns = base.ColumnsUsed(options, predicate);
        columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    internal override XLRangeColumns ColumnsUsed(Func<IXLRangeColumn, bool>? predicate = null)
    {
        var columns = base.ColumnsUsed(predicate);
        columns.Cast<XLRangeColumn>().ForEach(column => column.Table = this);
        return columns;
    }

    IXLPivotTable IXLRangeBase.CreatePivotTable(IXLCell targetCell, string name)
    {
        return CreatePivotTable(targetCell, name);
    }

    internal new XLPivotTable CreatePivotTable(IXLCell targetCell, string name)
    {
        return (XLPivotTable)targetCell.Worksheet.PivotTables.Add(name, targetCell, this);
    }

    public IEnumerable<dynamic> AsDynamicEnumerable()
    {
        foreach (var row in DataRange!.Rows())
        {
            dynamic expando = new ExpandoObject();
            foreach (var f in Fields)
            {
                var value = row.Cell(f.Index + 1).Value;
                // ExpandoObject supports IDictionary so we can extend it like this
                var expandoDict = (IDictionary<string, object>)expando;
                expandoDict[f.Name] = value;
            }

            yield return expando;
        }
    }

    public DataTable AsNativeDataTable()
    {
        var table = new DataTable(Name);

        foreach (var f in Fields.Cast<XLTableField>())
        {
            var type = typeof(object);
            if (f.IsConsistentDataType())
            {
                var c = f.Column.Cells().Skip(ShowHeaderRow ? 1 : 0).First();
                switch (c.DataType)
                {
                    case XLDataType.Text:
                        type = typeof(string);
                        break;

                    case XLDataType.Boolean:
                        type = typeof(bool);
                        break;

                    case XLDataType.DateTime:
                        type = typeof(DateTime);
                        break;

                    case XLDataType.TimeSpan:
                        type = typeof(TimeSpan);
                        break;

                    case XLDataType.Number:
                        type = typeof(double);
                        break;
                }
            }

            table.Columns.Add(f.Name, type);
        }

        foreach (var row in DataRange!.Rows())
        {
            var dr = table.NewRow();

            foreach (var f in Fields)
            {
                dr[f.Name] = row.Cell(f.Index + 1).Value.ToObject();
            }

            table.Rows.Add(dr);
        }

        return table;
    }

    public IXLTable CopyTo(IXLWorksheet targetSheet)
    {
        return CopyTo((XLWorksheet)targetSheet);
    }

    internal IXLTable CopyTo(XLWorksheet targetSheet, bool copyData = true)
    {
        if (targetSheet == Worksheet)
            throw new InvalidOperationException("Cannot copy table to the worksheet it already belongs to.");

        var targetRange = targetSheet.Range(RangeAddress.WithoutWorksheet());
        if (copyData)
            RangeUsed()!.CopyTo(targetRange);
        else
            HeadersRow()!.CopyTo(targetRange.FirstRow()!);

        var tableName = Name;
        var newTable = (XLTable)targetSheet.Table(targetRange, tableName, true);

        newTable.RelId = null;
        newTable.EmphasizeFirstColumn = EmphasizeFirstColumn;
        newTable.EmphasizeLastColumn = EmphasizeLastColumn;
        newTable.ShowRowStripes = ShowRowStripes;
        newTable.ShowColumnStripes = ShowColumnStripes;
        newTable.ShowAutoFilter = ShowAutoFilter;
        newTable.Theme = Theme;
        newTable._showTotalsRow = ShowTotalsRow;

        var fieldCount = ColumnCount();
        for (var f = 0; f < fieldCount; f++)
        {
            var tableField = (XLTableField)newTable.Field(f);
            var tField = (XLTableField)Field(f);
            tableField.Index = tField.Index;
            tableField.Name = tField.Name;
            tableField.totalsRowLabel = tField.totalsRowLabel;
            tableField.totalsRowFunction = tField.totalsRowFunction;
        }
        return newTable;
    }

    #region Append and replace data

    public IXLRange? AppendData(IEnumerable data, bool propagateExtraColumns = false)
    {
        return AppendData(data, transpose: false, propagateExtraColumns: propagateExtraColumns);
    }

    public IXLRange? AppendData(IEnumerable data, bool transpose, bool propagateExtraColumns = false)
    {
        var castedData = data?.Cast<object>().ToArray() ?? Array.Empty<object>();
        if (!castedData.Any() || data is string)
            return null;

        var numberOfNewRows = castedData.Length;

        var lastRowOfOldRange = DataRange!.LastRow()!;
        lastRowOfOldRange.InsertRowsBelow(numberOfNewRows);
        Fields.Cast<XLTableField>().ForEach(f => f.Column = null!);

        var insertedRange = lastRowOfOldRange.RowBelow().FirstCell().InsertData(castedData, transpose)!;

        PropagateExtraColumns(insertedRange.ColumnCount(), lastRowOfOldRange.RowNumber());

        return insertedRange;
    }

    public IXLRange? AppendData(DataTable dataTable, bool propagateExtraColumns = false)
    {
        return AppendData(dataTable.Rows.Cast<DataRow>(), propagateExtraColumns: propagateExtraColumns);
    }

    public IXLRange? AppendData<T>(IEnumerable<T> data, bool propagateExtraColumns = false)
    {
        var materializedData = data?.ToArray() ?? Array.Empty<T>();
        if (!materializedData.Any() || data is string)
            return null;

        var numberOfNewRows = materializedData.Length;

        if (numberOfNewRows == 0)
            return null;

        var lastRowOfOldRange = DataRange!.LastRow()!;
        lastRowOfOldRange.InsertRowsBelow(numberOfNewRows);
        Fields.Cast<XLTableField>().ForEach(f => f.Column = null!);

        var insertedRange = lastRowOfOldRange.RowBelow().FirstCell().InsertData(materializedData)!;

        PropagateExtraColumns(insertedRange.ColumnCount(), lastRowOfOldRange.RowNumber());

        return insertedRange;
    }

    public IXLRange ReplaceData(IEnumerable data, bool propagateExtraColumns = false)
    {
        return ReplaceData(data, transpose: false, propagateExtraColumns: propagateExtraColumns);
    }

    public IXLRange ReplaceData(IEnumerable data, bool transpose, bool propagateExtraColumns = false)
    {
        var castedData = data?.Cast<object>().ToArray() ?? Array.Empty<object>();
        if (!castedData.Any() || data is string)
            throw new InvalidOperationException("Cannot replace table data with empty enumerable.");

        var firstDataRowNumber = DataRange!.FirstRow()!.RowNumber();
        var lastDataRowNumber = DataRange.LastRow()!.RowNumber();

        // Resize table
        var sizeDifference = castedData.Length - DataRange.RowCount();
        if (sizeDifference > 0)
            DataRange.LastRow()!.InsertRowsBelow(sizeDifference);
        else if (sizeDifference < 0)
        {
            DataRange.Rows
                (
                    lastDataRowNumber + sizeDifference + 1 - firstDataRowNumber + 1,
                    lastDataRowNumber - firstDataRowNumber + 1
                )
                .Delete();

            // No propagation needed when reducing the number of rows
            propagateExtraColumns = false;
        }

        if (sizeDifference != 0)
            // Invalidate table fields' columns
            Fields.Cast<XLTableField>().ForEach(f => f.Column = null!);

        var replacedRange = DataRange.FirstCell().InsertData(castedData, transpose)!;

        if (propagateExtraColumns)
            PropagateExtraColumns(replacedRange.ColumnCount(), lastDataRowNumber);

        return replacedRange;
    }

    public IXLRange ReplaceData(DataTable dataTable, bool propagateExtraColumns = false)
    {
        return ReplaceData(dataTable.Rows.Cast<DataRow>(), propagateExtraColumns: propagateExtraColumns);
    }

    public IXLRange ReplaceData<T>(IEnumerable<T> data, bool propagateExtraColumns = false)
    {
        var materializedData = data?.ToArray() ?? Array.Empty<T>();
        if (!materializedData.Any() || data is string)
            throw new InvalidOperationException("Cannot replace table data with empty enumerable.");

        var firstDataRowNumber = DataRange!.FirstRow()!.RowNumber();
        var lastDataRowNumber = DataRange.LastRow()!.RowNumber();

        // Resize table
        var sizeDifference = materializedData.Length - DataRange.RowCount();
        if (sizeDifference > 0)
            DataRange.LastRow()!.InsertRowsBelow(sizeDifference);
        else if (sizeDifference < 0)
        {
            DataRange.Rows
                (
                    lastDataRowNumber + sizeDifference + 1 - firstDataRowNumber + 1,
                    lastDataRowNumber - firstDataRowNumber + 1
                )
                .Delete();

            // No propagation needed when reducing the number of rows
            propagateExtraColumns = false;
        }

        if (sizeDifference != 0)
            // Invalidate table fields' columns
            Fields.Cast<XLTableField>().ForEach(f => f.Column = null!);

        var replacedRange = DataRange.FirstCell().InsertData(materializedData)!;

        if (propagateExtraColumns)
            PropagateExtraColumns(replacedRange.ColumnCount(), lastDataRowNumber);

        return replacedRange;
    }

    private void PropagateExtraColumns(int numberOfNonExtraColumns, int previousLastDataRow)
    {
        for (var i = numberOfNonExtraColumns; i < Fields.Count(); i++)
        {
            var field = Field(i);

            var cell = Worksheet.Cell(previousLastDataRow, field.Column.ColumnNumber());
            field.Column.Cells(c => c.Address.RowNumber > previousLastDataRow)
                .ForEach(c =>
                {
                    if (cell.HasFormula)
                        c.FormulaR1C1 = cell.FormulaR1C1;
                    else
                        c.Value = cell.Value;
                });
        }
    }

    /// <summary>
    /// Update headers fields and totals fields by data from the cells. Do not add a new fields or names.
    /// </summary>
    /// <param name="refreshArea">Area that contains cells with changed values that might affect header and totals fields.</param>
    internal void RefreshFieldsFromCells(XLSheetRange refreshArea)
    {
        var tableArea = Area;
        if (ShowTotalsRow)
        {
            var totalsRow = tableArea.SliceFromBottom(1);
            var intersection = totalsRow.Intersect(refreshArea);
            if (intersection is not null)
            {
                var totalsRowNumber = totalsRow.BottomRow;
                var valueSlice = Worksheet.Internals.CellsCollection.ValueSlice;
                for (var column = intersection.Value.LeftColumn; column <= intersection.Value.RightColumn; ++column)
                {
                    var fieldIndex = column - totalsRow.LeftColumn;
                    var field = Field(fieldIndex);
                    var value = valueSlice.GetCellValue(new XLSheetPoint(totalsRowNumber, column));

                    // Convert value to text, because Excel always converts values to text when replacing totals row.
                    field.TotalsRowLabel = value.ToString(CultureInfo.CurrentCulture);
                }
            }
        }

        if (ShowHeaderRow)
        {
            var headersRow = Area.SliceFromTop(1);
            var intersection = headersRow.Intersect(refreshArea);
            if (intersection is not null)
            {
                var headersRowNumber = headersRow.TopRow;
                var valueSlice = Worksheet.Internals.CellsCollection.ValueSlice;
                for (var column = intersection.Value.LeftColumn; column <= intersection.Value.RightColumn; ++column)
                {
                    var fieldIndex = column - headersRow.LeftColumn;
                    var field = Field(fieldIndex);
                    var value = valueSlice.GetCellValue(new XLSheetPoint(headersRowNumber, column));

                    // Convert to text, because headers row of a table can be only
                    // string in OOXML and Excel converts it to string as well.
                    field.Name = value.ToString(CultureInfo.CurrentCulture);
                }
            }
        }
    }

    #endregion Append and replace data
}
