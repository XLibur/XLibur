using System;
using System.Linq;
using XLibur.Excel.Ranges;
using XLibur.Excel.Tables;

namespace XLibur.Excel;

internal sealed class XLRangeColumn : XLStoredRangeBase, IXLRangeColumn
{
    #region Constructor

    /// <summary>
    /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
    /// </summary>
    public XLRangeColumn(XLRangeParameters rangeParameters)
        : base(rangeParameters.RangeAddress, ((XLStyle)rangeParameters.DefaultStyle).Value)
    {
    }

    #endregion Constructor

    #region IXLRangeColumn Members

    IXLCell IXLRangeColumn.Cell(int rowNumber)
    {
        return Cell(rowNumber);
    }

    IXLCells IXLRangeColumn.Cells(string cellsInColumn) => Cells(cellsInColumn);

    public override XLCells Cells(string cells)
    {
        var retVal = new XLCells(false, XLCellsUsedOptions.AllContents);
        var rangePairs = cells.Split(',');
        foreach (var pair in rangePairs)
            retVal.Add(Range(pair.Trim()).RangeAddress);
        return retVal;
    }

    public IXLCells Cells(int firstRow, int lastRow)
    {
        return Cells(firstRow + ":" + lastRow);
    }

    public void Delete()
    {
        Delete(true);
    }

    internal void Delete(bool deleteTableField)
    {
        if (deleteTableField && IsTableColumn())
        {
            var table = (XLTable)Table!;
            if (!Cell(1).Value.TryGetText(out var firstCellValue))
                throw new InvalidOperationException("Top cell doesn't contain a text.");

            if (!table.FieldNames.ContainsKey(firstCellValue!))
                throw new InvalidOperationException($"Field {firstCellValue} not found.");

            var field = table.Fields.Cast<XLTableField>().Single(f => f.Name == firstCellValue);
            field.Delete(false);
        }

        Delete(XLShiftDeletedCells.ShiftCellsLeft);
    }

    public IXLCells InsertCellsAbove(int numberOfRows)
    {
        return InsertCellsAbove(numberOfRows, false);
    }

    public IXLCells InsertCellsAbove(int numberOfRows, bool expandRange)
    {
        return InsertRowsAbove(numberOfRows, expandRange).Cells();
    }

    public IXLCells InsertCellsBelow(int numberOfRows)
    {
        return InsertCellsBelow(numberOfRows, true);
    }

    public IXLCells InsertCellsBelow(int numberOfRows, bool expandRange)
    {
        return InsertRowsBelow(numberOfRows, expandRange).Cells();
    }

    public int CellCount()
    {
        return RangeAddress.LastAddress.RowNumber - RangeAddress.FirstAddress.RowNumber + 1;
    }

    public IXLRangeColumn Sort(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false, bool ignoreBlanks = true)
    {
        base.Sort(1, sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    public IXLRangeColumn CopyTo(IXLCell target)
    {
        base.CopyToCell((XLCell)target);
        return BuildCopyResult(target.Worksheet, target.Address.RowNumber, target.Address.ColumnNumber);
    }

    public new IXLRangeColumn CopyTo(IXLRangeBase target)
    {
        base.CopyTo(target);
        return BuildCopyResult(target.Worksheet, target.RangeAddress.FirstAddress.RowNumber, target.RangeAddress.FirstAddress.ColumnNumber);
    }

    private IXLRangeColumn BuildCopyResult(IXLWorksheet worksheet, int startRow, int startColumn)
    {
        var lastRowNumber = Math.Min(startRow + RowCount() - 1, XLHelper.MaxRowNumber);
        var lastColumnNumber = Math.Min(startColumn + ColumnCount() - 1, XLHelper.MaxColumnNumber);
        return worksheet.Range(startRow, startColumn, lastRowNumber, lastColumnNumber).Column(1);
    }

    public IXLRangeColumn Column(int start, int end)
    {
        return Range(start, end).FirstColumn()!;
    }

    public IXLRangeColumn Column(IXLCell start, IXLCell end)
    {
        return Column(start.Address.RowNumber, end.Address.RowNumber);
    }

    public IXLRangeColumns Columns(string columns)
    {
        var retVal = new XLRangeColumns();
        var rowPairs = columns.Split(',');
        foreach (var trimmedPair in rowPairs.Select(pair => pair.Trim()))
        {
            string firstRow;
            string lastRow;
            if (trimmedPair.Contains(':') || trimmedPair.Contains('-'))
            {
                var rowRange = trimmedPair.Split(':', '-');

                firstRow = rowRange[0];
                lastRow = rowRange[1];
            }
            else
            {
                firstRow = trimmedPair;
                lastRow = trimmedPair;
            }

            retVal.Add(Range(firstRow, lastRow).FirstColumn()!);
        }

        return retVal;
    }

    public IXLColumn WorksheetColumn()
    {
        return Worksheet.Column(RangeAddress.FirstAddress.ColumnNumber);
    }

    #endregion IXLRangeColumn Members

    public override XLRangeType RangeType
    {
        get { return XLRangeType.RangeColumn; }
    }

    public XLCell Cell(int row)
    {
        return Cell(row, 1);
    }

    internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
    {
        RangeAddress = (XLRangeAddress)ShiftColumns(RangeAddress, range, columnsShifted);
    }

    internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
    {
        RangeAddress = (XLRangeAddress)ShiftRows(RangeAddress, range, rowsShifted);
    }

    public XLRange Range(int firstRow, int lastRow)
    {
        return Range(firstRow, 1, lastRow, 1);
    }

    public override XLRange Range(string rangeAddressStr)
    {
        string rangeAddressToUse;
        if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
        {
            if (rangeAddressStr.Contains('-'))
                rangeAddressStr = rangeAddressStr.Replace('-', ':');

            var arrRange = rangeAddressStr.Split(':');
            var firstPart = arrRange[0];
            var secondPart = arrRange[1];
            rangeAddressToUse = FixColumnAddress(firstPart) + ":" + FixColumnAddress(secondPart);
        }
        else
            rangeAddressToUse = FixColumnAddress(rangeAddressStr);

        var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
        return Range(rangeAddress);
    }

    public int CompareTo(XLRangeColumn otherColumn, IXLSortElements rowsToSort)
    {
        foreach (var e in rowsToSort)
        {
            var thisCell = Cell(e.ElementNumber);
            var otherCell = otherColumn.Cell(e.ElementNumber);
            var comparison = CompareColumnCells(thisCell, otherCell, e);

            if (comparison != 0)
                return e.SortOrder == XLSortOrder.Ascending ? comparison : comparison * -1;
        }

        return 0;
    }

    private static int CompareColumnCells(XLCell thisCell, XLCell otherCell, IXLSortElement e)
    {
        return XLCellComparer.CompareCells(thisCell, otherCell, e);
    }

    private XLRangeColumn ColumnShift(int columnsToShift)
    {
        var columnNumber = ColumnNumber() + columnsToShift;
        return Worksheet.Range(
            RangeAddress.FirstAddress.RowNumber,
            columnNumber,
            RangeAddress.LastAddress.RowNumber,
            columnNumber).FirstColumn()!;
    }

    #region XLRangeColumn Left

    IXLRangeColumn IXLRangeColumn.ColumnLeft()
    {
        return ColumnLeft();
    }

    IXLRangeColumn IXLRangeColumn.ColumnLeft(int step)
    {
        return ColumnLeft(step);
    }

    public XLRangeColumn ColumnLeft()
    {
        return ColumnLeft(1);
    }

    public XLRangeColumn ColumnLeft(int step)
    {
        return ColumnShift(step * -1);
    }

    #endregion XLRangeColumn Left

    #region XLRangeColumn Right

    IXLRangeColumn IXLRangeColumn.ColumnRight()
    {
        return ColumnRight();
    }

    IXLRangeColumn IXLRangeColumn.ColumnRight(int step)
    {
        return ColumnRight(step);
    }

    public XLRangeColumn ColumnRight()
    {
        return ColumnRight(1);
    }

    public XLRangeColumn ColumnRight(int step)
    {
        return ColumnShift(step);
    }

    #endregion XLRangeColumn Right

    public IXLTable AsTable()
    {
        ThrowIfTableColumn();
        return AsRange().AsTable();
    }

    public IXLTable AsTable(string name)
    {
        ThrowIfTableColumn();
        return AsRange().AsTable(name);
    }

    public IXLTable CreateTable()
    {
        ThrowIfTableColumn();
        return AsRange().CreateTable();
    }

    public IXLTable CreateTable(string name)
    {
        ThrowIfTableColumn();
        return AsRange().CreateTable(name);
    }

    private void ThrowIfTableColumn()
    {
        if (IsTableColumn())
            throw new InvalidOperationException("This column is already part of a table.");
    }

    public new IXLRangeColumn Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }


    public IXLRangeColumn ColumnUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents)
    {
        return Column((this as IXLRangeBase).FirstCellUsed(options)!,
            (this as IXLRangeBase).LastCellUsed(options)!);
    }

    internal IXLTable? Table { get; set; }

    public bool IsTableColumn()
    {
        return Table != null;
    }
}
