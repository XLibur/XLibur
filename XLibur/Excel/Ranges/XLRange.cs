
using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Tables;
using XLibur.Extensions;

namespace XLibur.Excel;

internal class XLRange : XLStoredRangeBase, IXLRange
{
    #region Constructor

    public XLRange(XLRangeParameters xlRangeParameters)
        : base(xlRangeParameters.RangeAddress, ((XLStyle)xlRangeParameters.DefaultStyle).Value)
    {
    }

    #endregion Constructor

    public override XLRangeType RangeType => XLRangeType.Range;

    #region IXLRange Members

    IXLRangeRow IXLRange.Row(int row)
    {
        return Row(row);
    }

    IXLRangeColumn IXLRange.Column(int columnNumber)
    {
        return Column(columnNumber);
    }

    IXLRangeColumn IXLRange.Column(string columnLetter)
    {
        return Column(columnLetter);
    }

    public virtual IXLRangeColumns Columns(Func<IXLRangeColumn, bool>? predicate = null)
    {
        var retVal = new XLRangeColumns();
        var columnCount = ColumnCount();
        for (var c = 1; c <= columnCount; c++)
        {
            var column = Column(c);
            if (predicate == null || predicate(column))
                retVal.Add(column);
        }

        return retVal;
    }

    public virtual IXLRangeColumns Columns(int firstColumn, int lastColumn)
    {
        var retVal = new XLRangeColumns();

        for (var co = firstColumn; co <= lastColumn; co++)
            retVal.Add(Column(co));
        return retVal;
    }

    public virtual IXLRangeColumns Columns(string firstColumn, string lastColumn)
    {
        return Columns(XLHelper.GetColumnNumberFromLetter(firstColumn),
            XLHelper.GetColumnNumberFromLetter(lastColumn));
    }

    public virtual IXLRangeColumns Columns(string columns)
    {
        ArgumentException.ThrowIfNullOrEmpty(columns);
        var retVal = new XLRangeColumns();
        foreach (var segment in columns.Split(','))
            AddColumnsFromSegment(retVal, segment.Trim());

        return retVal;
    }

    private void AddColumnsFromSegment(XLRangeColumns result, string segment)
    {
        ParseRangeSegment(segment, out var first, out var last);

        var columnsToAdd = int.TryParse(first, out _)
            ? Columns(int.Parse(first), int.Parse(last))
            : Columns(first, last);

        foreach (var col in columnsToAdd)
            result.Add(col);
    }

    private static void ParseRangeSegment(string segment, out string first, out string last)
    {
        if (segment.Contains(':') || segment.Contains('-'))
        {
            var parts = XLHelper.SplitRange(segment);
            first = parts[0];
            last = parts[1];
        }
        else
        {
            first = segment;
            last = segment;
        }
    }

    IXLCell IXLRange.Cell(int row, int column)
    {
        return Cell(row, column);
    }

    IXLCell IXLRange.Cell(string cellAddressInRange)
    {
        return Cell(cellAddressInRange)!;
    }

    IXLCell IXLRange.Cell(int row, string column)
    {
        return Cell(row, column);
    }

    IXLCell IXLRange.Cell(IXLAddress cellAddressInRange)
    {
        return Cell(cellAddressInRange);
    }

    IXLRange IXLRange.Range(IXLRangeAddress rangeAddress)
    {
        return Range(rangeAddress);
    }

    IXLRange IXLRange.Range(string rangeAddress)
    {
        return Range(rangeAddress)!;
    }

    IXLRange IXLRange.Range(IXLCell firstCell, IXLCell lastCell)
    {
        return Range(firstCell, lastCell);
    }

    IXLRange IXLRange.Range(string firstCellAddress, string lastCellAddress)
    {
        return Range(firstCellAddress, lastCellAddress);
    }

    IXLRange IXLRange.Range(IXLAddress firstCellAddress, IXLAddress lastCellAddress)
    {
        return Range(firstCellAddress, lastCellAddress);
    }

    IXLRange IXLRange.Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
    {
        return Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn);
    }

    IXLRanges IXLRange.Ranges(string ranges) => Ranges(ranges);

    public IXLRangeRows Rows(Func<IXLRangeRow, bool>? predicate = null)
    {
        var retVal = new XLRangeRows();
        var rowCount = RowCount();
        for (var r = 1; r <= rowCount; r++)
        {
            var row = Row(r);
            if (predicate == null || predicate(row))
                retVal.Add(Row(r));
        }

        return retVal;
    }

    public IXLRangeRows Rows(int firstRow, int lastRow)
    {
        var retVal = new XLRangeRows();

        for (var ro = firstRow; ro <= lastRow; ro++)
            retVal.Add(Row(ro));
        return retVal;
    }

    public IXLRangeRows Rows(string rows)
    {
        var retVal = new XLRangeRows();
        foreach (var segment in rows.Split(','))
        {
            ParseRangeSegment(segment.Trim(), out var first, out var last);
            foreach (var row in Rows(int.Parse(first), int.Parse(last)))
                retVal.Add(row);
        }

        return retVal;
    }

    public void Transpose(XLTransposeOptions transposeOption)
    {
        var rowCount = RowCount();
        var columnCount = ColumnCount();
        var squareSide = rowCount > columnCount ? rowCount : columnCount;

        var firstCell = FirstCell();

        MoveOrClearForTranspose(transposeOption, rowCount, columnCount);
        TransposeMerged(squareSide);
        TransposeRange(squareSide);
        RangeAddress = new XLRangeAddress(
            RangeAddress.FirstAddress,
            new XLAddress(Worksheet,
                firstCell.Address.RowNumber + columnCount - 1,
                firstCell.Address.ColumnNumber + rowCount - 1,
                RangeAddress.LastAddress.FixedRow,
                RangeAddress.LastAddress.FixedColumn));

        if (rowCount > columnCount)
        {
            var rng = Worksheet.Range(
                RangeAddress.LastAddress.RowNumber + 1,
                RangeAddress.FirstAddress.ColumnNumber,
                RangeAddress.LastAddress.RowNumber + (rowCount - columnCount),
                RangeAddress.LastAddress.ColumnNumber);
            rng.Delete(XLShiftDeletedCells.ShiftCellsUp);
        }
        else if (columnCount > rowCount)
        {
            var rng = Worksheet.Range(
                RangeAddress.FirstAddress.RowNumber,
                RangeAddress.LastAddress.ColumnNumber + 1,
                RangeAddress.LastAddress.RowNumber,
                RangeAddress.LastAddress.ColumnNumber + (columnCount - rowCount));
            rng.Delete(XLShiftDeletedCells.ShiftCellsLeft);
        }

        foreach (var c in Range(1, 1, columnCount, rowCount).Cells())
        {
            var border = ((XLStyle)c.Style).Value.Border;
            c.Style.Border.TopBorder = border.LeftBorder;
            c.Style.Border.TopBorderColor = border.LeftBorderColor;
            c.Style.Border.LeftBorder = border.TopBorder;
            c.Style.Border.LeftBorderColor = border.TopBorderColor;
            c.Style.Border.RightBorder = border.BottomBorder;
            c.Style.Border.RightBorderColor = border.BottomBorderColor;
            c.Style.Border.BottomBorder = border.RightBorder;
            c.Style.Border.BottomBorderColor = border.RightBorderColor;
        }
    }

    public IXLTable AsTable()
    {
        return Worksheet.Table(this, false);
    }

    public IXLTable AsTable(string name)
    {
        return Worksheet.Table(this, name, false);
    }

    IXLTable IXLRange.CreateTable()
    {
        return CreateTable();
    }

    public XLTable CreateTable()
    {
        return (XLTable)Worksheet.Table(this, true);
    }

    IXLTable IXLRange.CreateTable(string name)
    {
        return CreateTable(name);
    }

    public XLTable CreateTable(string name)
    {
        return (XLTable)Worksheet.Table(this, name, true);
    }

    public IXLTable CreateTable(string name, bool setAutofilter)
    {
        return Worksheet.Table(this, name, true, setAutofilter);
    }

    public IXLRange CopyTo(IXLCell target)
    {
        base.CopyToCell((XLCell)target);

        var lastRowNumber = target.Address.RowNumber + RowCount() - 1;
        if (lastRowNumber > XLHelper.MaxRowNumber)
            lastRowNumber = XLHelper.MaxRowNumber;
        var lastColumnNumber = target.Address.ColumnNumber + ColumnCount() - 1;
        if (lastColumnNumber > XLHelper.MaxColumnNumber)
            lastColumnNumber = XLHelper.MaxColumnNumber;

        return target.Worksheet.Range(target.Address.RowNumber,
            target.Address.ColumnNumber,
            lastRowNumber,
            lastColumnNumber);
    }

    public new IXLRange CopyTo(IXLRangeBase target)
    {
        base.CopyTo(target);

        var lastRowNumber = target.RangeAddress.FirstAddress.RowNumber + RowCount() - 1;
        if (lastRowNumber > XLHelper.MaxRowNumber)
            lastRowNumber = XLHelper.MaxRowNumber;
        var lastColumnNumber = target.RangeAddress.FirstAddress.ColumnNumber + ColumnCount() - 1;
        if (lastColumnNumber > XLHelper.MaxColumnNumber)
            lastColumnNumber = XLHelper.MaxColumnNumber;

        return target.Worksheet.Range(target.RangeAddress.FirstAddress.RowNumber,
            target.RangeAddress.FirstAddress.ColumnNumber,
            lastRowNumber,
            lastColumnNumber);
    }

    public new IXLRange Sort()
    {
        return base.Sort().AsRange();
    }

    public new IXLRange Sort(string columnsToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending,
        bool matchCase = false, bool ignoreBlanks = true)
    {
        return base.Sort(columnsToSortBy, sortOrder, matchCase, ignoreBlanks).AsRange();
    }

    public new IXLRange Sort(int columnToSortBy, XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false,
        bool ignoreBlanks = true)
    {
        return base.Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks).AsRange();
    }

    public new IXLRange SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false,
        bool ignoreBlanks = true)
    {
        return base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks).AsRange();
    }

    #endregion IXLRange Members

    internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
    {
        RangeAddress = (XLRangeAddress)ShiftColumns(RangeAddress, range, columnsShifted);
    }

    internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
    {
        RangeAddress = (XLRangeAddress)ShiftRows(RangeAddress, range, rowsShifted);
    }

    IXLRangeColumn? IXLRange.FirstColumn(Func<IXLRangeColumn, bool>? predicate)
    {
        return FirstColumn(predicate);
    }

    internal XLRangeColumn? FirstColumn(Func<IXLRangeColumn, bool>? predicate = null)
    {
        if (predicate == null)
            return Column(1);

        var columnCount = ColumnCount();
        for (var c = 1; c <= columnCount; c++)
        {
            var column = Column(c);
            if (predicate(column)) return column;
        }

        return null;
    }

    IXLRangeColumn? IXLRange.LastColumn(Func<IXLRangeColumn, bool>? predicate)
    {
        return LastColumn(predicate);
    }

    internal XLRangeColumn? LastColumn(Func<IXLRangeColumn, bool>? predicate = null)
    {
        var columnCount = ColumnCount();
        if (predicate == null)
            return Column(columnCount);

        for (var c = columnCount; c >= 1; c--)
        {
            var column = Column(c);
            if (predicate(column)) return column;
        }

        return null;
    }

    IXLRangeColumn? IXLRange.FirstColumnUsed(Func<IXLRangeColumn, bool>? predicate)
    {
        return FirstColumnUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    internal XLRangeColumn? FirstColumnUsed(Func<IXLRangeColumn, bool>? predicate = null)
    {
        return FirstColumnUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    IXLRangeColumn? IXLRange.FirstColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, bool>? predicate)
    {
        return FirstColumnUsed(options, predicate);
    }

    internal XLRangeColumn? FirstColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, bool>? predicate = null)
    {
        if (predicate == null)
        {
            var firstColumnUsed = Worksheet.Internals.CellsCollection.FirstColumnUsed(
                XLSheetRange.FromRangeAddress(RangeAddress),
                options);

            return firstColumnUsed == 0 ? null : Column(firstColumnUsed - RangeAddress.FirstAddress.ColumnNumber + 1);
        }

        var columnCount = ColumnCount();
        for (var co = 1; co <= columnCount; co++)
        {
            var column = Column(co);

            if (!column.IsEmpty(options) && predicate(column))
                return column;
        }

        return null;
    }

    IXLRangeColumn? IXLRange.LastColumnUsed(Func<IXLRangeColumn, bool>? predicate)
    {
        return LastColumnUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    internal XLRangeColumn? LastColumnUsed(Func<IXLRangeColumn, bool>? predicate = null)
    {
        return LastColumnUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    IXLRangeColumn? IXLRange.LastColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, bool>? predicate)
    {
        return LastColumnUsed(options, predicate);
    }

    internal XLRangeColumn? LastColumnUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, bool>? predicate = null)
    {
        if (predicate == null)
        {
            var lastColumnUsed = Worksheet.Internals.CellsCollection.LastColumnUsed(
                XLSheetRange.FromRangeAddress(RangeAddress),
                options);

            return lastColumnUsed == 0 ? null : Column(lastColumnUsed - RangeAddress.FirstAddress.ColumnNumber + 1);
        }

        var columnCount = ColumnCount();
        for (var co = columnCount; co >= 1; co--)
        {
            var column = Column(co);

            if (!column.IsEmpty(options) && predicate(column))
                return column;
        }

        return null;
    }

    IXLRangeRow? IXLRange.FirstRow(Func<IXLRangeRow, bool>? predicate)
    {
        return FirstRow(predicate);
    }

    public XLRangeRow? FirstRow(Func<IXLRangeRow, bool>? predicate = null)
    {
        if (predicate == null)
            return Row(1);

        var rowCount = RowCount();
        for (var ro = 1; ro <= rowCount; ro++)
        {
            var row = Row(ro);
            if (predicate(row)) return row;
        }

        return null;
    }

    IXLRangeRow? IXLRange.LastRow(Func<IXLRangeRow, bool>? predicate)
    {
        return LastRow(predicate);
    }

    public XLRangeRow? LastRow(Func<IXLRangeRow, bool>? predicate = null)
    {
        var rowCount = RowCount();
        if (predicate == null)
            return Row(rowCount);

        for (var ro = rowCount; ro >= 1; ro--)
        {
            var row = Row(ro);
            if (predicate(row)) return row;
        }

        return null;
    }

    IXLRangeRow? IXLRange.FirstRowUsed(Func<IXLRangeRow, bool>? predicate)
    {
        return FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    internal XLRangeRow? FirstRowUsed(Func<IXLRangeRow, bool>? predicate = null)
    {
        return FirstRowUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    IXLRangeRow? IXLRange.FirstRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, bool>? predicate)
    {
        return FirstRowUsed(options, predicate);
    }

    internal XLRangeRow? FirstRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, bool>? predicate = null)
    {
        if (predicate == null)
        {
            var rowFromCells = Worksheet.Internals.CellsCollection.FirstRowUsed(
                XLSheetRange.FromRangeAddress(RangeAddress), options);

            return rowFromCells == 0 ? null : Row(rowFromCells - RangeAddress.FirstAddress.RowNumber + 1);
        }

        var rowCount = RowCount();
        for (var ro = 1; ro <= rowCount; ro++)
        {
            var row = Row(ro);

            if (!row.IsEmpty(options) && predicate(row))
                return row;
        }

        return null;
    }

    IXLRangeRow? IXLRange.LastRowUsed(Func<IXLRangeRow, bool>? predicate)
    {
        return LastRowUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    internal XLRangeRow? LastRowUsed(Func<IXLRangeRow, bool>? predicate = null)
    {
        return LastRowUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    IXLRangeRow? IXLRange.LastRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, bool>? predicate)
    {
        return LastRowUsed(options, predicate);
    }

    internal XLRangeRow? LastRowUsed(XLCellsUsedOptions options, Func<IXLRangeRow, bool>? predicate = null)
    {
        if (predicate == null)
        {
            var lastRowUsed = Worksheet.Internals.CellsCollection.LastRowUsed(
                XLSheetRange.FromRangeAddress(RangeAddress), options);

            return lastRowUsed == 0 ? null : Row(lastRowUsed - RangeAddress.FirstAddress.RowNumber + 1);
        }

        var rowCount = RowCount();
        for (var ro = rowCount; ro >= 1; ro--)
        {
            var row = Row(ro);

            if (!row.IsEmpty(options) && predicate(row))
                return row;
        }

        return null;
    }

    IXLRangeRows IXLRange.RowsUsed(XLCellsUsedOptions options, Func<IXLRangeRow, bool>? predicate)
    {
        return RowsUsed(options, predicate);
    }

    internal XLRangeRows RowsUsed(XLCellsUsedOptions options, Func<IXLRangeRow, bool>? predicate = null)
    {
        var rows = new XLRangeRows();
        var rowCount = RowCount(options);

        for (var ro = 1; ro <= rowCount; ro++)
        {
            var row = Row(ro);

            if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
                rows.Add(row);
        }

        return rows;
    }

    IXLRangeRows IXLRange.RowsUsed(Func<IXLRangeRow, bool>? predicate)
    {
        return RowsUsed(predicate);
    }

    internal XLRangeRows RowsUsed(Func<IXLRangeRow, bool>? predicate = null)
    {
        return RowsUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    IXLRangeColumns IXLRange.ColumnsUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, bool>? predicate)
    {
        return ColumnsUsed(options, predicate);
    }

    internal virtual XLRangeColumns ColumnsUsed(XLCellsUsedOptions options, Func<IXLRangeColumn, bool>? predicate = null)
    {
        var columns = new XLRangeColumns();
        var columnCount = ColumnCount(options);

        for (var co = 1; co <= columnCount; co++)
        {
            var column = Column(co);

            if (!column.IsEmpty(options) && (predicate == null || predicate(column)))
                columns.Add(column);
        }

        return columns;
    }

    IXLRangeColumns IXLRange.ColumnsUsed(Func<IXLRangeColumn, bool>? predicate)
    {
        return ColumnsUsed(predicate);
    }

    internal virtual XLRangeColumns ColumnsUsed(Func<IXLRangeColumn, bool>? predicate = null)
    {
        return ColumnsUsed(XLCellsUsedOptions.AllContents, predicate);
    }

    public XLRangeRow Row(int row)
    {
        if (row <= 0 || row > XLHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1)
            throw new ArgumentOutOfRangeException(nameof(row),
                $"Row number must be between 1 and {XLHelper.MaxRowNumber + RangeAddress.FirstAddress.RowNumber - 1}");

        var firstCellAddress = new XLAddress(Worksheet,
            RangeAddress.FirstAddress.RowNumber + row - 1,
            RangeAddress.FirstAddress.ColumnNumber,
            false,
            false);

        var lastCellAddress = new XLAddress(Worksheet,
            RangeAddress.FirstAddress.RowNumber + row - 1,
            RangeAddress.LastAddress.ColumnNumber,
            false,
            false);
        return Worksheet.RangeRow(new XLRangeAddress(firstCellAddress, lastCellAddress));
    }

    public virtual XLRangeColumn Column(int columnNumber)
    {
        if (columnNumber <= 0 || columnNumber > XLHelper.MaxColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1)
            throw new ArgumentOutOfRangeException(nameof(columnNumber),
                $"Column number must be between 1 and {XLHelper.MaxColumnNumber + RangeAddress.FirstAddress.ColumnNumber - 1}");

        var firstCellAddress = new XLAddress(Worksheet,
            RangeAddress.FirstAddress.RowNumber,
            RangeAddress.FirstAddress.ColumnNumber + columnNumber - 1,
            false,
            false);
        var lastCellAddress = new XLAddress(Worksheet,
            RangeAddress.LastAddress.RowNumber,
            RangeAddress.FirstAddress.ColumnNumber + columnNumber - 1,
            false,
            false);
        return Worksheet.RangeColumn(new XLRangeAddress(firstCellAddress, lastCellAddress));
    }

    public virtual XLRangeColumn Column(string columnLetter)
    {
        return Column(XLHelper.GetColumnNumberFromLetter(columnLetter));
    }

    internal IEnumerable<XLRange> Split(IXLRangeAddress anotherRange, bool includeIntersection)
    {
        if (!RangeAddress.Intersects(anotherRange))
        {
            yield return this;
            yield break;
        }

        var thisRow1 = RangeAddress.FirstAddress.RowNumber;
        var thisRow2 = RangeAddress.LastAddress.RowNumber;
        var thisColumn1 = RangeAddress.FirstAddress.ColumnNumber;
        var thisColumn2 = RangeAddress.LastAddress.ColumnNumber;

        var otherRow1 = Math.Min(Math.Max(thisRow1, anotherRange.FirstAddress.RowNumber), thisRow2 + 1);
        var otherRow2 = Math.Max(Math.Min(thisRow2, anotherRange.LastAddress.RowNumber), thisRow1 - 1);
        var otherColumn1 = Math.Min(Math.Max(thisColumn1, anotherRange.FirstAddress.ColumnNumber), thisColumn2 + 1);
        var otherColumn2 = Math.Max(Math.Min(thisColumn2, anotherRange.LastAddress.ColumnNumber), thisColumn1 - 1);

        var candidates = new[]
        {
            // to the top of the intersection
            new XLRangeAddress(
                new XLAddress(thisRow1, thisColumn1, false, false),
                new XLAddress(otherRow1 - 1, thisColumn2, false, false)),

            // to the left of the intersection
            new XLRangeAddress(
                new XLAddress(otherRow1, thisColumn1, false, false),
                new XLAddress(otherRow2, otherColumn1 - 1, false, false)),

            includeIntersection
                ? new XLRangeAddress(
                    new XLAddress(otherRow1, otherColumn1, false, false),
                    new XLAddress(otherRow2, otherColumn2, false, false))
                : XLRangeAddress.Invalid,

            // to the right of the intersection
            new XLRangeAddress(
                new XLAddress(otherRow1, otherColumn2 + 1, false, false),
                new XLAddress(otherRow2, thisColumn2, false, false)),

            // to the bottom of the intersection
            new XLRangeAddress(
                new XLAddress(otherRow2 + 1, thisColumn1, false, false),
                new XLAddress(thisRow2, thisColumn2, false, false)),
        };

        foreach (var rangeAddress in candidates.Where(c => c is { IsValid: true, IsNormalized: true }))
        {
            yield return Worksheet.Range(rangeAddress);
        }
    }

    private void TransposeRange(int squareSide)
    {
        var rowOffset = RangeAddress.FirstAddress.RowNumber - 1;
        var colOffset = RangeAddress.FirstAddress.ColumnNumber - 1;
        for (var row = 1; row <= squareSide; ++row)
        {
            for (var col = row + 1; col <= squareSide; ++col)
            {
                var oldAddress = new XLSheetPoint(row + rowOffset, col + colOffset);
                var newAddress = new XLSheetPoint(col + colOffset, row + rowOffset);
                Worksheet.Internals.CellsCollection.SwapCellsContent(oldAddress, newAddress);
            }
        }
    }

    private void TransposeMerged(int squareSide)
    {
        var rngToTranspose = Worksheet.Range(
            RangeAddress.FirstAddress.RowNumber,
            RangeAddress.FirstAddress.ColumnNumber,
            RangeAddress.FirstAddress.RowNumber + squareSide - 1,
            RangeAddress.FirstAddress.ColumnNumber + squareSide - 1);

        foreach (var merge in Worksheet.Internals.MergedRanges.Where(Contains).Cast<XLRange>())
        {
            merge.RangeAddress = new XLRangeAddress(
                merge.RangeAddress.FirstAddress,
                rngToTranspose.Cell(merge.ColumnCount(), merge.RowCount()).Address);
        }
    }

    private void MoveOrClearForTranspose(XLTransposeOptions transposeOption, int rowCount, int columnCount)
    {
        if (transposeOption == XLTransposeOptions.MoveCells)
        {
            if (rowCount > columnCount)
                InsertColumnsAfter(false, rowCount - columnCount, false);
            else if (columnCount > rowCount)
                InsertRowsBelow(false, columnCount - rowCount, false);
        }
        else
        {
            if (rowCount > columnCount)
            {
                var toMove = rowCount - columnCount;
                var rngToClear = Worksheet.Range(
                    RangeAddress.FirstAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber + 1,
                    RangeAddress.LastAddress.RowNumber,
                    RangeAddress.LastAddress.ColumnNumber + toMove);
                rngToClear.Clear();
            }
            else if (columnCount > rowCount)
            {
                var toMove = columnCount - rowCount;
                var rngToClear = Worksheet.Range(
                    RangeAddress.LastAddress.RowNumber + 1,
                    RangeAddress.FirstAddress.ColumnNumber,
                    RangeAddress.LastAddress.RowNumber + toMove,
                    RangeAddress.LastAddress.ColumnNumber);
                rngToClear.Clear();
            }
        }
    }

    public override bool Equals(object? obj)
    {
        var other = obj as XLRange;
        if (other == null)
            return false;
        return RangeAddress.Equals(other.RangeAddress)
               && Worksheet.Equals(other.Worksheet);
    }

    public override int GetHashCode()
    {
        return RangeAddress.GetHashCode()
               ^ Worksheet.GetHashCode();
    }

    public new IXLRange Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public IXLRangeColumn? FindColumn(Func<IXLRangeColumn, bool> predicate)
    {
        var columnCount = ColumnCount();
        for (var c = 1; c <= columnCount; c++)
        {
            var column = Column(c);
            if (predicate == null || predicate(column))
                return column;
        }

        return null;
    }

    public IXLRangeRow? FindRow(Func<IXLRangeRow, bool> predicate)
    {
        var rowCount = RowCount();
        for (var r = 1; r <= rowCount; r++)
        {
            var row = Row(r);
            if (predicate(row))
                return row;
        }

        return null;
    }

    public override string ToString()
    {
        if (IsEntireSheet())
        {
            return Worksheet.Name;
        }

        if (IsEntireRow())
        {
            return string.Concat(
                Worksheet.Name.EscapeSheetName(),
                '!',
                RangeAddress.FirstAddress.RowNumber,
                ':',
                RangeAddress.LastAddress.RowNumber);
        }

        if (IsEntireColumn())
        {
            return string.Concat(
                Worksheet.Name.EscapeSheetName(),
                '!',
                RangeAddress.FirstAddress.ColumnLetter,
                ':',
                RangeAddress.LastAddress.ColumnLetter);
        }

        return base.ToString();
    }
}
