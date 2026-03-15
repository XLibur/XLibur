using System.Linq;

namespace XLibur.Excel;

/// <summary>
/// Contains the heavy algorithmic logic for inserting rows and columns into a range.
/// <see cref="XLRangeBase"/> delegates to these methods to keep the main class smaller.
/// </summary>
internal static class XLRangeInsertHelper
{
    internal static IXLRangeColumns? InsertColumnsBefore(XLRangeBase range, bool onlyUsedCells, int numberOfColumns, bool formatFromLeft, bool nullReturn)
    {
        if (numberOfColumns <= 0 || numberOfColumns > XLHelper.MaxColumnNumber)
            throw new System.ArgumentOutOfRangeException(nameof(numberOfColumns),
                $"Number of columns to insert must be a positive number no more than {XLHelper.MaxColumnNumber}");

        ShiftFormulasForColumns(range, numberOfColumns);

        range.Worksheet.SparklineGroupsInternal.ShiftColumns(XLSheetRange.FromRangeAddress(range.RangeAddress), numberOfColumns);

        ShiftColumnWidths(range, onlyUsedCells, numberOfColumns);

        var insertedRange = new XLSheetRange(
            XLSheetPoint.FromAddress(range.RangeAddress.FirstAddress),
            new XLSheetPoint(range.RangeAddress.LastAddress.RowNumber, range.RangeAddress.FirstAddress.ColumnNumber + numberOfColumns - 1));

        range.Worksheet.Internals.CellsCollection.InsertAreaAndShiftRight(insertedRange);

        var firstRowReturn = range.RangeAddress.FirstAddress.RowNumber;
        var lastRowReturn = range.RangeAddress.LastAddress.RowNumber;
        var firstColumnReturn = range.RangeAddress.FirstAddress.ColumnNumber;
        var lastColumnReturn = range.RangeAddress.FirstAddress.ColumnNumber + numberOfColumns - 1;

        range.Worksheet.NotifyRangeShiftedColumns(range.AsRange(), numberOfColumns);

        var rangeToReturn = range.Worksheet.Range(firstRowReturn, firstColumnReturn, lastRowReturn, lastColumnReturn);

        var contentFlags = XLCellsUsedOptions.All
                           & ~XLCellsUsedOptions.ConditionalFormats
                           & ~XLCellsUsedOptions.DataValidation;

        ApplyColumnFormatting(range, rangeToReturn, formatFromLeft, contentFlags);

        if (nullReturn)
            return null;

        return rangeToReturn.Columns();
    }

    private static void ShiftFormulasForColumns(XLRangeBase range, int numberOfColumns)
    {
        foreach (var ws in range.Worksheet.Workbook.WorksheetsInternal)
        {
            foreach (var cell in ws.Internals.CellsCollection.GetCells(c => c.Formula is not null))
                cell.ShiftFormulaColumns(range.AsRange(), numberOfColumns);
        }
    }

    private static void ShiftColumnWidths(XLRangeBase range, bool onlyUsedCells, int numberOfColumns)
    {
        if (onlyUsedCells)
            return;

        var lastColumn = range.Worksheet.Internals.CellsCollection.MaxColumnUsed;
        if (lastColumn <= 0)
            return;

        var firstColumn = range.RangeAddress.FirstAddress.ColumnNumber;
        for (var co = lastColumn; co >= firstColumn; co--)
        {
            var newColumn = co + numberOfColumns;
            if (range.IsEntireColumn())
                range.Worksheet.Column(newColumn).Width = range.Worksheet.Column(co).Width;
        }
    }

    private static void ApplyColumnFormatting(XLRangeBase range, IXLRange rangeToReturn, bool formatFromLeft, XLCellsUsedOptions contentFlags)
    {
        if (formatFromLeft && rangeToReturn.RangeAddress.FirstAddress.ColumnNumber > 1)
            ApplyColumnFormattingFromLeft(rangeToReturn, contentFlags);
        else
            ApplyColumnFormattingFromExistingRows(range, rangeToReturn, contentFlags);
    }

    private static void ApplyColumnFormattingFromLeft(IXLRange rangeToReturn, XLCellsUsedOptions contentFlags)
    {
        var firstColumnUsed = rangeToReturn!.FirstColumn()!;
        var model = firstColumnUsed.ColumnLeft()!;
        var modelFirstRow = ((IXLRangeBase)model).FirstCellUsed(contentFlags);
        var modelLastRow = ((IXLRangeBase)model).LastCellUsed(contentFlags);
        if (modelFirstRow == null || modelLastRow == null)
            return;

        var firstRoReturned = modelFirstRow.Address.RowNumber
            - model.RangeAddress.FirstAddress.RowNumber + 1;
        var lastRoReturned = modelLastRow.Address.RowNumber
            - model.RangeAddress.FirstAddress.RowNumber + 1;
        for (var ro = firstRoReturned; ro <= lastRoReturned; ro++)
            rangeToReturn.Row(ro).Style = model.Cell(ro).Style;
    }

    private static void ApplyColumnFormattingFromExistingRows(XLRangeBase range, IXLRange rangeToReturn, XLCellsUsedOptions contentFlags)
    {
        var lastRoUsed = rangeToReturn.LastRowUsed(contentFlags);
        if (lastRoUsed == null)
            return;

        var firstWsRow = rangeToReturn.RangeAddress.FirstAddress.RowNumber;
        var lastRoReturned = lastRoUsed.RowNumber() - firstWsRow + 1;
        for (var ro = 1; ro <= lastRoReturned; ro++)
        {
            var wsRow = firstWsRow + ro - 1;
            var styleToUse =
                range.Worksheet.Internals.RowsCollection.TryGetValue(wsRow, out var row)
                    ? row.Style
                    : range.Worksheet.Style;

            rangeToReturn.Row(ro).Style = styleToUse;
        }
    }

    internal static IXLRangeRows? InsertRowsAbove(XLRangeBase range, bool onlyUsedCells, int numberOfRows, bool formatFromAbove, bool nullReturn)
    {
        if (numberOfRows <= 0 || numberOfRows > XLHelper.MaxRowNumber)
            throw new System.ArgumentOutOfRangeException(nameof(numberOfRows),
                $"Number of rows to insert must be a positive number no more than {XLHelper.MaxRowNumber}");

        ShiftFormulasForRows(range, numberOfRows);

        range.Worksheet.SparklineGroupsInternal.ShiftRows(XLSheetRange.FromRangeAddress(range.RangeAddress), numberOfRows);

        ShiftRowHeights(range, onlyUsedCells, numberOfRows);

        var insertedRange = new XLSheetRange(
            XLSheetPoint.FromAddress(range.RangeAddress.FirstAddress),
            new XLSheetPoint(range.RangeAddress.FirstAddress.RowNumber + numberOfRows - 1, range.RangeAddress.LastAddress.ColumnNumber));
        range.Worksheet.Internals.CellsCollection.InsertAreaAndShiftDown(insertedRange);

        var firstRowReturn = range.RangeAddress.FirstAddress.RowNumber;
        var lastRowReturn = range.RangeAddress.FirstAddress.RowNumber + numberOfRows - 1;
        var firstColumnReturn = range.RangeAddress.FirstAddress.ColumnNumber;
        var lastColumnReturn = range.RangeAddress.LastAddress.ColumnNumber;

        range.Worksheet.NotifyRangeShiftedRows(range.AsRange(), numberOfRows);

        var rangeToReturn = range.Worksheet.Range(firstRowReturn, firstColumnReturn, lastRowReturn, lastColumnReturn);

        var contentFlags = XLCellsUsedOptions.All
                           & ~XLCellsUsedOptions.ConditionalFormats
                           & ~XLCellsUsedOptions.DataValidation;

        ApplyRowFormatting(range, rangeToReturn, formatFromAbove, contentFlags);

        // Skip calling .Rows() for performance reasons if required.
        if (nullReturn)
            return null;

        return rangeToReturn.Rows();
    }

    private static void ShiftFormulasForRows(XLRangeBase range, int numberOfRows)
    {
        var asRange = range.AsRange();
        foreach (var ws in range.Worksheet.Workbook.WorksheetsInternal)
        {
            foreach (var cell in ws.Internals.CellsCollection.GetCells(c => c.Formula is not null))
                cell.ShiftFormulaRows(asRange, numberOfRows);
        }
    }

    private static void ShiftRowHeights(XLRangeBase range, bool onlyUsedCells, int numberOfRows)
    {
        if (onlyUsedCells)
            return;

        var lastRow = range.Worksheet.Internals.CellsCollection.MaxRowUsed;
        if (lastRow <= 0)
            return;

        var firstRow = range.RangeAddress.FirstAddress.RowNumber;
        for (var ro = lastRow; ro >= firstRow; ro--)
        {
            var newRow = ro + numberOfRows;
            if (range.IsEntireRow())
                range.Worksheet.Row(newRow).Height = range.Worksheet.Row(ro).Height;
        }
    }

    private static void ApplyRowFormatting(XLRangeBase range, IXLRange rangeToReturn, bool formatFromAbove, XLCellsUsedOptions contentFlags)
    {
        if (formatFromAbove && rangeToReturn.RangeAddress.FirstAddress.RowNumber > 1)
            ApplyRowFormattingFromAbove(rangeToReturn, contentFlags);
        else
            ApplyRowFormattingFromExistingColumns(range, rangeToReturn, contentFlags);
    }

    private static void ApplyRowFormattingFromAbove(IXLRange rangeToReturn, XLCellsUsedOptions contentFlags)
    {
        var fr = rangeToReturn!.FirstRow()!;
        var model = fr.RowAbove()!;
        var modelFirstColumn = ((IXLRangeBase)model).FirstCellUsed(contentFlags);
        var modelLastColumn = ((IXLRangeBase)model).LastCellUsed(contentFlags);
        if (modelFirstColumn == null || modelLastColumn == null)
            return;

        var firstCoReturned = modelFirstColumn.Address.ColumnNumber
            - model.RangeAddress.FirstAddress.ColumnNumber + 1;
        var lastCoReturned = modelLastColumn.Address.ColumnNumber
            - model.RangeAddress.FirstAddress.ColumnNumber + 1;
        for (var co = firstCoReturned; co <= lastCoReturned; co++)
            rangeToReturn.Column(co).Style = model.Cell(co).Style;
    }

    private static void ApplyRowFormattingFromExistingColumns(XLRangeBase range, IXLRange rangeToReturn, XLCellsUsedOptions contentFlags)
    {
        var lastCoUsed = rangeToReturn.LastColumnUsed(contentFlags);
        if (lastCoUsed == null)
            return;

        var firstWsCol = rangeToReturn.RangeAddress.FirstAddress.ColumnNumber;
        var lastCoReturned = lastCoUsed.ColumnNumber() - firstWsCol + 1;
        for (var co = 1; co <= lastCoReturned; co++)
        {
            var wsCol = firstWsCol + co - 1;
            var styleToUse =
                range.Worksheet.Internals.ColumnsCollection.TryGetValue(wsCol, out var column)
                    ? column.Style
                    : range.Worksheet.Style;

            rangeToReturn.Column(co).Style = styleToUse;
        }
    }
}
