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

        foreach (var ws in range.Worksheet.Workbook.WorksheetsInternal)
        {
            foreach (var cell in ws.Internals.CellsCollection.GetCells(c => c.Formula is not null))
                cell.ShiftFormulaColumns(range.AsRange(), numberOfColumns);
        }

        range.Worksheet.SparklineGroupsInternal.ShiftColumns(XLSheetRange.FromRangeAddress(range.RangeAddress), numberOfColumns);

        // Inserting and shifting of whole columns is rather inconsistent across the codebase. In some places, the columns collection
        // is shifted before this method is called and thus the we can't shift column properties again. In others, the code relies on
        // shifting in this method.
        if (!onlyUsedCells)
        {
            var lastColumn = range.Worksheet.Internals.CellsCollection.MaxColumnUsed;
            if (lastColumn > 0)
            {
                var firstColumn = range.RangeAddress.FirstAddress.ColumnNumber;
                for (var co = lastColumn; co >= firstColumn; co--)
                {
                    var newColumn = co + numberOfColumns;
                    if (range.IsEntireColumn())
                    {
                        range.Worksheet.Column(newColumn).Width = range.Worksheet.Column(co).Width;
                    }
                }
            }
        }

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

        // We deliberately ignore conditional formats and data validation here. Their shifting is handled elsewhere
        var contentFlags = XLCellsUsedOptions.All
                           & ~XLCellsUsedOptions.ConditionalFormats
                           & ~XLCellsUsedOptions.DataValidation;

        if (formatFromLeft && rangeToReturn.RangeAddress.FirstAddress.ColumnNumber > 1)
        {
            var firstColumnUsed = rangeToReturn!.FirstColumn()!;
            var model = firstColumnUsed.ColumnLeft()!;
            var modelFirstRow = ((IXLRangeBase)model).FirstCellUsed(contentFlags);
            var modelLastRow = ((IXLRangeBase)model).LastCellUsed(contentFlags);
            if (modelFirstRow != null && modelLastRow != null)
            {
                var firstRoReturned = modelFirstRow.Address.RowNumber
                    - model.RangeAddress.FirstAddress.RowNumber + 1;
                var lastRoReturned = modelLastRow.Address.RowNumber
                    - model.RangeAddress.FirstAddress.RowNumber + 1;
                for (var ro = firstRoReturned; ro <= lastRoReturned; ro++)
                {
                    rangeToReturn.Row(ro).Style = model.Cell(ro).Style;
                }
            }
        }
        else
        {
            var lastRoUsed = rangeToReturn.LastRowUsed(contentFlags);
            if (lastRoUsed != null)
            {
                var lastRoReturned = lastRoUsed.RowNumber();
                for (var ro = 1; ro <= lastRoReturned; ro++)
                {
                    var styleToUse =
                        range.Worksheet.Internals.RowsCollection.TryGetValue(ro, out var row)
                            ? row.Style
                            : range.Worksheet.Style;

                    rangeToReturn.Row(ro).Style = styleToUse;
                }
            }
        }

        if (nullReturn)
            return null;

        return rangeToReturn.Columns();
    }

    internal static IXLRangeRows? InsertRowsAbove(XLRangeBase range, bool onlyUsedCells, int numberOfRows, bool formatFromAbove, bool nullReturn)
    {
        if (numberOfRows <= 0 || numberOfRows > XLHelper.MaxRowNumber)
            throw new System.ArgumentOutOfRangeException(nameof(numberOfRows),
                $"Number of rows to insert must be a positive number no more than {XLHelper.MaxRowNumber}");

        var asRange = range.AsRange();
        foreach (var ws in range.Worksheet.Workbook.WorksheetsInternal)
        {
            foreach (var cell in ws.Internals.CellsCollection.GetCells(c => c.Formula is not null))
                cell.ShiftFormulaRows(asRange, numberOfRows);
        }

        range.Worksheet.SparklineGroupsInternal.ShiftRows(XLSheetRange.FromRangeAddress(range.RangeAddress), numberOfRows);

        if (!onlyUsedCells)
        {
            var lastRow = range.Worksheet.Internals.CellsCollection.MaxRowUsed;
            if (lastRow > 0)
            {
                var firstRow = range.RangeAddress.FirstAddress.RowNumber;
                for (var ro = lastRow; ro >= firstRow; ro--)
                {
                    var newRow = ro + numberOfRows;
                    if (range.IsEntireRow())
                    {
                        range.Worksheet.Row(newRow).Height = range.Worksheet.Row(ro).Height;
                    }
                }
            }
        }

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

        // We deliberately ignore conditional formats and data validation here. Their shifting is handled elsewhere
        var contentFlags = XLCellsUsedOptions.All
                           & ~XLCellsUsedOptions.ConditionalFormats
                           & ~XLCellsUsedOptions.DataValidation;

        if (formatFromAbove && rangeToReturn.RangeAddress.FirstAddress.RowNumber > 1)
        {
            var fr = rangeToReturn!.FirstRow()!;
            var model = fr.RowAbove()!;
            var modelFirstColumn = ((IXLRangeBase)model).FirstCellUsed(contentFlags);
            var modelLastColumn = ((IXLRangeBase)model).LastCellUsed(contentFlags);
            if (modelFirstColumn != null && modelLastColumn != null)
            {
                var firstCoReturned = modelFirstColumn.Address.ColumnNumber
                    - model.RangeAddress.FirstAddress.ColumnNumber + 1;
                var lastCoReturned = modelLastColumn.Address.ColumnNumber
                    - model.RangeAddress.FirstAddress.ColumnNumber + 1;
                for (var co = firstCoReturned; co <= lastCoReturned; co++)
                {
                    rangeToReturn.Column(co).Style = model.Cell(co).Style;
                }
            }
        }
        else
        {
            var lastCoUsed = rangeToReturn.LastColumnUsed(contentFlags);
            if (lastCoUsed != null)
            {
                var lastCoReturned = lastCoUsed.ColumnNumber();
                for (var co = 1; co <= lastCoReturned; co++)
                {
                    var styleToUse =
                        range.Worksheet.Internals.ColumnsCollection.TryGetValue(co, out var column)
                            ? column.Style
                            : range.Worksheet.Style;

                    rangeToReturn.Style = styleToUse;
                }
            }
        }

        // Skip calling .Rows() for performance reasons if required.
        if (nullReturn)
            return null;

        return rangeToReturn.Rows();
    }
}
