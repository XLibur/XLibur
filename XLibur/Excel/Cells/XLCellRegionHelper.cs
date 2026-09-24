using System;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

internal static class XLCellRegionHelper
{
    internal static IXLRangeAddress FindCurrentRegion(XLWorksheet sheet, int rowNumber, int columnNumber)
    {
        var minRow = rowNumber;
        var minCol = columnNumber;
        var maxRow = rowNumber;
        var maxCol = columnNumber;

        bool hasRegionExpanded;

        do
        {
            var borderMinRow = Math.Max(minRow - 1, XLHelper.MinRowNumber);
            var borderMaxRow = Math.Min(maxRow + 1, XLHelper.MaxRowNumber);
            var borderMinColumn = Math.Max(minCol - 1, XLHelper.MinColumnNumber);
            var borderMaxColumn = Math.Min(maxCol + 1, XLHelper.MaxColumnNumber);

            // Both passes test the borders of the region as it was at the start of this iteration.
            var hasColumnsExpanded = TryExpandColumns(sheet, ref minCol, ref maxCol, borderMinRow, borderMaxRow);
            var hasRowsExpanded = TryExpandRows(sheet, ref minRow, ref maxRow, borderMinColumn, borderMaxColumn);
            hasRegionExpanded = hasColumnsExpanded || hasRowsExpanded;
        } while (hasRegionExpanded);

        return new XLRangeAddress(
            new XLAddress(sheet, minRow, minCol, false, false),
            new XLAddress(sheet, maxRow, maxCol, false, false));
    }

    /// <summary>
    /// Extends the region by one column to the left and/or right when the neighbouring column
    /// (between <paramref name="borderMinRow"/> and <paramref name="borderMaxRow"/>) holds a non-empty cell.
    /// </summary>
    private static bool TryExpandColumns(XLWorksheet sheet, ref int minCol, ref int maxCol, int borderMinRow, int borderMaxRow)
    {
        var hasExpanded = false;
        var borderMinColumn = Math.Max(minCol - 1, XLHelper.MinColumnNumber);
        var borderMaxColumn = Math.Min(maxCol + 1, XLHelper.MaxColumnNumber);

        if (minCol > XLHelper.MinColumnNumber &&
            !IsVerticalBorderBlank(sheet, borderMinColumn, borderMinRow, borderMaxRow))
        {
            hasExpanded = true;
            minCol = borderMinColumn;
        }

        if (maxCol < XLHelper.MaxColumnNumber &&
            !IsVerticalBorderBlank(sheet, borderMaxColumn, borderMinRow, borderMaxRow))
        {
            hasExpanded = true;
            maxCol = borderMaxColumn;
        }

        return hasExpanded;
    }

    /// <summary>
    /// Extends the region by one row up and/or down when the neighbouring row (between
    /// <paramref name="borderMinColumn"/> and <paramref name="borderMaxColumn"/>) holds a non-empty cell.
    /// </summary>
    private static bool TryExpandRows(XLWorksheet sheet, ref int minRow, ref int maxRow, int borderMinColumn, int borderMaxColumn)
    {
        var hasExpanded = false;
        var borderMinRow = Math.Max(minRow - 1, XLHelper.MinRowNumber);
        var borderMaxRow = Math.Min(maxRow + 1, XLHelper.MaxRowNumber);

        if (minRow > XLHelper.MinRowNumber &&
            !IsHorizontalBorderBlank(sheet, borderMinRow, borderMinColumn, borderMaxColumn))
        {
            hasExpanded = true;
            minRow = borderMinRow;
        }

        if (maxRow < XLHelper.MaxRowNumber &&
            !IsHorizontalBorderBlank(sheet, borderMaxRow, borderMinColumn, borderMaxColumn))
        {
            hasExpanded = true;
            maxRow = borderMaxRow;
        }

        return hasExpanded;
    }

    private static bool IsVerticalBorderBlank(XLWorksheet sheet, int borderColumn, int borderMinRow, int borderMaxRow)
    {
        for (var row = borderMinRow; row <= borderMaxRow; row++)
        {
            var verticalBorderCell = sheet.Cell(row, borderColumn);
            if (!verticalBorderCell.IsEmpty(XLCellsUsedOptions.AllContents))
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsHorizontalBorderBlank(XLWorksheet sheet, int borderRow, int borderMinColumn, int borderMaxColumn)
    {
        for (var col = borderMinColumn; col <= borderMaxColumn; col++)
        {
            var horizontalBorderCell = sheet.Cell(borderRow, col);
            if (!horizontalBorderCell.IsEmpty(XLCellsUsedOptions.AllContents))
            {
                return false;
            }
        }

        return true;
    }
}
