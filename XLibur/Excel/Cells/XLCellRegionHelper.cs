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
            hasRegionExpanded = false;

            var borderMinRow = Math.Max(minRow - 1, XLHelper.MinRowNumber);
            var borderMaxRow = Math.Min(maxRow + 1, XLHelper.MaxRowNumber);
            var borderMinColumn = Math.Max(minCol - 1, XLHelper.MinColumnNumber);
            var borderMaxColumn = Math.Min(maxCol + 1, XLHelper.MaxColumnNumber);

            if (minCol > XLHelper.MinColumnNumber &&
                !IsVerticalBorderBlank(sheet, borderMinColumn, borderMinRow, borderMaxRow))
            {
                hasRegionExpanded = true;
                minCol = borderMinColumn;
            }

            if (maxCol < XLHelper.MaxColumnNumber &&
                !IsVerticalBorderBlank(sheet, borderMaxColumn, borderMinRow, borderMaxRow))
            {
                hasRegionExpanded = true;
                maxCol = borderMaxColumn;
            }

            if (minRow > XLHelper.MinRowNumber &&
                !IsHorizontalBorderBlank(sheet, borderMinRow, borderMinColumn, borderMaxColumn))
            {
                hasRegionExpanded = true;
                minRow = borderMinRow;
            }

            if (maxRow < XLHelper.MaxRowNumber &&
                !IsHorizontalBorderBlank(sheet, borderMaxRow, borderMinColumn, borderMaxColumn))
            {
                hasRegionExpanded = true;
                maxRow = borderMaxRow;
            }
        } while (hasRegionExpanded);

        return new XLRangeAddress(
            new XLAddress(sheet, minRow, minCol, false, false),
            new XLAddress(sheet, maxRow, maxCol, false, false));
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
