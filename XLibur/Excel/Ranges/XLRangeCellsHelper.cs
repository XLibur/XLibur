using System;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel;

/// <summary>
/// Contains the algorithm bodies for finding used cells and the internal cell-enumeration helper.
/// <see cref="XLRangeBase"/> delegates <c>FirstCellUsed</c>, <c>LastCellUsed</c>, and
/// <c>CellsUsedInternal</c> here.
/// </summary>
internal static class XLRangeCellsHelper
{
    internal static XLCell? FirstCellUsed(XLRangeBase range, XLCellsUsedOptions options, Func<IXLCell, bool>? predicate)
    {
        var cellsUsed = CellsUsedInternal(range, options, r => r.FirstCell(), predicate).ToList();

        if (!cellsUsed.Any())
            return null;

        var firstRow = cellsUsed.Min(c => c.Address.RowNumber);
        var firstColumn = cellsUsed.Min(c => c.Address.ColumnNumber);

        if (firstRow < range.RangeAddress.FirstAddress.RowNumber)
            firstRow = range.RangeAddress.FirstAddress.RowNumber;

        if (firstColumn < range.RangeAddress.FirstAddress.ColumnNumber)
            firstColumn = range.RangeAddress.FirstAddress.ColumnNumber;

        return range.Worksheet.Cell(firstRow, firstColumn);
    }

    internal static XLCell? LastCellUsed(XLRangeBase range, XLCellsUsedOptions options, Func<IXLCell, bool>? predicate)
    {
        var cellsUsed = CellsUsedInternal(range, options, r => r.LastCell(), predicate).ToList();

        if (!cellsUsed.Any())
            return null;

        var lastRow = cellsUsed.Max(c => c.Address.RowNumber);
        var lastColumn = cellsUsed.Max(c => c.Address.ColumnNumber);

        if (lastRow > range.RangeAddress.LastAddress.RowNumber)
            lastRow = range.RangeAddress.LastAddress.RowNumber;

        if (lastColumn > range.RangeAddress.LastAddress.ColumnNumber)
            lastColumn = range.RangeAddress.LastAddress.ColumnNumber;

        return range.Worksheet.Cell(lastRow, lastColumn);
    }

    internal static IEnumerable<IXLCell> CellsUsedInternal(XLRangeBase range, XLCellsUsedOptions options, Func<IXLRange, IXLCell> selector, Func<IXLCell, bool>? predicate)
    {
        predicate ??= (t => true);

        //To avoid unnecessary initialization of thousands of cells
        var opt = options
                  & ~XLCellsUsedOptions.ConditionalFormats
                  & ~XLCellsUsedOptions.DataValidation
                  & ~XLCellsUsedOptions.MergedRanges;

        // If opt == 0 then we're basically back at unconstrained, so just set back the original options
        if (opt == XLCellsUsedOptions.NoConstraints)
            opt = options;

        IEnumerable<IXLCell> cellsUsed = range.CellsUsed(opt, predicate);

        if (options.HasFlag(XLCellsUsedOptions.ConditionalFormats))
        {
            cellsUsed = cellsUsed.Union(
                range.Worksheet.ConditionalFormats
                    .SelectMany(cf => cf.Ranges.GetIntersectedRanges(range.RangeAddress))
                    .Select(selector)
                    .Where(predicate)
            );
        }
        if (options.HasFlag(XLCellsUsedOptions.DataValidation))
        {
            cellsUsed = cellsUsed.Union(
                range.Worksheet.DataValidations
                    .GetAllInRange(range.RangeAddress)
                    .SelectMany(dv => dv.Ranges)
                    .Select(selector)
                    .Where(predicate)
            );
        }
        if (options.HasFlag(XLCellsUsedOptions.MergedRanges))
        {
            cellsUsed = cellsUsed.Union(
                range.Worksheet.MergedRanges.GetIntersectedRanges(range.RangeAddress)
                    .Select(selector)
                    .Where(predicate)
            );
        }

        return cellsUsed;
    }
}
