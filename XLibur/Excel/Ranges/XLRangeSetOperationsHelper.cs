using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel;

/// <summary>
/// Contains set-algebra operations on ranges: Grow, Shrink, Intersection, Union, Difference,
/// and SurroundingCells.
/// <see cref="XLRangeBase"/> delegates to these methods to keep the main class smaller.
/// </summary>
internal static class XLRangeSetOperationsHelper
{
    internal static IXLRangeBase Grow(XLRangeBase range, int growCount)
    {
        var firstRow = Math.Max(1, range.RangeAddress.FirstAddress.RowNumber - growCount);
        var firstColumn = Math.Max(1, range.RangeAddress.FirstAddress.ColumnNumber - growCount);

        var lastRow = Math.Min(XLHelper.MaxRowNumber, range.RangeAddress.LastAddress.RowNumber + growCount);
        var lastColumn = Math.Min(XLHelper.MaxColumnNumber, range.RangeAddress.LastAddress.ColumnNumber + growCount);

        return range.Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
    }

    internal static IXLRangeBase? Shrink(XLRangeBase range, int shrinkCount)
    {
        var firstRow = range.RangeAddress.FirstAddress.RowNumber + shrinkCount;
        var firstColumn = range.RangeAddress.FirstAddress.ColumnNumber + shrinkCount;

        var lastRow = range.RangeAddress.LastAddress.RowNumber - shrinkCount;
        var lastColumn = range.RangeAddress.LastAddress.ColumnNumber - shrinkCount;

        if (firstRow > lastRow || firstColumn > lastColumn)
            return null;

        return range.Worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);
    }

    internal static IXLRangeAddress? Intersection(XLRangeBase range, IXLRangeBase otherRange, Func<IXLCell, bool>? thisRangePredicate, Func<IXLCell, bool>? otherRangePredicate)
    {
        if (otherRange == null)
            return null;

        if (!range.Worksheet.Equals(otherRange.Worksheet))
            return null;

        if (thisRangePredicate == null && otherRangePredicate == null)
        {
            // Special case, no predicates. We can optimise this a bit then.
            return range.RangeAddress.Intersection(otherRange.RangeAddress);
        }

        thisRangePredicate ??= _ => true;
        otherRangePredicate ??= _ => true;

        var intersectionCells = range.Cells(c => thisRangePredicate(c) && otherRange.Cells(otherRangePredicate).Contains(c));

        if (!intersectionCells.Any())
            return null;

        var firstRow = intersectionCells.Min(c => c.Address.RowNumber);
        var firstColumn = intersectionCells.Min(c => c.Address.ColumnNumber);

        var lastRow = intersectionCells.Max(c => c.Address.RowNumber);
        var lastColumn = intersectionCells.Max(c => c.Address.ColumnNumber);

        return new XLRangeAddress
        (
            new XLAddress(range.Worksheet, firstRow, firstColumn, fixedRow: false, fixedColumn: false),
            new XLAddress(range.Worksheet, lastRow, lastColumn, fixedRow: false, fixedColumn: false)
        );
    }

    internal static IXLCells SurroundingCells(XLRangeBase range, Func<IXLCell, bool>? predicate)
    {
        var cells = new XLCells(false, XLCellsUsedOptions.AllContents, predicate);
        range.Grow().Cells(c => !range.Contains(c)).ForEach(c => cells.Add((XLCell)c));
        return cells;
    }

    internal static IXLCells Union(XLRangeBase range, IXLRangeBase otherRange, Func<IXLCell, bool>? thisRangePredicate, Func<IXLCell, bool>? otherRangePredicate)
    {
        if (otherRange == null)
            return range.Cells(thisRangePredicate!);

        var cells = new XLCells(false, XLCellsUsedOptions.AllContents);
        if (!range.Worksheet.Equals(otherRange.Worksheet))
            return cells;

        thisRangePredicate ??= c => true;
        otherRangePredicate ??= c => true;

        range.Cells(thisRangePredicate).Concat(otherRange.Cells(otherRangePredicate)).Distinct().ForEach(c => cells.Add((XLCell)c));
        return cells;
    }

    internal static IXLCells Difference(XLRangeBase range, IXLRangeBase otherRange, Func<IXLCell, bool>? thisRangePredicate, Func<IXLCell, bool>? otherRangePredicate)
    {
        if (otherRange == null)
            return range.Cells(thisRangePredicate!);

        var cells = new XLCells(false, XLCellsUsedOptions.AllContents);
        if (!range.Worksheet.Equals(otherRange.Worksheet))
            return cells;

        thisRangePredicate ??= _ => true;
        otherRangePredicate ??= _ => true;

        range.Cells(c => thisRangePredicate(c) && !otherRange.Cells(otherRangePredicate).Contains(c)).ForEach(c => cells.Add((XLCell)c));
        return cells;
    }
}
