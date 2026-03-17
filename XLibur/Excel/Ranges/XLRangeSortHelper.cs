using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// Contains the sort algorithm logic for ranges.
/// <see cref="XLRangeBase"/> delegates private sort helpers here.
/// </summary>
internal static class XLRangeSortHelper
{
    internal static void SortRangeRows(XLRangeBase range, IXLSortElements sortColumns)
    {
        var sortRange = range.SheetRange;
        var cellsCollection = range.Worksheet.Internals.CellsCollection;
        if (sortRange.IsEntireColumn())
        {
            // If we're dealing with the entire column, we're not interested in the unused cells
            var lastRowUsed = cellsCollection.LastRowUsed(XLSheetRange.Full, XLCellsUsedOptions.Contents);
            sortRange = new XLSheetRange(sortRange.FirstPoint, new XLSheetPoint(lastRowUsed, sortRange.RightColumn));
        }

        var comparer = new XLRangeRowsSortComparer(range.Worksheet, sortRange, sortColumns);
        var rows = new int[sortRange.Height];
        for (var i = 0; i < sortRange.Height; ++i)
            rows[i] = i + sortRange.TopRow;

        System.Array.Sort(rows, comparer);

        cellsCollection.RemapRows(rows, sortRange);
    }

    internal static void SortRangeColumns(XLRangeBase range, IXLSortElements sortRows)
    {
        var sortRange = range.SheetRange;
        var cellsCollection = range.Worksheet.Internals.CellsCollection;
        if (sortRange.IsEntireRow())
        {
            // If we're dealing with the entire row, we're not interested in the unused cells
            var lastColumnCell = cellsCollection.LastColumnUsed(XLSheetRange.Full, XLCellsUsedOptions.Contents);
            sortRange = new XLSheetRange(sortRange.FirstPoint, new XLSheetPoint(sortRange.BottomRow, lastColumnCell));
        }

        var comparer = new XLRangeColumnsSortComparer(range.Worksheet, sortRange, sortRows);
        var columns = new int[sortRange.Width];
        for (var i = 0; i < sortRange.Width; ++i)
            columns[i] = i + sortRange.LeftColumn;

        System.Array.Sort(columns, comparer);

        cellsCollection.RemapColumns(columns, sortRange);
    }

    internal static IEnumerable<XLSortElement> ParseSortOrder(string columnsToSortBy, XLSortOrder defaultSortOrder, bool matchCase, bool ignoreBlanks)
    {
        foreach (var sortColumn in columnsToSortBy.Split(',').Select(coPair => coPair.Trim()))
        {
            var sortOrder = defaultSortOrder;

            string columnName;
            if (sortColumn.Contains(' '))
            {
                var pair = sortColumn.Split(' ');
                columnName = pair[0];
                sortOrder = pair[1].Equals("ASC", System.StringComparison.OrdinalIgnoreCase) ? XLSortOrder.Ascending : XLSortOrder.Descending;
            }
            else
            {
                columnName = sortColumn;
            }

            if (!int.TryParse(columnName, out var columnNumber))
                columnNumber = XLHelper.GetColumnNumberFromLetter(columnName);

            yield return new XLSortElement(
                columnNumber,
                sortOrder,
                ignoreBlanks,
                matchCase);
        }
    }

    internal static string DefaultSortString(XLRangeBase range)
    {
        var sb = new StringBuilder();
        var maxColumn = range.ColumnCount();
        if (maxColumn == XLHelper.MaxColumnNumber)
            maxColumn = ((IXLRangeBase)range).LastCellUsed(XLCellsUsedOptions.All)!.Address.ColumnNumber;
        for (var i = 1; i <= maxColumn; i++)
        {
            if (sb.Length > 0)
                sb.Append(',');

            sb.Append(i);
        }

        return sb.ToString();
    }
}
