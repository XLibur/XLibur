using System.Collections.Generic;
using XLibur.Excel;

namespace XLibur.Report.Ranges;

/// <summary>
/// A defined name that resolved to a collection, and so marks a block of template rows to be
/// repeated once per item.
/// </summary>
internal sealed class BoundRange
{
    public BoundRange(IXLDefinedName definedName, IXLWorksheet worksheet, IReadOnlyList<object?> items, RangeArea area)
    {
        DefinedName = definedName;
        Worksheet = worksheet;
        Items = items;
        Area = area;
    }

    /// <summary>The defined name marking the template rows.</summary>
    public IXLDefinedName DefinedName { get; }

    /// <summary>The sheet the name points at.</summary>
    public IXLWorksheet Worksheet { get; }

    /// <summary>The data bound to it, materialised because expansion enumerates it more than once.</summary>
    public IReadOnlyList<object?> Items { get; }

    /// <summary>
    /// Where the range sat when it was resolved. Used to order expansion and to keep the
    /// pre-expansion cell pass out of the range; the live address is re-read at expansion time,
    /// because expanding an earlier range shifts this one down.
    /// </summary>
    public RangeArea Area { get; }

    public string Name => DefinedName.Name;
}

/// <summary>A rectangle of cells on a sheet, as row and column numbers.</summary>
internal readonly struct RangeArea
{
    public RangeArea(int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        FirstRow = firstRow;
        FirstColumn = firstColumn;
        LastRow = lastRow;
        LastColumn = lastColumn;
    }

    public int FirstRow { get; }

    public int FirstColumn { get; }

    public int LastRow { get; }

    public int LastColumn { get; }

    public int RowCount => LastRow - FirstRow + 1;

    public int ColumnCount => LastColumn - FirstColumn + 1;

    public bool Contains(int row, int column) =>
        row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;

    public static RangeArea From(IXLRangeBase range)
    {
        var address = range.RangeAddress;
        return new RangeArea(
            address.FirstAddress.RowNumber,
            address.FirstAddress.ColumnNumber,
            address.LastAddress.RowNumber,
            address.LastAddress.ColumnNumber);
    }
}
