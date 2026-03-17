using System;

namespace XLibur.Excel.Coordinates;

/// <summary>
/// An offset of a cell in a sheet.
/// </summary>
/// <param name="RowOfs">The row offset in the number of rows from the original point.</param>
/// <param name="ColOfs">The column offset in the number of columns from the original point</param>
internal readonly record struct XLSheetOffset(int RowOfs, int ColOfs) : IComparable<XLSheetOffset>
{
    public int CompareTo(XLSheetOffset other)
    {
        var rowComparison = RowOfs.CompareTo(other.RowOfs);
        return rowComparison != 0 ? rowComparison : ColOfs.CompareTo(other.ColOfs);
    }
}