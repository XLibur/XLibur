using XLibur.Excel.Coordinates;

namespace XLibur.Excel.Rows;

/// <summary>The materialised rows of a worksheet, keyed by row number.</summary>
internal sealed class XLRowsCollection : XLLineCollection<XLRow, RowAxis>
{
    /// <inheritdoc cref="XLLineCollection{TLine, TAxis}.ShiftLines"/>
    public void ShiftRowsDown(int startingRow, int rowsToShift) => ShiftLines(startingRow, rowsToShift);
}
