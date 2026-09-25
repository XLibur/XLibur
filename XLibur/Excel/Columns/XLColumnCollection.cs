using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>The materialised columns of a worksheet, keyed by column number.</summary>
internal sealed class XLColumnsCollection : XLLineCollection<XLColumn, ColumnAxis>
{
    /// <inheritdoc cref="XLLineCollection{TLine, TAxis}.ShiftLines"/>
    public void ShiftColumnsRight(int startingColumn, int columnsToShift) => ShiftLines(startingColumn, columnsToShift);
}
