using System.Drawing;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Columns;

/// <summary>
/// Load options at a DPI, for the column width tests: a column's width in pixels, and its padding
/// in digits, depend on the digit width, and with it on the DPI.
/// </summary>
internal static class ColumnWidthDpi
{
    /// <summary>
    /// 120 DPI, the DPI of the machine the Excel measurements in these tests were taken on. Calibri
    /// 11 has a 9 px digit there.
    /// </summary>
    internal static readonly LoadOptions AtExcelDpi = At(120);

    /// <summary>Options that load and create a workbook at <paramref name="dpi"/> both ways.</summary>
    internal static LoadOptions At(int dpi) => new() { Dpi = new Point(dpi, dpi) };
}
