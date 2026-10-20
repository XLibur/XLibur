using System;
using System.Drawing;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Graphics;

namespace XLibur.Tests.Excel.Columns;

/// <summary>
/// Autofit pads the width of the text the way Excel does (#789). These tests take the text
/// measurement out of the picture: every glyph is 4 px wide, as <c>i</c> is in Calibri 11 at
/// 120 DPI, so only the padding decides the result.
/// </summary>
/// <remarks>
/// <para>
/// Measured in Excel (Microsoft 365, 16.0 build 20430, 120 DPI, over COM): a run of <c>i</c> in
/// Calibri 11, <c>Columns(n).AutoFit()</c>, then <c>ColumnWidth</c> and <c>Width</c>. Excel measures
/// text with GDI, at an em of whole pixels and with hinted advance widths, and the GDI width of every
/// string in the survey plus the padding below gave Excel's width to the pixel, for 108 strings in
/// nine fonts and sizes and for runs of 1 to 120 glyphs. For text <c>W</c> px wide the cell is
/// </para>
/// <code>
/// W + 2 * ceil(cellMdw / 4) + 1 + max(2, floor(W / (4 * mdw)))
/// </code>
/// <para>
/// where <c>cellMdw</c> is the digit width of the cell's font and <c>mdw</c> that of the column, the
/// first font's. So Calibri 11 text is padded by 9 px up to 107 px, and by one more pixel for every
/// four digits of width after that. The last term follows the column's digit: with a Normal style of
/// Calibri 20 (17 px digits) the same text gains a pixel every 68 px instead of every 36 px.
/// </para>
/// <para>
/// Excel wraps text by the same rule: the longest run of <c>i</c> it keeps on one line is the
/// longest whose padded width fits the column.
/// </para>
/// </remarks>
public class AutofitPaddingTests
{
    private const int GlyphWidth = 4;

    [Test]
    // 4 px of text and 9 px of padding is 13 px, less than one character of 9 px plus 7 px.
    [Arguments(1, 0.81)]
    [Arguments(17, 7.78)]
    [Arguments(18, 8.22)]
    // 104 px is padded by 9 px, 108 px by 10 px.
    [Arguments(26, 11.78)]
    [Arguments(27, 12.33)]
    [Arguments(35, 15.89)]
    [Arguments(36, 16.44)]
    [Arguments(70, 31.89)]
    [Arguments(71, 32.33)]
    [Arguments(100, 45.67)]
    // 480 px is padded by 7 + 13 px: 500 px.
    [Arguments(120, 54.78)]
    public async Task Column_autofit_pads_the_text_as_Excel_does(int glyphs, double excelWidth)
    {
        using var wb = Workbook(firstFontMdw: 9);
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = new string('i', glyphs);

        ws.Column(1).AdjustToContents();

        // Excel shows a width to two decimals.
        await Assert.That(ws.Column(1).Width).IsEqualTo(excelWidth).Within(0.005);
    }

    [Test]
    // The column is in 17 px digits, so the text gains a pixel every 68 px from 204 px: 50 glyphs
    // (200 px) are 209 px wide, 51 glyphs (204 px) 214 px.
    [Arguments(1, 0.46)]
    [Arguments(50, 11.65)]
    [Arguments(51, 11.94)]
    [Arguments(68, 16)]
    [Arguments(69, 16.24)]
    [Arguments(100, 23.59)]
    [Arguments(120, 28.41)]
    public async Task Column_autofit_pads_long_text_in_digits_of_the_column(int glyphs, double excelWidth)
    {
        // The first font has 17 px digits, as Calibri 20 has in Excel, and the cell's font 9 px.
        using var wb = Workbook(firstFontMdw: 17);
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = new string('i', glyphs);
        ws.Cell("A1").Style.Font.FontSize = CellFontSize;

        ws.Column(1).AdjustToContents();

        await Assert.That(ws.Column(1).Width).IsEqualTo(excelWidth).Within(0.005);
    }

    [Test]
    // A column of 8 characters is 79 px, of 16 characters 151 px, and so on.
    [Arguments(8, 17)]
    [Arguments(16, 35)]
    [Arguments(32, 70)]
    [Arguments(48, 105)]
    public async Task Row_autofit_wraps_where_Excel_does(double columnWidth, int excelLongestLine)
    {
        var oneLine = WrappedHeight(columnWidth, 1);

        await Assert.That(WrappedHeight(columnWidth, excelLongestLine)).IsEqualTo(oneLine);
        await Assert.That(WrappedHeight(columnWidth, excelLongestLine + 1)).IsEqualTo(2 * oneLine).Within(0.01);
    }

    private static double WrappedHeight(double columnWidth, int glyphs)
    {
        using var wb = Workbook(firstFontMdw: 9);
        var ws = wb.AddWorksheet();
        ws.Column(1).Width = columnWidth;
        var cell = ws.Cell("A1");
        cell.Value = new string('i', glyphs);
        cell.Style.Alignment.WrapText = true;

        ws.Row(1).AdjustToContents();

        return ws.Row(1).Height;
    }

    private const double CellFontSize = 10;

    private static XLWorkbook Workbook(int firstFontMdw)
    {
        using var probe = new XLWorkbook();
        return new XLWorkbook(new LoadOptions
        {
            Dpi = new Point(120, 120),
            FontEngine = new FixedGlyphFontEngine(probe.FontEngine, firstFontMdw),
        });
    }

    /// <summary>
    /// The test font engine, except that every glyph is <see cref="GlyphWidth"/> px wide, and digits
    /// are <paramref name="firstFontMdw"/> px wide in the workbook's default font (the first font of
    /// the stylesheet) and 9 px in a font of <see cref="CellFontSize"/>.
    /// </summary>
    private sealed class FixedGlyphFontEngine(IXLFontEngine inner, int firstFontMdw) : IXLFontEngine
    {
        public double GetTextHeight(IXLFontBase font, double dpiY) => inner.GetTextHeight(font, dpiY);

        public double GetTextWidth(string text, IXLFontBase font, double dpiX) => text.Length * GlyphWidth;

        public double GetMaxDigitWidth(IXLFontBase font, double dpiX) =>
            Math.Abs(font.FontSize - CellFontSize) < 0.01 ? 9 : firstFontMdw;

        public double GetDescent(IXLFontBase font, double dpiY) => inner.GetDescent(font, dpiY);

        public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi)
        {
            var box = inner.GetGlyphBox(graphemeCluster, font, dpi);
            return new GlyphBox(GlyphWidth, box.EmSize, box.Descent);
        }
    }
}
