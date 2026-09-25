using System;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Graphics;

namespace XLibur.Tests.Excel.Columns;

/// <summary>
/// Maximum digit width (MDW) is rounded to whole pixels in one place, <c>XLHelper.GetMdw</c>, and
/// that rounding is half away from zero (#621 finding 17). Column autofit and the wrap width used
/// by row autofit used banker's rounding before, so an MDW of exactly 8.5 px was 8 there but 9 in
/// the load path and in the per-cell padding.
/// </summary>
public class MaxDigitWidthRoundingTests
{
    [Test]
    [Arguments(7.43359375, 7)] // Calibri 11pt at 96 DPI
    [Arguments(7.5, 8)]
    [Arguments(8.5, 9)]
    [Arguments(8.49, 8)]
    public async Task GetMdw_RoundsHalfAwayFromZero(double engineMdw, int expected)
    {
        using var probe = new XLWorkbook();
        var engine = new FixedMdwFontEngine(probe.FontEngine, engineMdw);

        var mdw = XLHelper.GetMdw(engine, probe.Style.Font, 96);

        await Assert.That(mdw).IsEqualTo(expected);
    }

    [Test]
    public async Task ColumnAdjustToContents_UsesTheSameMdwAsTheRestOfTheLibrary()
    {
        // With an MDW of 8.5 px the column is sized in 9 px characters: a minimum width of 10.3
        // characters is 10.3 * 9 + 7 px padding = 99.7 px, rounded up to 100 px = 93/9 characters.
        // Banker's rounding made it 8 px, and the width 10.375.
        using var probe = new XLWorkbook();
        using var wb = new XLWorkbook(new LoadOptions
        {
            FontEngine = new FixedMdwFontEngine(probe.FontEngine, 8.5),
        });
        var ws = wb.AddWorksheet();

        ws.Column(1).AdjustToContents(10.3, 100);

        await Assert.That(ws.Column(1).Width).IsEqualTo(93d / 9).Within(1e-9);
    }

    [Test]
    public async Task RowAdjustToContents_WrapsAtTheSameMdwAsTheRestOfTheLibrary()
    {
        // Every glyph is 10 px wide, so a 7-letter word is 70 px, 81 px with the cell padding. A
        // default column (8.43 characters) is 83 px wide with a 9 px MDW and 72 px with an 8 px
        // one, so the word fits on its line only when 8.5 rounds to 9.
        var at8_5 = WrappedRowHeight(8.5);
        var at9 = WrappedRowHeight(9);
        var at8 = WrappedRowHeight(8);

        await Assert.That(at8).IsNotEqualTo(at9);
        await Assert.That(at8_5).IsEqualTo(at9);
    }

    private static double WrappedRowHeight(double engineMdw)
    {
        using var probe = new XLWorkbook();
        using var wb = new XLWorkbook(new LoadOptions
        {
            FontEngine = new FixedMdwFontEngine(probe.FontEngine, engineMdw, fixedGlyphWidth: 10),
        });
        var ws = wb.AddWorksheet();
        var cell = ws.Cell("A1");
        cell.Value = "aaaaaaa aaaaaaa";
        cell.Style.Alignment.WrapText = true;

        ws.Row(1).AdjustToContents();

        return ws.Row(1).Height;
    }

    /// <summary>
    /// The test font engine, except for a fixed maximum digit width and, optionally, a fixed
    /// glyph advance width.
    /// </summary>
    private sealed class FixedMdwFontEngine(IXLFontEngine inner, double mdw, float? fixedGlyphWidth = null)
        : IXLFontEngine
    {
        public double GetTextHeight(IXLFontBase font, double dpiY) => inner.GetTextHeight(font, dpiY);

        public double GetTextWidth(string text, IXLFontBase font, double dpiX) => inner.GetTextWidth(text, font, dpiX);

        public double GetMaxDigitWidth(IXLFontBase font, double dpiX) => mdw;

        public double GetDescent(IXLFontBase font, double dpiY) => inner.GetDescent(font, dpiY);

        public GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi)
        {
            var box = inner.GetGlyphBox(graphemeCluster, font, dpi);
            return fixedGlyphWidth is { } width ? new GlyphBox(width, box.EmSize, box.Descent) : box;
        }
    }
}
