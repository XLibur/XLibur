using System;
using System.IO;
using XLibur.Excel;
using XLibur.Graphics;
using System.Threading.Tasks;

namespace XLibur.Fonts.SixLabors.Tests;

public class SixLaborsFontEngineTests
{
    /// <summary>
    /// Stream-based engine using TestFontA as fallback — works on all platforms including CI (no system fonts needed).
    /// </summary>
    private static IXLFontEngine CreateTestEngine()
    {
        var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        return SixLaborsFontEngine.CreateOnlyWithFonts(fallbackStream);
    }

    private readonly IXLFontEngine _engine = CreateTestEngine();

    #region Text width

    [Test]
    public async Task GetTextWidth_ReturnsPositiveValue()
    {
        var font = new DummyFont("TestFontA", 20);
        var width = _engine.GetTextWidth("Lorem ipsum dolor sit amet", font, 96);

        await Assert.That(width).IsGreaterThan(0);
    }

    [Test]
    public async Task GetTextWidth_LongerTextIsWider()
    {
        var font = new DummyFont("TestFontA", 11);
        var shortWidth = _engine.GetTextWidth("AB", font, 96);
        var longWidth = _engine.GetTextWidth("ABCDEF", font, 96);

        await Assert.That(longWidth).IsGreaterThan(shortWidth);
    }

    [Test]
    public async Task GetTextWidth_LargerFontIsWider()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 20);
        var smallWidth = _engine.GetTextWidth("Test", smallFont, 96);
        var largeWidth = _engine.GetTextWidth("Test", largeFont, 96);

        await Assert.That(largeWidth).IsGreaterThan(smallWidth);
    }

    [Test]
    public async Task GetTextWidth_HigherDpiIsWider()
    {
        var font = new DummyFont("TestFontA", 11);
        var width96 = _engine.GetTextWidth("Test", font, 96);
        var width120 = _engine.GetTextWidth("Test", font, 120);

        await Assert.That(width120).IsGreaterThan(width96);
    }

    [Test]
    public async Task GetTextWidth_EmptyStringReturnsZero()
    {
        var font = new DummyFont("TestFontA", 11);
        var width = _engine.GetTextWidth("", font, 96);

        await Assert.That(width).IsEqualTo(0);
    }

    #endregion

    #region Text height

    [Test]
    public async Task GetTextHeight_ReturnsPositiveValue()
    {
        var font = new DummyFont("TestFontA", 11);
        var height = _engine.GetTextHeight(font, 96);

        await Assert.That(height).IsGreaterThan(0);
    }

    [Test]
    public async Task GetTextHeight_LargerFontIsTaller()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 30);
        var smallHeight = _engine.GetTextHeight(smallFont, 96);
        var largeHeight = _engine.GetTextHeight(largeFont, 96);

        await Assert.That(largeHeight).IsGreaterThan(smallHeight);
    }

    [Test]
    public async Task GetTextHeight_HigherDpiIsTaller()
    {
        var font = new DummyFont("TestFontA", 11);
        var height96 = _engine.GetTextHeight(font, 96);
        var height120 = _engine.GetTextHeight(font, 120);

        await Assert.That(height120).IsGreaterThan(height96);
    }

    #endregion

    #region Max digit width

    [Test]
    public async Task GetMaxDigitWidth_ReturnsPositiveValue()
    {
        var font = new DummyFont("TestFontA", 11);
        var mdw = _engine.GetMaxDigitWidth(font, 96);

        await Assert.That(mdw).IsGreaterThan(0);
    }

    [Test]
    public async Task GetMaxDigitWidth_LargerFontIsWider()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 20);
        var smallMdw = _engine.GetMaxDigitWidth(smallFont, 96);
        var largeMdw = _engine.GetMaxDigitWidth(largeFont, 96);

        await Assert.That(largeMdw).IsGreaterThan(smallMdw);
    }

    #endregion

    #region Descent

    [Test]
    public async Task GetDescent_ReturnsPositiveValue()
    {
        var font = new DummyFont("TestFontA", 11);
        var descent = _engine.GetDescent(font, 96);

        await Assert.That(descent).IsGreaterThan(0);
    }

    [Test]
    public async Task GetDescent_LargerFontHasLargerDescent()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 30);
        var smallDescent = _engine.GetDescent(smallFont, 96);
        var largeDescent = _engine.GetDescent(largeFont, 96);

        await Assert.That(largeDescent).IsGreaterThan(smallDescent);
    }

    #endregion

    #region Glyph box

    [Test]
    public async Task GetGlyphBox_ReturnsPositiveAdvanceWidth()
    {
        var font = new DummyFont("TestFontA", 11);
        Span<int> codePoints = ['A'];
        var box = _engine.GetGlyphBox(codePoints, font, new Dpi(96, 96));

        await Assert.That(box.AdvanceWidth).IsGreaterThan(0);
        await Assert.That(box.EmSize).IsGreaterThan(0);
    }

    [Test]
    public async Task GetGlyphBox_MultipleCharactersProduceValidWidths()
    {
        var font = new DummyFont("TestFontA", 11);
        Span<int> charA = ['A'];
        Span<int> charB = ['B'];

        var boxA = _engine.GetGlyphBox(charA, font, new Dpi(96, 96));
        var boxB = _engine.GetGlyphBox(charB, font, new Dpi(96, 96));

        await Assert.That(boxA.AdvanceWidth).IsGreaterThan(0);
        await Assert.That(boxB.AdvanceWidth).IsGreaterThan(0);
    }

    [Test]
    public async Task GetGlyphBox_DescentIsPositive()
    {
        var font = new DummyFont("TestFontA", 11);
        Span<int> codePoints = ['g'];
        var box = _engine.GetGlyphBox(codePoints, font, new Dpi(96, 96));

        await Assert.That(box.Descent).IsGreaterThanOrEqualTo(0);
    }

    [Test]
    public async Task GetGlyphBox_LargerFontProducesLargerBox()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 20);
        Span<int> codePoints = ['A'];

        var smallBox = _engine.GetGlyphBox(codePoints, smallFont, new Dpi(96, 96));
        var largeBox = _engine.GetGlyphBox(codePoints, largeFont, new Dpi(96, 96));

        await Assert.That(largeBox.AdvanceWidth).IsGreaterThan(smallBox.AdvanceWidth);
        await Assert.That(largeBox.EmSize).IsGreaterThan(smallBox.EmSize);
    }

    #endregion

    #region Fallback behavior

    [Test]
    public async Task NonExistentFont_UsesFallback()
    {
        // With stream-based engine, non-existent fonts fall back to the provided fallback font
        var nonExistent = new DummyFont("TotallyFakeNonExistentFont12345", 11);
        var fallback = new DummyFont("TestFontA", 11);

        var nonExistentWidth = _engine.GetTextWidth("Test", nonExistent, 96);
        var fallbackWidth = _engine.GetTextWidth("Test", fallback, 96);

        await Assert.That(nonExistentWidth).IsEqualTo(fallbackWidth);
    }

    [Test]
    public async Task NonExistentFont_UsesFallbackForHeight()
    {
        var nonExistent = new DummyFont("TotallyFakeNonExistentFont12345", 14);
        var fallback = new DummyFont("TestFontA", 14);

        var nonExistentHeight = _engine.GetTextHeight(nonExistent, 96);
        var fallbackHeight = _engine.GetTextHeight(fallback, 96);

        await Assert.That(nonExistentHeight).IsEqualTo(fallbackHeight);
    }

    #endregion

    #region Stream-based factory methods

    [Test]
    public async Task CreateOnlyWithFonts_UsesProvidedFallback()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var engine = SixLaborsFontEngine.CreateOnlyWithFonts(fallbackStream);

        var font = new DummyFont("Nonexistent Font", 20);
        var width = engine.GetTextWidth("A", font, 120);

        // TestFontA at 20pt, 120 DPI — v2 may have slightly different measurement than v1
        await Assert.That(width).IsEqualTo(31.25d).Within(1.0);
    }

    [Test]
    public async Task CreateOnlyWithFonts_CanLoadExtraFonts()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        using var fontBStream = TestHelper.GetStreamFromResource("Fonts.TestFontB.ttf");
        var engine = SixLaborsFontEngine.CreateOnlyWithFonts(fallbackStream, fontBStream);

        var widthB = engine.GetTextWidth("B", new DummyFont("TestFontB", 30), 96);

        await Assert.That(widthB).IsEqualTo(25d).Within(1.5);
    }

    [Test]
    public async Task CreateWithFontsAndSystemFonts_CanUseFallbackFont()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var engine = SixLaborsFontEngine.CreateWithFontsAndSystemFonts(fallbackStream);

        // Even if system fonts aren't available, the fallback font should work
        var font = new DummyFont("NonexistentFont", 11);
        var width = engine.GetTextWidth("Test", font, 96);

        await Assert.That(width).IsGreaterThan(0);
    }

    #endregion

    #region Workbook integration

    [Test]
    public async Task FontEngine_WorksWithWorkbookViaLoadOptions()
    {
        var loadOptions = new LoadOptions { FontEngine = _engine };
        using var wb = new XLWorkbook(loadOptions);
        var ws = wb.AddWorksheet();

        ws.Cell(1, 1).Value = "Hello World";
        ws.Column(1).AdjustToContents();

        await Assert.That(ws.Column(1).Width).IsGreaterThan(0);
    }

    [Test]
    public async Task FontEngine_AdjustToContents_ProducesReasonableWidth()
    {
        var loadOptions = new LoadOptions { FontEngine = _engine };
        using var wb = new XLWorkbook(loadOptions);
        var ws = wb.AddWorksheet();

        ws.Cell(1, 1).Value = "Short";
        ws.Cell(2, 1).Value = "A much longer text that should need more width";

        ws.Column(1).AdjustToContents();

        // Width should accommodate the longer text
        await Assert.That(ws.Column(1).Width).IsGreaterThan(8.43); // 8.43 is the default column width
    }

    [Test]
    public async Task FontEngine_AdjustRowHeight_ProducesReasonableHeight()
    {
        var loadOptions = new LoadOptions { FontEngine = _engine };
        using var wb = new XLWorkbook(loadOptions);
        var ws = wb.AddWorksheet();

        ws.Cell(1, 1).Value = "Test";
        ws.Row(1).AdjustToContents();

        await Assert.That(ws.Row(1).Height).IsGreaterThan(0);
    }

    [Test]
    public async Task FontEngine_CanSaveAndReloadWorkbook()
    {
        var loadOptions = new LoadOptions { FontEngine = _engine };
        using var wb = new XLWorkbook(loadOptions);
        var ws = wb.AddWorksheet();
        ws.Cell(1, 1).Value = "Saved with SixLabors v2";
        ws.Column(1).AdjustToContents();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        // Reload with same font engine
        ms.Position = 0;
        using var wb2 = new XLWorkbook(ms, new LoadOptions { FontEngine = _engine });
        var value = wb2.Worksheet(1).Cell(1, 1).GetString();

        await Assert.That(value).IsEqualTo("Saved with SixLabors v2");
    }

    [Test]
    public async Task FontEngine_StreamBased_WorksWithWorkbook()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var engine = SixLaborsFontEngine.CreateOnlyWithFonts(fallbackStream);

        var loadOptions = new LoadOptions { FontEngine = engine };
        using var wb = new XLWorkbook(loadOptions);
        var ws = wb.AddWorksheet();
        ws.Cell(1, 1).Value = "Stream-based font";
        ws.Column(1).AdjustToContents();

        await Assert.That(ws.Column(1).Width).IsGreaterThan(0);
    }

    #endregion

    #region Bold / Italic variants

    [Test]
    public async Task BoldFont_ProducesValidMetrics()
    {
        var bold = new DummyFont("TestFontA", 11) { Bold = true };

        var boldWidth = _engine.GetTextWidth("Test text", bold, 96);

        // Bold font should still produce valid positive width
        await Assert.That(boldWidth).IsGreaterThan(0);
    }

    [Test]
    public async Task ItalicFont_ProducesValidMetrics()
    {
        var italic = new DummyFont("TestFontA", 11) { Italic = true };

        var italicWidth = _engine.GetTextWidth("Test text", italic, 96);

        // Italic may have different metrics — just verify it resolves without error
        await Assert.That(italicWidth).IsGreaterThan(0);
    }

    #endregion

    #region Constructor validation

    [Test]
    public async Task Constructor_ThrowsOnNullFallbackFont()
    {
        await Assert.That(() => new SixLaborsFontEngine(null!)).Throws<ArgumentException>();
    }

    [Test]
    public async Task Constructor_ThrowsOnWhitespaceFallbackFont()
    {
        await Assert.That(() => new SixLaborsFontEngine("   ")).Throws<ArgumentException>();
    }

    [Test]
    public async Task CreateOnlyWithFonts_ThrowsOnNullStream()
    {
        await Assert.That(() => SixLaborsFontEngine.CreateOnlyWithFonts(null!)).Throws<ArgumentNullException>();
    }

    #endregion

    private class DummyFont : IXLFontBase
    {
        public DummyFont(string name, double size)
        {
            FontName = name;
            FontSize = size;
        }

        public string FontName { get; set; }
        public double FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Strikethrough { get; set; }
        public XLFontUnderlineValues Underline { get; set; } = XLFontUnderlineValues.None;
        public XLFontVerticalTextAlignmentValues VerticalAlignment { get; set; }
        public bool Shadow { get; set; }
        public XLColor FontColor { get; set; } = XLColor.Black;
        public XLFontFamilyNumberingValues FontFamilyNumbering { get; set; } = XLFontFamilyNumberingValues.NotApplicable;
        public XLFontCharSet FontCharSet { get; set; } = XLFontCharSet.Default;
        public XLFontScheme FontScheme { get; set; }
    }
}
