using System;
using System.IO;
using NUnit.Framework;
using XLibur.Excel;
using XLibur.Graphics;

namespace XLibur.Fonts.SkiaSharp.Tests;

[TestFixture]
public class SkiaSharpFontEngineTests
{
    /// <summary>
    /// Stream-based engine using TestFontA as fallback — works on all platforms including CI (no system fonts needed).
    /// </summary>
    private static IXLFontEngine CreateTestEngine()
    {
        var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        return SkiaSharpFontEngine.CreateOnlyWithFonts(fallbackStream);
    }

    private readonly IXLFontEngine _engine = CreateTestEngine();

    #region Text width

    [Test]
    public void GetTextWidth_ReturnsPositiveValue()
    {
        var font = new DummyFont("TestFontA", 20);
        var width = _engine.GetTextWidth("Lorem ipsum dolor sit amet", font, 96);

        Assert.That(width, Is.GreaterThan(0));
    }

    [Test]
    public void GetTextWidth_LongerTextIsWider()
    {
        var font = new DummyFont("TestFontA", 11);
        var shortWidth = _engine.GetTextWidth("AB", font, 96);
        var longWidth = _engine.GetTextWidth("ABCDEF", font, 96);

        Assert.That(longWidth, Is.GreaterThan(shortWidth));
    }

    [Test]
    public void GetTextWidth_LargerFontIsWider()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 20);
        var smallWidth = _engine.GetTextWidth("Test", smallFont, 96);
        var largeWidth = _engine.GetTextWidth("Test", largeFont, 96);

        Assert.That(largeWidth, Is.GreaterThan(smallWidth));
    }

    [Test]
    public void GetTextWidth_HigherDpiIsWider()
    {
        var font = new DummyFont("TestFontA", 11);
        var width96 = _engine.GetTextWidth("Test", font, 96);
        var width120 = _engine.GetTextWidth("Test", font, 120);

        Assert.That(width120, Is.GreaterThan(width96));
    }

    [Test]
    public void GetTextWidth_EmptyStringReturnsZero()
    {
        var font = new DummyFont("TestFontA", 11);
        var width = _engine.GetTextWidth("", font, 96);

        Assert.That(width, Is.EqualTo(0));
    }

    #endregion

    #region Text height

    [Test]
    public void GetTextHeight_ReturnsPositiveValue()
    {
        var font = new DummyFont("TestFontA", 11);
        var height = _engine.GetTextHeight(font, 96);

        Assert.That(height, Is.GreaterThan(0));
    }

    [Test]
    public void GetTextHeight_LargerFontIsTaller()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 30);
        var smallHeight = _engine.GetTextHeight(smallFont, 96);
        var largeHeight = _engine.GetTextHeight(largeFont, 96);

        Assert.That(largeHeight, Is.GreaterThan(smallHeight));
    }

    [Test]
    public void GetTextHeight_HigherDpiIsTaller()
    {
        var font = new DummyFont("TestFontA", 11);
        var height96 = _engine.GetTextHeight(font, 96);
        var height120 = _engine.GetTextHeight(font, 120);

        Assert.That(height120, Is.GreaterThan(height96));
    }

    #endregion

    #region Max digit width

    [Test]
    public void GetMaxDigitWidth_ReturnsPositiveValue()
    {
        var font = new DummyFont("TestFontA", 11);
        var mdw = _engine.GetMaxDigitWidth(font, 96);

        Assert.That(mdw, Is.GreaterThan(0));
    }

    [Test]
    public void GetMaxDigitWidth_LargerFontIsWider()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 20);
        var smallMdw = _engine.GetMaxDigitWidth(smallFont, 96);
        var largeMdw = _engine.GetMaxDigitWidth(largeFont, 96);

        Assert.That(largeMdw, Is.GreaterThan(smallMdw));
    }

    #endregion

    #region Descent

    [Test]
    public void GetDescent_ReturnsPositiveValue()
    {
        var font = new DummyFont("TestFontA", 11);
        var descent = _engine.GetDescent(font, 96);

        Assert.That(descent, Is.GreaterThan(0));
    }

    [Test]
    public void GetDescent_LargerFontHasLargerDescent()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 30);
        var smallDescent = _engine.GetDescent(smallFont, 96);
        var largeDescent = _engine.GetDescent(largeFont, 96);

        Assert.That(largeDescent, Is.GreaterThan(smallDescent));
    }

    #endregion

    #region Glyph box

    [Test]
    public void GetGlyphBox_ReturnsPositiveAdvanceWidth()
    {
        var font = new DummyFont("TestFontA", 11);
        Span<int> codePoints = ['A'];
        var box = _engine.GetGlyphBox(codePoints, font, new Dpi(96, 96));

        Assert.That(box.AdvanceWidth, Is.GreaterThan(0));
        Assert.That(box.EmSize, Is.GreaterThan(0));
    }

    [Test]
    public void GetGlyphBox_MultipleCharactersProduceValidWidths()
    {
        var font = new DummyFont("TestFontA", 11);
        Span<int> charA = ['A'];
        Span<int> charB = ['B'];

        var boxA = _engine.GetGlyphBox(charA, font, new Dpi(96, 96));
        var boxB = _engine.GetGlyphBox(charB, font, new Dpi(96, 96));

        Assert.That(boxA.AdvanceWidth, Is.GreaterThan(0));
        Assert.That(boxB.AdvanceWidth, Is.GreaterThan(0));
    }

    [Test]
    public void GetGlyphBox_DescentIsPositive()
    {
        var font = new DummyFont("TestFontA", 11);
        Span<int> codePoints = ['g'];
        var box = _engine.GetGlyphBox(codePoints, font, new Dpi(96, 96));

        Assert.That(box.Descent, Is.GreaterThanOrEqualTo(0));
    }

    [Test]
    public void GetGlyphBox_LargerFontProducesLargerBox()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 20);
        Span<int> codePoints = ['A'];

        var smallBox = _engine.GetGlyphBox(codePoints, smallFont, new Dpi(96, 96));
        var largeBox = _engine.GetGlyphBox(codePoints, largeFont, new Dpi(96, 96));

        Assert.That(largeBox.AdvanceWidth, Is.GreaterThan(smallBox.AdvanceWidth));
        Assert.That(largeBox.EmSize, Is.GreaterThan(smallBox.EmSize));
    }

    #endregion

    #region Fallback behavior

    [Test]
    public void NonExistentFont_UsesFallback()
    {
        // With stream-based engine, non-existent fonts fall back to the provided fallback font
        var nonExistent = new DummyFont("TotallyFakeNonExistentFont12345", 11);
        var fallback = new DummyFont("TestFontA", 11);

        var nonExistentWidth = _engine.GetTextWidth("Test", nonExistent, 96);
        var fallbackWidth = _engine.GetTextWidth("Test", fallback, 96);

        Assert.That(nonExistentWidth, Is.EqualTo(fallbackWidth));
    }

    [Test]
    public void NonExistentFont_UsesFallbackForHeight()
    {
        var nonExistent = new DummyFont("TotallyFakeNonExistentFont12345", 14);
        var fallback = new DummyFont("TestFontA", 14);

        var nonExistentHeight = _engine.GetTextHeight(nonExistent, 96);
        var fallbackHeight = _engine.GetTextHeight(fallback, 96);

        Assert.That(nonExistentHeight, Is.EqualTo(fallbackHeight));
    }

    #endregion

    #region Stream-based factory methods

    [Test]
    public void CreateOnlyWithFonts_UsesProvidedFallback()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var engine = SkiaSharpFontEngine.CreateOnlyWithFonts(fallbackStream);

        var font = new DummyFont("Nonexistent Font", 20);
        var width = engine.GetTextWidth("A", font, 120);

        // Unknown font resolves to TestFontA; scaling with size and DPI must produce a sensible positive width.
        Assert.That(width, Is.GreaterThan(0));
    }

    [Test]
    public void CreateOnlyWithFonts_CanLoadExtraFonts()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        using var fontBStream = TestHelper.GetStreamFromResource("Fonts.TestFontB.ttf");
        var engine = SkiaSharpFontEngine.CreateOnlyWithFonts(fallbackStream, fontBStream);

        var widthB = engine.GetTextWidth("B", new DummyFont("TestFontB", 30), 96);
        var widthFallback = engine.GetTextWidth("B", new DummyFont("TestFontA", 30), 96);

        // TestFontB is loaded as an extra font, so it resolves to itself (not the TestFontA fallback).
        Assert.That(widthB, Is.GreaterThan(0));
        Assert.That(widthB, Is.Not.EqualTo(widthFallback));
    }

    [Test]
    public void CreateWithFontsAndSystemFonts_CanUseFallbackFont()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var engine = SkiaSharpFontEngine.CreateWithFontsAndSystemFonts(fallbackStream);

        // Even if system fonts aren't available, the fallback font should work
        var font = new DummyFont("NonexistentFont", 11);
        var width = engine.GetTextWidth("Test", font, 96);

        Assert.That(width, Is.GreaterThan(0));
    }

    #endregion

    #region Workbook integration

    [Test]
    public void FontEngine_WorksWithWorkbookViaLoadOptions()
    {
        var loadOptions = new LoadOptions { FontEngine = _engine };
        using var wb = new XLWorkbook(loadOptions);
        var ws = wb.AddWorksheet();

        ws.Cell(1, 1).Value = "Hello World";
        ws.Column(1).AdjustToContents();

        Assert.That(ws.Column(1).Width, Is.GreaterThan(0));
    }

    [Test]
    public void FontEngine_AdjustToContents_ProducesReasonableWidth()
    {
        var loadOptions = new LoadOptions { FontEngine = _engine };
        using var wb = new XLWorkbook(loadOptions);
        var ws = wb.AddWorksheet();

        ws.Cell(1, 1).Value = "Short";
        ws.Cell(2, 1).Value = "A much longer text that should need more width";

        ws.Column(1).AdjustToContents();

        // Width should accommodate the longer text
        Assert.That(ws.Column(1).Width, Is.GreaterThan(8.43)); // 8.43 is the default column width
    }

    [Test]
    public void FontEngine_AdjustRowHeight_ProducesReasonableHeight()
    {
        var loadOptions = new LoadOptions { FontEngine = _engine };
        using var wb = new XLWorkbook(loadOptions);
        var ws = wb.AddWorksheet();

        ws.Cell(1, 1).Value = "Test";
        ws.Row(1).AdjustToContents();

        Assert.That(ws.Row(1).Height, Is.GreaterThan(0));
    }

    [Test]
    public void FontEngine_CanSaveAndReloadWorkbook()
    {
        var loadOptions = new LoadOptions { FontEngine = _engine };
        using var wb = new XLWorkbook(loadOptions);
        var ws = wb.AddWorksheet();
        ws.Cell(1, 1).Value = "Saved with SkiaSharp";
        ws.Column(1).AdjustToContents();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        // Reload with same font engine
        ms.Position = 0;
        using var wb2 = new XLWorkbook(ms, new LoadOptions { FontEngine = _engine });
        var value = wb2.Worksheet(1).Cell(1, 1).GetString();

        Assert.That(value, Is.EqualTo("Saved with SkiaSharp"));
    }

    [Test]
    public void FontEngine_StreamBased_WorksWithWorkbook()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var engine = SkiaSharpFontEngine.CreateOnlyWithFonts(fallbackStream);

        var loadOptions = new LoadOptions { FontEngine = engine };
        using var wb = new XLWorkbook(loadOptions);
        var ws = wb.AddWorksheet();
        ws.Cell(1, 1).Value = "Stream-based font";
        ws.Column(1).AdjustToContents();

        Assert.That(ws.Column(1).Width, Is.GreaterThan(0));
    }

    #endregion

    #region Bold / Italic variants

    [Test]
    public void BoldFont_ProducesValidMetrics()
    {
        var bold = new DummyFont("TestFontA", 11) { Bold = true };

        var boldWidth = _engine.GetTextWidth("Test text", bold, 96);

        // Bold font should still produce valid positive width
        Assert.That(boldWidth, Is.GreaterThan(0));
    }

    [Test]
    public void ItalicFont_ProducesValidMetrics()
    {
        var italic = new DummyFont("TestFontA", 11) { Italic = true };

        var italicWidth = _engine.GetTextWidth("Test text", italic, 96);

        // Italic may have different metrics — just verify it resolves without error
        Assert.That(italicWidth, Is.GreaterThan(0));
    }

    #endregion

    #region Constructor validation

    [Test]
    public void Constructor_ThrowsOnNullFallbackFont()
    {
        Assert.Throws<ArgumentException>(() => new SkiaSharpFontEngine(null!));
    }

    [Test]
    public void Constructor_ThrowsOnWhitespaceFallbackFont()
    {
        Assert.Throws<ArgumentException>(() => new SkiaSharpFontEngine("   "));
    }

    [Test]
    public void CreateOnlyWithFonts_ThrowsOnNullStream()
    {
        Assert.Throws<ArgumentNullException>(() => SkiaSharpFontEngine.CreateOnlyWithFonts(null!));
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
