using System;
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Graphics;

namespace XLibur.Fonts.Tests;

/// <summary>
/// The behaviour every <see cref="IXLFontEngine"/> adapter must share, run once per engine.
/// </summary>
/// <remarks>
/// <para>
/// Linked into each font engine test project as source, and inherited by one small class per
/// engine that supplies the factories and keeps the tests whose expectations really differ. TUnit
/// registers the tests of a base class on a derived class only when the derived class carries
/// <c>[InheritsTests]</c>; an abstract base is never run on its own.
/// </para>
/// <para>
/// Every engine is built from the embedded TestFontA, so the suite needs no system fonts and runs
/// the same on CI as on a developer machine.
/// </para>
/// </remarks>
public abstract class FontEngineContractTests
{
    private IXLFontEngine? _engine;

    /// <summary>
    /// The engine under test: stream-based, with TestFontA as the fallback. Built on first use
    /// rather than in the constructor, which must not call the abstract factory.
    /// </summary>
    private IXLFontEngine Engine =>
        _engine ??= CreateOnlyWithFonts(TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf"));

    /// <summary>The engine's <c>CreateOnlyWithFonts</c> factory.</summary>
    protected abstract IXLFontEngine CreateOnlyWithFonts(Stream fallbackFontStream, params Stream[] fontStreams);

    /// <summary>The engine's <c>CreateWithFontsAndSystemFonts</c> factory.</summary>
    protected abstract IXLFontEngine CreateWithFontsAndSystemFonts(Stream fallbackFontStream);

    /// <summary>The engine's public constructor, which takes the name of the fallback font.</summary>
    protected abstract IXLFontEngine CreateWithFallbackFont(string fallbackFont);

    /// <summary>A short name for the engine, written into a cell by the round-trip test.</summary>
    protected abstract string EngineName { get; }

    #region Text width

    [Test]
    public async Task GetTextWidth_ReturnsPositiveValue()
    {
        var font = new DummyFont("TestFontA", 20);
        var width = Engine.GetTextWidth("Lorem ipsum dolor sit amet", font, 96);

        await Assert.That(width).IsGreaterThan(0);
    }

    [Test]
    public async Task GetTextWidth_LongerTextIsWider()
    {
        var font = new DummyFont("TestFontA", 11);
        var shortWidth = Engine.GetTextWidth("AB", font, 96);
        var longWidth = Engine.GetTextWidth("ABCDEF", font, 96);

        await Assert.That(longWidth).IsGreaterThan(shortWidth);
    }

    [Test]
    public async Task GetTextWidth_LargerFontIsWider()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 20);
        var smallWidth = Engine.GetTextWidth("Test", smallFont, 96);
        var largeWidth = Engine.GetTextWidth("Test", largeFont, 96);

        await Assert.That(largeWidth).IsGreaterThan(smallWidth);
    }

    [Test]
    public async Task GetTextWidth_HigherDpiIsWider()
    {
        var font = new DummyFont("TestFontA", 11);
        var width96 = Engine.GetTextWidth("Test", font, 96);
        var width120 = Engine.GetTextWidth("Test", font, 120);

        await Assert.That(width120).IsGreaterThan(width96);
    }

    [Test]
    public async Task GetTextWidth_EmptyStringReturnsZero()
    {
        var font = new DummyFont("TestFontA", 11);
        var width = Engine.GetTextWidth("", font, 96);

        await Assert.That(width).IsEqualTo(0);
    }

    #endregion

    #region Text height

    [Test]
    public async Task GetTextHeight_ReturnsPositiveValue()
    {
        var font = new DummyFont("TestFontA", 11);
        var height = Engine.GetTextHeight(font, 96);

        await Assert.That(height).IsGreaterThan(0);
    }

    [Test]
    public async Task GetTextHeight_LargerFontIsTaller()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 30);
        var smallHeight = Engine.GetTextHeight(smallFont, 96);
        var largeHeight = Engine.GetTextHeight(largeFont, 96);

        await Assert.That(largeHeight).IsGreaterThan(smallHeight);
    }

    [Test]
    public async Task GetTextHeight_HigherDpiIsTaller()
    {
        var font = new DummyFont("TestFontA", 11);
        var height96 = Engine.GetTextHeight(font, 96);
        var height120 = Engine.GetTextHeight(font, 120);

        await Assert.That(height120).IsGreaterThan(height96);
    }

    #endregion

    #region Max digit width

    [Test]
    public async Task GetMaxDigitWidth_ReturnsPositiveValue()
    {
        var font = new DummyFont("TestFontA", 11);
        var mdw = Engine.GetMaxDigitWidth(font, 96);

        await Assert.That(mdw).IsGreaterThan(0);
    }

    [Test]
    public async Task GetMaxDigitWidth_LargerFontIsWider()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 20);
        var smallMdw = Engine.GetMaxDigitWidth(smallFont, 96);
        var largeMdw = Engine.GetMaxDigitWidth(largeFont, 96);

        await Assert.That(largeMdw).IsGreaterThan(smallMdw);
    }

    #endregion

    #region Descent

    [Test]
    public async Task GetDescent_ReturnsPositiveValue()
    {
        var font = new DummyFont("TestFontA", 11);
        var descent = Engine.GetDescent(font, 96);

        await Assert.That(descent).IsGreaterThan(0);
    }

    [Test]
    public async Task GetDescent_LargerFontHasLargerDescent()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 30);
        var smallDescent = Engine.GetDescent(smallFont, 96);
        var largeDescent = Engine.GetDescent(largeFont, 96);

        await Assert.That(largeDescent).IsGreaterThan(smallDescent);
    }

    #endregion

    #region Glyph box

    [Test]
    public async Task GetGlyphBox_ReturnsPositiveAdvanceWidth()
    {
        var font = new DummyFont("TestFontA", 11);
        Span<int> codePoints = ['A'];
        var box = Engine.GetGlyphBox(codePoints, font, new Dpi(96, 96));

        await Assert.That(box.AdvanceWidth).IsGreaterThan(0);
        await Assert.That(box.EmSize).IsGreaterThan(0);
    }

    [Test]
    public async Task GetGlyphBox_MultipleCharactersProduceValidWidths()
    {
        var font = new DummyFont("TestFontA", 11);
        Span<int> charA = ['A'];
        Span<int> charB = ['B'];

        var boxA = Engine.GetGlyphBox(charA, font, new Dpi(96, 96));
        var boxB = Engine.GetGlyphBox(charB, font, new Dpi(96, 96));

        await Assert.That(boxA.AdvanceWidth).IsGreaterThan(0);
        await Assert.That(boxB.AdvanceWidth).IsGreaterThan(0);
    }

    [Test]
    public async Task GetGlyphBox_DescentIsPositive()
    {
        var font = new DummyFont("TestFontA", 11);
        Span<int> codePoints = ['g'];
        var box = Engine.GetGlyphBox(codePoints, font, new Dpi(96, 96));

        await Assert.That(box.Descent).IsGreaterThanOrEqualTo(0);
    }

    [Test]
    public async Task GetGlyphBox_LargerFontProducesLargerBox()
    {
        var smallFont = new DummyFont("TestFontA", 10);
        var largeFont = new DummyFont("TestFontA", 20);
        Span<int> codePoints = ['A'];

        var smallBox = Engine.GetGlyphBox(codePoints, smallFont, new Dpi(96, 96));
        var largeBox = Engine.GetGlyphBox(codePoints, largeFont, new Dpi(96, 96));

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

        var nonExistentWidth = Engine.GetTextWidth("Test", nonExistent, 96);
        var fallbackWidth = Engine.GetTextWidth("Test", fallback, 96);

        await Assert.That(nonExistentWidth).IsEqualTo(fallbackWidth);
    }

    [Test]
    public async Task NonExistentFont_UsesFallbackForHeight()
    {
        var nonExistent = new DummyFont("TotallyFakeNonExistentFont12345", 14);
        var fallback = new DummyFont("TestFontA", 14);

        var nonExistentHeight = Engine.GetTextHeight(nonExistent, 96);
        var fallbackHeight = Engine.GetTextHeight(fallback, 96);

        await Assert.That(nonExistentHeight).IsEqualTo(fallbackHeight);
    }

    #endregion

    #region Stream-based factory methods

    // CreateOnlyWithFonts_UsesProvidedFallback and CreateOnlyWithFonts_CanLoadExtraFonts assert
    // engine-specific numbers, so each engine's subclass holds its own version.

    [Test]
    public async Task CreateWithFontsAndSystemFonts_CanUseFallbackFont()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var engine = CreateWithFontsAndSystemFonts(fallbackStream);

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
        var loadOptions = new LoadOptions { FontEngine = Engine };
        using var wb = new XLWorkbook(loadOptions);
        var ws = wb.AddWorksheet();

        ws.Cell(1, 1).Value = "Hello World";
        ws.Column(1).AdjustToContents();

        await Assert.That(ws.Column(1).Width).IsGreaterThan(0);
    }

    [Test]
    public async Task FontEngine_AdjustToContents_ProducesReasonableWidth()
    {
        var loadOptions = new LoadOptions { FontEngine = Engine };
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
        var loadOptions = new LoadOptions { FontEngine = Engine };
        using var wb = new XLWorkbook(loadOptions);
        var ws = wb.AddWorksheet();

        ws.Cell(1, 1).Value = "Test";
        ws.Row(1).AdjustToContents();

        await Assert.That(ws.Row(1).Height).IsGreaterThan(0);
    }

    [Test]
    public async Task FontEngine_CanSaveAndReloadWorkbook()
    {
        var text = $"Saved with {EngineName}";
        var loadOptions = new LoadOptions { FontEngine = Engine };
        using var wb = new XLWorkbook(loadOptions);
        var ws = wb.AddWorksheet();
        ws.Cell(1, 1).Value = text;
        ws.Column(1).AdjustToContents();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        // Reload with same font engine
        ms.Position = 0;
        using var wb2 = new XLWorkbook(ms, new LoadOptions { FontEngine = Engine });
        var value = wb2.Worksheet(1).Cell(1, 1).GetString();

        await Assert.That(value).IsEqualTo(text);
    }

    [Test]
    public async Task FontEngine_StreamBased_WorksWithWorkbook()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var engine = CreateOnlyWithFonts(fallbackStream);

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

        var boldWidth = Engine.GetTextWidth("Test text", bold, 96);

        // Bold font should still produce valid positive width
        await Assert.That(boldWidth).IsGreaterThan(0);
    }

    [Test]
    public async Task ItalicFont_ProducesValidMetrics()
    {
        var italic = new DummyFont("TestFontA", 11) { Italic = true };

        var italicWidth = Engine.GetTextWidth("Test text", italic, 96);

        // Italic may have different metrics — just verify it resolves without error
        await Assert.That(italicWidth).IsGreaterThan(0);
    }

    #endregion

    #region Constructor validation

    [Test]
    public async Task Constructor_ThrowsOnNullFallbackFont()
    {
        await Assert.That(() => CreateWithFallbackFont(null!)).Throws<ArgumentException>();
    }

    [Test]
    public async Task Constructor_ThrowsOnWhitespaceFallbackFont()
    {
        await Assert.That(() => CreateWithFallbackFont("   ")).Throws<ArgumentException>();
    }

    [Test]
    public async Task CreateOnlyWithFonts_ThrowsOnNullStream()
    {
        await Assert.That(() => CreateOnlyWithFonts(null!)).Throws<ArgumentNullException>();
    }

    #endregion

    protected sealed class DummyFont : IXLFontBase
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
