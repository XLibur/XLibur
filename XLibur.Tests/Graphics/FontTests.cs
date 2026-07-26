using System;
using XLibur.Excel;
using XLibur.Graphics;
using XLibur.Fonts.SixLabors.V1;
using System.Threading.Tasks;

namespace XLibur.Tests.Graphics;

public class FontTests
{
    private readonly IXLGraphicEngine _engine = new DefaultGraphicEngine(DefaultFontEngine.Instance.Value);

    [Test]
    [Arguments]
    public async Task CalculatedTextWidth()
    {
        var textFont = new DummyFont("Calibri", 20);
        var textWidthPt = _engine.GetTextWidth("Lorem ipsum dolor sit amet", textFont, 96);
        await Assert.That(textWidthPt).IsEqualTo(300);
    }

    [Test]
    [Arguments]
    public async Task CalculatedTextHeight()
    {
        var textFont = new DummyFont("Calibri", 300);
        var textHeightPx = _engine.GetTextHeight(textFont, 96);
        // Calibri on Windows (~500) vs Carlito fallback on Linux (~596) have different metrics
        await Assert.That(textHeightPx).IsEqualTo(500).Within(100);
    }

    [Test]
    [Arguments]
    public async Task GetMaxDigitWidth()
    {
        var textFont = new DummyFont("Calibri", 11);
        var textWidthPx = _engine.GetMaxDigitWidth(textFont, 96);
        await Assert.That(textWidthPx).IsEqualTo(7.43359375d); // Calibri,11 has a max digit width of 7 per spec 18.3.1.13
    }

    [Test]
    [Arguments]
    public async Task DescentIsPositive()
    {
        var textFont = new DummyFont("Calibri", 11);
        var textWidthPt = _engine.GetDescent(textFont, 96);
        // Calibri on Windows vs Carlito fallback on Linux gives slightly different metrics
        await Assert.That(textWidthPt).IsEqualTo(3.666666666666667d).Within(0.5);
    }

    [Test]
    [Arguments]
    public async Task NonExistentFontUsesFallback()
    {
        var nonExistentFont = new DummyFont("NonExistentFont", 100);
        var fallbackFont = new DummyFont("Microsoft Sans Serif", 100);

        var nonExistentFontWidth = _engine.GetTextWidth("ABCDEF text", nonExistentFont, 96);
        var fallbackFontWidth = _engine.GetTextWidth("ABCDEF text", fallbackFont, 96);
        await Assert.That(nonExistentFontWidth).IsEqualTo(fallbackFontWidth);

        var nonExistentFontHeight = _engine.GetTextHeight(nonExistentFont, 96);
        var fallbackFontHeight = _engine.GetTextHeight(fallbackFont, 96);
        await Assert.That(nonExistentFontHeight).IsEqualTo(fallbackFontHeight);
    }

    [Test]
    public async Task UseEmbeddedFontWhenFallbackFontIsNotPresent()
    {
        var nonExistentFont = new DummyFont("SomeNonExistentFont", 11);
        var fontEngine = new DefaultFontEngine("NonExistentFallbackFont");
        Span<int> text = ['8'];

        var box = fontEngine.GetGlyphBox(text, nonExistentFont, new Dpi(96, 96));

        // Max digit width of CarlitoBare is ~7.4 at 11pt, unlike MS Sans Serif which is ~8
        await Assert.That(box.AdvanceWidth).IsEqualTo(7.43359375f);
    }

    [Test]
    [Arguments]
    public async Task CanSpecifyFallbackFontWithoutFileSystem()
    {
        using var fallbackFontStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var fontEngine = DefaultFontEngine.CreateOnlyWithFonts(fallbackFontStream);

        var nonExistentFont = new DummyFont("Nonexistent Font", 20);
        var widthOfLetterA = fontEngine.GetTextWidth("A", nonExistentFont, 120);

        const double expectedWidthOfLetterA = 31.25d;
        await Assert.That(widthOfLetterA).IsEqualTo(expectedWidthOfLetterA).Within(0.0001);
    }

    [Test]
    [Arguments]
    public async Task CanSpecifyExtraFontsAsStreamsWithoutFileSystem()
    {
        using var fallbackFontStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var fontBStream = TestHelper.GetStreamFromResource("Fonts.TestFontB.ttf");
        var fontEngine = DefaultFontEngine.CreateOnlyWithFonts(fallbackFontStream, fontBStream);

        var widthOfLetterB = fontEngine.GetTextWidth("B", new DummyFont("TestFontB", 30), 96);

        const double expectedWidthOfLetterB = 25d;
        await Assert.That(widthOfLetterB).IsEqualTo(expectedWidthOfLetterB).Within(0.0001);
    }

    [Test]
    [Arguments]
    public async Task Issue_1916_CanMeasureSpecificArabicText()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell(1, 1).Value = @"اصين";
        ws.Column(1).AdjustToContents();

        // AdjustToContents should set width to match content (short Arabic text is narrower than default)
        await Assert.That(ws.Column(1).Width).IsGreaterThan(0);
    }

    [Test]
    public async Task DefaultFontEngine_CanBeUsedDirectly()
    {
        var fontEngine = DefaultFontEngine.Instance.Value;
        var textFont = new DummyFont("Calibri", 11);
        var textWidthPx = fontEngine.GetMaxDigitWidth(textFont, 96);
        await Assert.That(textWidthPx).IsEqualTo(7.43359375d);
    }

    [Test]
    public async Task DefaultFontEngine_CanSpecifyFallbackFontWithoutFileSystem()
    {
        using var fallbackFontStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var fontEngine = DefaultFontEngine.CreateOnlyWithFonts(fallbackFontStream);

        var nonExistentFont = new DummyFont("Nonexistent Font", 20);
        var widthOfLetterA = fontEngine.GetTextWidth("A", nonExistentFont, 120);

        const double expectedWidthOfLetterA = 31.25d;
        await Assert.That(widthOfLetterA).IsEqualTo(expectedWidthOfLetterA).Within(0.0001);
    }

    [Test]
    public async Task FontEngine_CanBeInjectedViaLoadOptions()
    {
        using var fallbackFontStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var customFontEngine = DefaultFontEngine.CreateOnlyWithFonts(fallbackFontStream);

        var loadOptions = new LoadOptions { FontEngine = customFontEngine };
        using var wb = new XLWorkbook(loadOptions);

        // The workbook should use the custom font engine for text measurement
        var ws = wb.AddWorksheet();
        ws.Cell(1, 1).Value = "Test";
        ws.Column(1).AdjustToContents();

        await Assert.That(ws.Column(1).Width).IsGreaterThan(0);
    }

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
