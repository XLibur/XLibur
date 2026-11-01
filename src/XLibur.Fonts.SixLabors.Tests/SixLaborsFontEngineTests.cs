using System.IO;
using System.Threading.Tasks;
using XLibur.Fonts.Tests;
using XLibur.Graphics;

namespace XLibur.Fonts.SixLabors.Tests;

/// <summary>
/// The shared font engine contract, run against <see cref="SixLaborsFontEngine"/>, plus the two
/// factory tests that pin SixLabors' own measurements.
/// </summary>
[InheritsTests]
public class SixLaborsFontEngineTests : FontEngineContractTests
{
    protected override IXLFontEngine CreateOnlyWithFonts(Stream fallbackFontStream, params Stream[] fontStreams) =>
        SixLaborsFontEngine.CreateOnlyWithFonts(fallbackFontStream, fontStreams);

    protected override IXLFontEngine CreateWithFontsAndSystemFonts(Stream fallbackFontStream) =>
        SixLaborsFontEngine.CreateWithFontsAndSystemFonts(fallbackFontStream);

    protected override IXLFontEngine CreateWithFallbackFont(string fallbackFont) =>
        new SixLaborsFontEngine(fallbackFont);

    protected override string EngineName => "SixLabors v2";

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
}
