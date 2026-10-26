using System.IO;
using System.Threading.Tasks;
using XLibur.Fonts.Tests;
using XLibur.Graphics;

namespace XLibur.Fonts.PinataLabs.Tests;

/// <summary>
/// The shared font engine contract, run against <see cref="PinataLabsFontEngine"/>, plus the two
/// factory tests that pin PinataLabs.Typography's own measurements.
/// </summary>
[InheritsTests]
public class PinataLabsFontEngineTests : FontEngineContractTests
{
    protected override IXLFontEngine CreateOnlyWithFonts(Stream fallbackFontStream, params Stream[] fontStreams) =>
        PinataLabsFontEngine.CreateOnlyWithFonts(fallbackFontStream, fontStreams);

    protected override IXLFontEngine CreateWithFontsAndSystemFonts(Stream fallbackFontStream) =>
        PinataLabsFontEngine.CreateWithFontsAndSystemFonts(fallbackFontStream);

    protected override IXLFontEngine CreateWithFallbackFont(string fallbackFont) =>
        new PinataLabsFontEngine(fallbackFont);

    protected override string EngineName => "PinataLabs";

    [Test]
    public async Task CreateOnlyWithFonts_UsesProvidedFallback()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var engine = PinataLabsFontEngine.CreateOnlyWithFonts(fallbackStream);

        var font = new DummyFont("Nonexistent Font", 20);
        var width = engine.GetTextWidth("A", font, 120);

        // TestFontA at 20pt, 120 DPI, the value XLibur.Tests pins for the V1 engine.
        await Assert.That(width).IsEqualTo(31.25d).Within(0.01);
    }

    [Test]
    public async Task CreateOnlyWithFonts_CanLoadExtraFonts()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        using var fontBStream = TestHelper.GetStreamFromResource("Fonts.TestFontB.ttf");
        var engine = PinataLabsFontEngine.CreateOnlyWithFonts(fallbackStream, fontBStream);

        var widthB = engine.GetTextWidth("B", new DummyFont("TestFontB", 30), 96);

        await Assert.That(widthB).IsEqualTo(25d).Within(0.01);
    }
}
