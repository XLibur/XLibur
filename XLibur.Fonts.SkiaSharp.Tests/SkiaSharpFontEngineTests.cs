using System.IO;
using System.Threading.Tasks;
using XLibur.Fonts.Tests;
using XLibur.Graphics;

namespace XLibur.Fonts.SkiaSharp.Tests;

/// <summary>
/// The shared font engine contract, run against <see cref="SkiaSharpFontEngine"/>, plus the two
/// factory tests whose expectations are SkiaSharp's own.
/// </summary>
[InheritsTests]
public class SkiaSharpFontEngineTests : FontEngineContractTests
{
    protected override IXLFontEngine CreateOnlyWithFonts(Stream fallbackFontStream, params Stream[] fontStreams) =>
        SkiaSharpFontEngine.CreateOnlyWithFonts(fallbackFontStream, fontStreams);

    protected override IXLFontEngine CreateWithFontsAndSystemFonts(Stream fallbackFontStream) =>
        SkiaSharpFontEngine.CreateWithFontsAndSystemFonts(fallbackFontStream);

    protected override IXLFontEngine CreateWithFallbackFont(string fallbackFont) =>
        new SkiaSharpFontEngine(fallbackFont);

    protected override string EngineName => "SkiaSharp";

    [Test]
    public async Task CreateOnlyWithFonts_UsesProvidedFallback()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        var engine = SkiaSharpFontEngine.CreateOnlyWithFonts(fallbackStream);

        var font = new DummyFont("Nonexistent Font", 20);
        var width = engine.GetTextWidth("A", font, 120);

        // Unknown font resolves to TestFontA; scaling with size and DPI must produce a sensible positive width.
        await Assert.That(width).IsGreaterThan(0);
    }

    [Test]
    public async Task CreateOnlyWithFonts_CanLoadExtraFonts()
    {
        using var fallbackStream = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        using var fontBStream = TestHelper.GetStreamFromResource("Fonts.TestFontB.ttf");
        var engine = SkiaSharpFontEngine.CreateOnlyWithFonts(fallbackStream, fontBStream);

        var widthB = engine.GetTextWidth("B", new DummyFont("TestFontB", 30), 96);
        var widthFallback = engine.GetTextWidth("B", new DummyFont("TestFontA", 30), 96);

        // TestFontB is loaded as an extra font, so it resolves to itself (not the TestFontA fallback).
        await Assert.That(widthB).IsGreaterThan(0);
        await Assert.That(widthB).IsNotEqualTo(widthFallback);
    }
}
