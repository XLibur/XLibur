using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Fonts.Tests;
using XLibur.Graphics;
using V1FontEngine = XLibur.Fonts.SixLabors.V1.DefaultFontEngine;

namespace XLibur.Fonts.PinataLabs.Tests;

/// <summary>
/// The engine supersedes <c>XLibur.Fonts.SixLabors.V1</c>, so a user who switches must get the same
/// column widths and row heights. Each test measures the same fonts, styles, sizes, resolutions and
/// texts with both engines and requires every result to be identical.
/// </summary>
/// <remarks>Remove this class, and the project reference to V1, when V1 is removed.</remarks>
[NotInParallel]
public class V1ParityTests
{
    private static readonly string[] FontNames = ["TestFontA", "TestFontB", "CarlitoBare", "Missing Font", "Aptos Narrow"];

    private static readonly string[] SystemFontNames = ["Calibri", "Arial", "Times New Roman", "Segoe UI", "Consolas", "Missing Font", "Aptos Narrow"];

    private static readonly double[] Sizes = [8, 11, 20.5];

    private static readonly double[] Dpis = [72, 96, 120];

    private static readonly string[] Texts =
    [
        "",
        "A",
        "Hello World",
        "0123456789",
        "Příliš žluťoučký kůň",
        "Ωμέγα и кириллица",
        "日本語のテキスト",
        "Tab\tand  spaces",
        "👨‍👩‍👧 emoji",
    ];

    private IXLFontEngine? _previousDefault;

    // Loading V1 runs its module initializer, which sets the process-wide default engine.
    [Before(HookType.Test)]
    public void SaveDefault() => _previousDefault = LoadOptions.DefaultFontEngine;

    [After(HookType.Test)]
    public void RestoreDefault() => LoadOptions.DefaultFontEngine = _previousDefault;

    [Test]
    public async Task StreamFonts_MeasureExactlyAsV1()
    {
        var pinata = CreateOnlyWithFonts(PinataLabsFontEngine.CreateOnlyWithFonts);
        var v1 = CreateOnlyWithFonts(V1FontEngine.CreateOnlyWithFonts);

        var mismatches = Compare(pinata, v1, FontNames);

        await Assert.That(mismatches).IsEmpty();
    }

    [Test]
    public async Task SystemFonts_MeasureExactlyAsV1()
    {
        var pinata = new PinataLabsFontEngine("Microsoft Sans Serif");
        var v1 = new V1FontEngine("Microsoft Sans Serif");

        var mismatches = Compare(pinata, v1, SystemFontNames);

        await Assert.That(mismatches).IsEmpty();
    }

    private static IXLFontEngine CreateOnlyWithFonts(Func<Stream, Stream[], IXLFontEngine> factory)
    {
        using var fallback = TestHelper.GetStreamFromResource("Fonts.TestFontA.ttf");
        using var fontB = TestHelper.GetStreamFromResource("Fonts.TestFontB.ttf");
        using var carlito = TestHelper.GetStreamFromResource("Fonts.CarlitoBare-Regular.ttf");
        return factory(fallback, [fontB, carlito]);
    }

    private static List<string> Compare(IXLFontEngine pinata, IXLFontEngine v1, string[] fontNames)
    {
        var mismatches = new List<string>();

        void Check<T>(string what, T actual, T expected)
        {
            if (!EqualityComparer<T>.Default.Equals(actual, expected))
                mismatches.Add($"{what}: PinataLabs {actual}, V1 {expected}");
        }

        foreach (var name in fontNames)
        foreach (var (bold, italic) in new[] { (false, false), (true, false), (false, true), (true, true) })
        foreach (var size in Sizes)
        foreach (var dpi in Dpis)
        {
            var font = new ParityFont(name, size, bold, italic);
            var id = $"{name} {size}pt{(bold ? " bold" : "")}{(italic ? " italic" : "")} @{dpi}";

            Check($"{id} height", pinata.GetTextHeight(font, dpi), v1.GetTextHeight(font, dpi));
            Check($"{id} descent", pinata.GetDescent(font, dpi), v1.GetDescent(font, dpi));
            Check($"{id} max digit", pinata.GetMaxDigitWidth(font, dpi), v1.GetMaxDigitWidth(font, dpi));

            foreach (var text in Texts)
            {
                Check($"{id} width '{text}'", pinata.GetTextWidth(text, font, dpi), v1.GetTextWidth(text, font, dpi));

                var elements = StringInfo.GetTextElementEnumerator(text);
                while (elements.MoveNext())
                {
                    var cluster = elements.GetTextElement().EnumerateRunes().Select(r => r.Value).ToArray();
                    var glyphDpi = new Dpi(dpi, dpi);
                    var actual = pinata.GetGlyphBox(cluster, font, glyphDpi);
                    var expected = v1.GetGlyphBox(cluster, font, glyphDpi);
                    var glyphId = $"{id} glyph '{elements.GetTextElement()}'";
                    Check($"{glyphId} advance", actual.AdvanceWidth, expected.AdvanceWidth);
                    Check($"{glyphId} em", actual.EmSize, expected.EmSize);
                    Check($"{glyphId} descent", actual.Descent, expected.Descent);
                }
            }
        }

        return mismatches;
    }

    private sealed class ParityFont : IXLFontBase
    {
        public ParityFont(string name, double size, bool bold, bool italic)
        {
            FontName = name;
            FontSize = size;
            Bold = bold;
            Italic = italic;
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
