using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Utils;
using XLibur.Utils;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// Indexed colours above 64 are system colours (65 is the window background, 81 the tooltip
/// text). They must load and save as indexed colours rather than collapse to automatic (#957).
/// </summary>
public class IndexedSystemColorTests
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [Test]
    public async Task Note_runs_keep_their_system_text_colour_and_charset()
    {
        using var saved = TestHelper.LoadAndSave(@"Other\Comments\InsetsUnitConversion.xlsx");

        var runProperties = XDocument.Parse(saved.ReadPartUnder("xl/comments")).Descendants(Main + "rPr").ToList();
        await Assert.That(runProperties).IsNotEmpty();
        foreach (var rPr in runProperties)
        {
            await Assert.That((string?)rPr.Element(Main + "color")?.Attribute("indexed")).IsEqualTo("81");
            await Assert.That((string?)rPr.Element(Main + "charset")?.Attribute("val")).IsEqualTo("1");
        }
    }

    [Test]
    public async Task An_index_too_large_for_an_int_reads_as_automatic()
    {
        var color = new DocumentFormat.OpenXml.Spreadsheet.Color { Indexed = uint.MaxValue }.ToXLiburColor();

        await Assert.That(color).IsEqualTo(XLColor.Automatic);
    }

    [Test]
    public async Task Note_run_loads_its_system_text_colour_as_indexed()
    {
        using var source = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Comments\InsetsUnitConversion.xlsx"));
        using var wb = new XLWorkbook(source);

        var color = wb.Worksheet(1).Cell("A1").GetComment().First().FontColor;
        await Assert.That(color.ColorType).IsEqualTo(XLColorType.Indexed);
        await Assert.That(color.Indexed).IsEqualTo(81);
    }

    [Test]
    [Arguments(@"TryToLoad\LO\xlsx\formats.xlsx")]
    [Arguments(@"TryToLoad\LO\xlsx\preserve_space.xlsx")]
    public async Task Shared_string_run_keeps_a_stated_default_charset(string fixture)
    {
        using var saved = TestHelper.LoadAndSave(fixture);

        var charsets = XDocument.Parse(saved.ReadPart("xl/sharedStrings.xml"))
            .Descendants(Main + "rPr")
            .Select(rPr => (string?)rPr.Element(Main + "charset")?.Attribute("val"))
            .ToList();
        await Assert.That(charsets).IsNotEmpty();
        foreach (var charset in charsets)
            await Assert.That(charset).IsEqualTo("1");
    }

    [Test]
    public async Task A_run_that_never_stated_a_charset_writes_none()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet().Cell("A1").GetRichText().AddText("plain").SetBold();
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        var sst = saved.TryReadPart("xl/sharedStrings.xml", out var xml) ? xml : saved.Sheet1Xml();
        await Assert.That(sst).DoesNotContain("charset");
    }

    [Test]
    public async Task Solid_fill_in_the_system_background_colour_loads_as_solid()
    {
        using var source = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EscapedApostrophe.xlsx"));
        using var wb = new XLWorkbook(source);

        var fill = wb.Worksheet(1).Cell("G1").Style.Fill;
        await Assert.That(fill.PatternType).IsEqualTo(XLFillPatternValues.Solid);
        await Assert.That(fill.BackgroundColor.ColorType).IsEqualTo(XLColorType.Indexed);
        await Assert.That(fill.BackgroundColor.Indexed).IsEqualTo(65);
    }

    [Test]
    public async Task Solid_fill_in_the_system_background_colour_stays_solid_when_the_sheet_is_rewritten()
    {
        using var source = new MemoryStream();
        using (var fixture = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\EscapedApostrophe.xlsx")))
            fixture.CopyTo(source);

        using var saved = source.LoadAndSave(wb => wb.Worksheet(1).Cell("Z100").Value = 1);
        using var reloaded = new XLWorkbook(saved);

        var fill = reloaded.Worksheet(1).Cell("G1").Style.Fill;
        await Assert.That(fill.PatternType).IsEqualTo(XLFillPatternValues.Solid);
        await Assert.That(fill.BackgroundColor.Indexed).IsEqualTo(65);
    }

    [Test]
    public async Task Cell_font_in_a_system_colour_round_trips_as_indexed()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet().Cell("A1").SetValue("x").Style.Font.FontColor = XLColor.FromIndex(81);
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        using var reloaded = new XLWorkbook(saved);
        var color = reloaded.Worksheet(1).Cell("A1").Style.Font.FontColor;
        await Assert.That(color.ColorType).IsEqualTo(XLColorType.Indexed);
        await Assert.That(color.Indexed).IsEqualTo(81);
    }

    [Test]
    [Arguments(65, 0xFFFFFF)]
    [Arguments(78, 0xFFFFFF)]
    [Arguments(80, 0xFFFFE1)]
    [Arguments(81, 0x000000)]
    [Arguments(32767, 0x000000)]
    public async Task System_colour_resolves_to_the_default_windows_colour(int index, int rgb)
    {
        var color = XLColor.FromIndex(index).Color;

        await Assert.That(color.ToArgb()).IsEqualTo(unchecked((int)0xFF000000) | rgb);
    }
}
