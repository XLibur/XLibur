using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// The fonts of a loaded stylesheet keep their order on save, so font 0 is still the file's first
/// font, which need not be the default font. A cell with the default style must be saved with the
/// index of the default font, wherever it sits, not with font 0 (#792). Borders keep their order
/// too, so the same holds for the default border.
/// </summary>
public class DefaultStyleFontIdSaveTests
{
    private const string Arial20 =
        "<x:font><x:sz val=\"20\" /><x:color rgb=\"FF000000\" /><x:name val=\"Arial\" /><x:family val=\"2\" /></x:font>";

    private const string Calibri11 =
        "<x:font><x:sz val=\"11\" /><x:color rgb=\"FF000000\" /><x:name val=\"Calibri\" /><x:family val=\"2\" /></x:font>";

    private const string ThinLeft =
        "<x:border><x:left style=\"thin\"><x:color indexed=\"64\" /></x:left><x:right /><x:top /><x:bottom /><x:diagonal /></x:border>";

    private const string NoBorder = "<x:border><x:left /><x:right /><x:top /><x:bottom /><x:diagonal /></x:border>";

    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(1);

    [Test]
    // The Normal style names the cells' font, Calibri 11.
    [Arguments("1")]
    // The Normal style names the first font, Arial 20, as Excel reports it whatever the style says.
    [Arguments("0")]
    public async Task Cells_in_the_default_font_keep_it_when_the_first_font_is_another(string normalStyleFontId)
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(Fixture(normalStyleFontId)))
        {
            wb.Worksheet(1).Cell("E5").Value = "edited";
            wb.SaveAs(saved);
        }

        var styles = XDocument.Parse(saved.ReadPart("xl/styles.xml"));
        var fonts = styles.Root!.Element(Main + "fonts")!.Elements(Main + "font").ToList();
        var cellXfs = styles.Root.Element(Main + "cellXfs")!.Elements(Main + "xf").ToList();

        // The first font stays first, so the columns are still measured in its digits.
        await Assert.That(FontName(fonts[0])).IsEqualTo("Arial 20");

        foreach (var cell in new[] { "E5", "E6" })
        {
            var cellXf = cellXfs[StyleIndex(saved.Sheet1Xml(), cell)];
            var fontId = (int)cellXf.Attribute("fontId")!;
            await Assert.That(FontName(fonts[fontId])).IsEqualTo("Calibri 11").Because(cell);
        }

        // And the reloaded cells read it.
        using var reloaded = new XLWorkbook(saved);
        foreach (var cell in new[] { "E5", "E6" })
        {
            var font = reloaded.Worksheet(1).Cell(cell).Style.Font;
            await Assert.That($"{font.FontName} {font.FontSize}").IsEqualTo("Calibri 11").Because(cell);
        }
    }

    [Test]
    public async Task Cells_without_a_border_keep_none_when_the_first_border_is_another()
    {
        // The same holds for borders: the first border is a thin left border nothing uses.
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(Fixture("1")))
        {
            wb.Worksheet(1).Cell("E5").Value = "edited";
            wb.SaveAs(saved);
        }

        var styles = XDocument.Parse(saved.ReadPart("xl/styles.xml"));
        var borders = styles.Root!.Element(Main + "borders")!.Elements(Main + "border").ToList();
        var cellXfs = styles.Root.Element(Main + "cellXfs")!.Elements(Main + "xf").ToList();

        foreach (var cell in new[] { "E5", "E6" })
        {
            var cellXf = cellXfs[StyleIndex(saved.Sheet1Xml(), cell)];
            var border = borders[(int)cellXf.Attribute("borderId")!];
            await Assert.That(border.Descendants().Any(e => e.Attribute("style") is not null)).IsFalse().Because(cell);
        }

        using var reloaded = new XLWorkbook(saved);
        foreach (var cell in new[] { "E5", "E6" })
            await Assert.That(reloaded.Worksheet(1).Cell(cell).Style.Border.LeftBorder)
                .IsEqualTo(XLBorderStyleValues.None).Because(cell);
    }

    private static string FontName(XElement font) =>
        $"{(string)font.Element(Main + "name")!.Attribute("val")!} {(string)font.Element(Main + "sz")!.Attribute("val")!}";

    private static int StyleIndex(string sheetXml, string cellReference)
    {
        var sheet = XDocument.Parse(sheetXml);
        var cell = sheet.Descendants(Main + "c").Single(c => (string?)c.Attribute("r") == cellReference);
        return (int?)cell.Attribute("s") ?? 0;
    }

    /// <summary>
    /// A one-sheet workbook with a 1 in E5 and E6, whose first font is Arial 20 and whose first
    /// border is a thin left border. The cells are Calibri 11 with no border, through font 1 and
    /// border 1. The Normal style's format names font <paramref name="normalStyleFontId"/>.
    /// </summary>
    private static MemoryStream Fixture(string normalStyleFontId)
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("E5").Value = 1;
            ws.Cell("E6").Value = 1;
            wb.SaveAs(package);
        }

        return package.RewritePart("xl/styles.xml", xml =>
        {
            xml = Regex.Replace(xml, "<x:fonts.*?</x:fonts>", $"<x:fonts count=\"2\">{Arial20}{Calibri11}</x:fonts>",
                RegexOptions.None, RegexTimeout);
            xml = Regex.Replace(xml, "<x:borders.*?</x:borders>", $"<x:borders count=\"2\">{ThinLeft}{NoBorder}</x:borders>",
                RegexOptions.None, RegexTimeout);
            xml = Regex.Replace(xml, "<x:cellStyleXfs.*?</x:cellStyleXfs>",
                $"<x:cellStyleXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"{normalStyleFontId}\" fillId=\"0\" borderId=\"1\" /></x:cellStyleXfs>",
                RegexOptions.None, RegexTimeout);
            xml = Regex.Replace(xml, "<x:cellXfs.*?</x:cellXfs>",
                "<x:cellXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" /></x:cellXfs>",
                RegexOptions.None, RegexTimeout);
            return Regex.Replace(xml, "<x:cellStyles.*?</x:cellStyles>",
                "<x:cellStyles count=\"1\"><x:cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></x:cellStyles>",
                RegexOptions.None, RegexTimeout);
        });
    }
}
