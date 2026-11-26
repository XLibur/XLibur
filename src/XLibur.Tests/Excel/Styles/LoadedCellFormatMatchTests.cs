using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// A save writes each cell's style to a <c>cellXf</c> of the loaded stylesheet, and that
/// <c>cellXf</c>'s font to a <c>font</c> of it. Two loaded records can describe the same format
/// and still differ: two <c>cellXf</c>s in different named cell styles (#977), or two fonts of
/// which one states an element the other leaves to its default (#978). A cell must keep the one it
/// had.
/// </summary>
public class LoadedCellFormatMatchTests
{
    private const string Calibri11 =
        "<x:font><x:sz val=\"11\" /><x:color theme=\"1\" /><x:name val=\"Calibri\" /><x:family val=\"2\" /><x:scheme val=\"minor\" /></x:font>";

    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [Test]
    public async Task A_cell_keeps_its_named_cell_style_when_a_cell_format_of_another_style_has_the_same_format()
    {
        // cellXfs 1 and 2 have one format, the Comma number format. 1 is in the Normal style, 2 in
        // the Comma style. A1 is in Comma, A2 in Normal.
        using var saved = SaveAfterEdit(NamedStyleFixture());

        using (Assert.Multiple())
        {
            await Assert.That(CellStyleOf(saved, "A1")).IsEqualTo("Comma");
            await Assert.That(CellStyleOf(saved, "A2")).IsEqualTo("Normal");
        }
    }

    [Test]
    public async Task A_cell_given_a_new_format_keeps_its_named_cell_style()
    {
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(NamedStyleFixture()))
        {
            wb.Worksheet(1).Cell("A1").Style.Font.Bold = true;
            wb.SaveAs(saved);
        }

        await Assert.That(CellStyleOf(saved, "A1")).IsEqualTo("Comma");
    }

    [Test]
    public async Task A_cell_keeps_a_font_that_states_its_family_when_another_font_leaves_it_out()
    {
        // Font 1 leaves out the family, font 2 states it. Both decode to Arial 10 in the default
        // family. A1 has font 2, A2 font 1.
        using var saved = SaveAfterEdit(FontFixture(
            "<x:font><x:sz val=\"10\" /><x:name val=\"Arial\" /></x:font>",
            "<x:font><x:sz val=\"10\" /><x:name val=\"Arial\" /><x:family val=\"2\" /></x:font>"));

        using (Assert.Multiple())
        {
            await Assert.That(FontOf(saved, "A1")).IsEqualTo("sz=10 name=Arial family=2");
            await Assert.That(FontOf(saved, "A2")).IsEqualTo("sz=10 name=Arial");
        }
    }

    [Test]
    public async Task A_cell_keeps_a_font_that_states_its_name_and_size_when_another_font_leaves_them_out()
    {
        // Font 1 states only the colour, so it decodes to Calibri 11 like font 2.
        using var saved = SaveAfterEdit(FontFixture(
            "<x:font><x:color theme=\"1\" /></x:font>",
            "<x:font><x:name val=\"Calibri\" /><x:color theme=\"1\" /><x:sz val=\"11\" /></x:font>"));

        using (Assert.Multiple())
        {
            await Assert.That(FontOf(saved, "A1")).IsEqualTo("name=Calibri color=theme:1 sz=11");
            await Assert.That(FontOf(saved, "A2")).IsEqualTo("color=theme:1");
        }
    }

    /// <summary>
    /// Fonts 0 (Calibri 11) and 1. cellStyleXfs 0 is Normal and 1 Comma. cellXfs 0 is Normal,
    /// 1 the Comma format in the Normal style, 2 the same format in the Comma style. A1 has cellXf 2,
    /// A2 cellXf 1.
    /// </summary>
    private static MemoryStream NamedStyleFixture() =>
        Fixture(
            ("fonts", $"<x:fonts count=\"1\">{Calibri11}</x:fonts>"),
            ("cellStyleXfs",
                "<x:cellStyleXfs count=\"2\">" +
                "<x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" />" +
                "<x:xf numFmtId=\"43\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"0\" applyBorder=\"0\" " +
                "applyAlignment=\"0\" applyProtection=\"0\" /></x:cellStyleXfs>"),
            ("cellXfs",
                "<x:cellXfs count=\"3\">" +
                "<x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" />" +
                "<x:xf numFmtId=\"43\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" />" +
                "<x:xf numFmtId=\"43\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"1\" /></x:cellXfs>"),
            ("cellStyles",
                "<x:cellStyles count=\"2\"><x:cellStyle name=\"Comma\" xfId=\"1\" builtinId=\"3\" />" +
                "<x:cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></x:cellStyles>"));

    /// <summary>
    /// Fonts 0 (Calibri 11), 1 and 2, with one cellXf for each. A1 has font 2, A2 font 1.
    /// </summary>
    private static MemoryStream FontFixture(string font1, string font2) =>
        Fixture(
            ("fonts", $"<x:fonts count=\"3\">{Calibri11}{font1}{font2}</x:fonts>"),
            ("cellStyleXfs",
                "<x:cellStyleXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" /></x:cellStyleXfs>"),
            ("cellXfs",
                "<x:cellXfs count=\"3\">" +
                "<x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" />" +
                "<x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" />" +
                "<x:xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" /></x:cellXfs>"),
            ("cellStyles",
                "<x:cellStyles count=\"1\"><x:cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></x:cellStyles>"));

    /// <summary>
    /// A one-sheet workbook with a 1 in A1 and A2, with the given stylesheet elements, whose A1 has
    /// cellXf 2 and A2 cellXf 1.
    /// </summary>
    private static MemoryStream Fixture(params (string Element, string Xml)[] styles)
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");
            ws.Cell("A1").Value = 1;
            ws.Cell("A2").Value = 1;
            wb.SaveAs(package);
        }

        package.RewriteStyles(styles);
        return package.RewriteSheet1(xml => WithStyleIndex(WithStyleIndex(xml, "A1", 2), "A2", 1));
    }

    private static string WithStyleIndex(string sheetXml, string cell, int styleIndex) =>
        Regex.Replace(sheetXml, $"<x:c r=\"{cell}\"( s=\"\\d+\")?", $"<x:c r=\"{cell}\" s=\"{styleIndex}\"");

    /// <summary>Loads <paramref name="package"/>, writes a value to C1, and saves it.</summary>
    private static MemoryStream SaveAfterEdit(MemoryStream package)
    {
        var saved = new MemoryStream();
        using var wb = new XLWorkbook(package);
        wb.Worksheet(1).Cell("C1").Value = "edited";
        wb.SaveAs(saved);
        return saved;
    }

    /// <summary>The name of the cell style the saved cell format of <paramref name="cell"/> is in.</summary>
    private static string CellStyleOf(Stream saved, string cell)
    {
        var styles = XDocument.Parse(saved.ReadPart("xl/styles.xml")).Root!;
        var xfId = (int?)CellXf(saved, styles, cell).Attribute("xfId") ?? 0;
        return styles.Element(Main + "cellStyles")!.Elements(Main + "cellStyle")
            .Where(s => (int)s.Attribute("xfId")! == xfId)
            .Select(s => (string)s.Attribute("name")!)
            .FirstOrDefault() ?? $"no cell style for xfId {xfId}";
    }

    /// <summary>
    /// The saved font of <paramref name="cell"/>, as each element's name and value in document order,
    /// e.g. <c>sz=10 name=Arial</c>.
    /// </summary>
    private static string FontOf(Stream saved, string cell)
    {
        var styles = XDocument.Parse(saved.ReadPart("xl/styles.xml")).Root!;
        var fontId = (int)CellXf(saved, styles, cell).Attribute("fontId")!;
        var font = styles.Element(Main + "fonts")!.Elements(Main + "font").ElementAt(fontId);
        return string.Join(" ", font.Elements().Select(e =>
            $"{e.Name.LocalName}={string.Join(",", e.Attributes().Select(a => a.Name.LocalName == "val" ? a.Value : $"{a.Name.LocalName}:{a.Value}"))}"));
    }

    private static XElement CellXf(Stream saved, XElement styles, string cell)
    {
        var sheet = XDocument.Parse(saved.Sheet1Xml());
        var c = sheet.Descendants(Main + "c").Single(e => (string?)e.Attribute("r") == cell);
        var styleIndex = (int?)c.Attribute("s") ?? 0;
        return styles.Element(Main + "cellXfs")!.Elements(Main + "xf").ElementAt(styleIndex);
    }
}
