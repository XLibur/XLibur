using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Styles;

/// <summary>
/// The workbook's style is the Normal cell style (<c>cellStyle builtinId="0"</c>). A cell style's
/// <c>xfId</c> is an index into <c>cellStyleXfs</c>, the style formats, not into <c>cellXfs</c>,
/// the cell formats (ECMA-376 Part 1, §18.8.7). Before #799 the load read it from <c>cellXfs</c>,
/// so the workbook took the style of whatever cell format sat at that index.
/// </summary>
public class NormalStyleLoadTests
{
    private const string Arial14Bold =
        "<x:font><x:b /><x:sz val=\"14\" /><x:color rgb=\"FF000000\" /><x:name val=\"Arial\" /><x:family val=\"2\" /></x:font>";

    private const string Calibri11 =
        "<x:font><x:sz val=\"11\" /><x:color rgb=\"FF000000\" /><x:name val=\"Calibri\" /><x:family val=\"2\" /></x:font>";

    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(1);

    [Test]
    public async Task The_workbook_style_is_the_Normal_style_format_not_the_cell_format_at_its_index()
    {
        // cellStyleXfs[0] is Arial 14 bold in General; cellXfs[0] is Calibri 11 in a date format.
        using var wb = new XLWorkbook(Fixture(
            "<x:cellStyleXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" /></x:cellStyleXfs>",
            "<x:cellXfs count=\"1\"><x:xf numFmtId=\"14\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" " +
            "applyNumberFormat=\"1\" applyFont=\"1\" /></x:cellXfs>",
            "<x:cellStyles count=\"1\"><x:cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></x:cellStyles>"));

        using (Assert.Multiple())
        {
            await Assert.That(wb.Style.Font.FontName).IsEqualTo("Arial");
            await Assert.That(wb.Style.Font.FontSize).IsEqualTo(14);
            await Assert.That(wb.Style.Font.Bold).IsTrue();
            await Assert.That(wb.Style.NumberFormat.NumberFormatId).IsEqualTo(0);
        }

        // A cell still takes its cell format.
        var cell = wb.Worksheet(1).Cell("E5");
        using (Assert.Multiple())
        {
            await Assert.That(cell.Style.Font.FontName).IsEqualTo("Calibri");
            await Assert.That(cell.Style.Font.FontSize).IsEqualTo(11);
            await Assert.That(cell.Style.Font.Bold).IsFalse();
            await Assert.That(cell.Style.NumberFormat.NumberFormatId).IsEqualTo(14);
        }
    }

    [Test]
    public async Task A_sheet_added_after_the_load_takes_the_Normal_style()
    {
        using var wb = new XLWorkbook(Fixture(
            "<x:cellStyleXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" /></x:cellStyleXfs>",
            "<x:cellXfs count=\"1\"><x:xf numFmtId=\"14\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" " +
            "applyNumberFormat=\"1\" applyFont=\"1\" /></x:cellXfs>",
            "<x:cellStyles count=\"1\"><x:cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></x:cellStyles>"));

        var cell = wb.AddWorksheet("Added").Cell("A1");

        using (Assert.Multiple())
        {
            await Assert.That(cell.Style.Font.FontName).IsEqualTo("Arial");
            await Assert.That(cell.Style.Font.FontSize).IsEqualTo(14);
            await Assert.That(cell.Style.NumberFormat.NumberFormatId).IsEqualTo(0);
        }
    }

    [Test]
    public async Task A_Normal_style_naming_a_missing_style_format_keeps_the_default_style()
    {
        // xfId 1 is past the end of cellStyleXfs. cellXfs has an entry at 1, which must not be used.
        using var wb = new XLWorkbook(Fixture(
            "<x:cellStyleXfs count=\"1\"><x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" /></x:cellStyleXfs>",
            "<x:cellXfs count=\"2\"><x:xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" />" +
            "<x:xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\" applyFont=\"1\" /></x:cellXfs>",
            "<x:cellStyles count=\"1\"><x:cellStyle name=\"Normal\" xfId=\"1\" builtinId=\"0\" /></x:cellStyles>"));

        using var inMemory = new XLWorkbook();
        using (Assert.Multiple())
        {
            await Assert.That(wb.Style.Font.FontName).IsEqualTo(inMemory.Style.Font.FontName);
            await Assert.That(wb.Style.Font.FontSize).IsEqualTo(inMemory.Style.Font.FontSize);
            await Assert.That(wb.Style.Font.Bold).IsFalse();
            await Assert.That(wb.Style.NumberFormat.NumberFormatId).IsEqualTo(0);
        }
    }

    /// <summary>
    /// A one-sheet workbook (a 1 in E5) whose stylesheet has two fonts, Arial 14 bold (0) and
    /// Calibri 11 (1), and the given style formats, cell formats and cell styles.
    /// </summary>
    private static MemoryStream Fixture(string cellStyleXfs, string cellXfs, string cellStyles)
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.Worksheets.Add("Sheet1").Cell("E5").Value = 1;
            wb.SaveAs(package);
        }

        return package.RewritePart("xl/styles.xml", xml =>
        {
            xml = Regex.Replace(xml, "<x:fonts.*?</x:fonts>", $"<x:fonts count=\"2\">{Arial14Bold}{Calibri11}</x:fonts>",
                RegexOptions.None, RegexTimeout);
            xml = Regex.Replace(xml, "<x:cellStyleXfs.*?</x:cellStyleXfs>", cellStyleXfs, RegexOptions.None, RegexTimeout);
            xml = Regex.Replace(xml, "<x:cellXfs.*?</x:cellXfs>", cellXfs, RegexOptions.None, RegexTimeout);
            return Regex.Replace(xml, "<x:cellStyles.*?</x:cellStyles>", cellStyles, RegexOptions.None, RegexTimeout);
        });
    }
}
