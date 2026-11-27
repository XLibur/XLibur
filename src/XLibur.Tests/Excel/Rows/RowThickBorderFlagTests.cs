using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Rows;

/// <summary>
/// A row's <c>thickTop</c> and <c>thickBot</c> follow the borders of its cells and of the cells next to
/// it, as Excel writes them (#1031). Excel draws a row with either flag about 0.6pt taller, and does not
/// work the flags out again on load.
/// </summary>
public class RowThickBorderFlagTests
{
    [Test]
    [Arguments("B2", XLBorderStyleValues.Thin, "", "")]
    [Arguments("B2", XLBorderStyleValues.Medium, "thickBot", "")]
    [Arguments("B2", XLBorderStyleValues.MediumDashed, "thickBot", "")]
    [Arguments("B2", XLBorderStyleValues.SlantDashDot, "thickBot", "")]
    [Arguments("B2", XLBorderStyleValues.Thick, "thickBot", "thickTop")]
    [Arguments("B2", XLBorderStyleValues.Double, "thickBot", "thickTop")]
    public async Task Bottom_border_sets_the_flags_of_its_row_and_the_row_below(
        string cell, XLBorderStyleValues style, string row2, string row3)
    {
        using var saved = SaveNew(ws => ws.Cell(cell).Style.Border.BottomBorder = style);

        await Assert.That(Flags(saved, 2)).IsEqualTo(row2);
        await Assert.That(Flags(saved, 3)).IsEqualTo(row3);
    }

    [Test]
    [Arguments("B3", XLBorderStyleValues.Thin, "", "")]
    [Arguments("B3", XLBorderStyleValues.Medium, "thickBot", "")]
    [Arguments("B3", XLBorderStyleValues.Thick, "thickBot", "thickTop")]
    [Arguments("B3", XLBorderStyleValues.Double, "thickBot", "thickTop")]
    public async Task Top_border_sets_the_flags_of_its_row_and_the_row_above(
        string cell, XLBorderStyleValues style, string row2, string row3)
    {
        using var saved = SaveNew(ws => ws.Cell(cell).Style.Border.TopBorder = style);

        await Assert.That(Flags(saved, 2)).IsEqualTo(row2);
        await Assert.That(Flags(saved, 3)).IsEqualTo(row3);
    }

    [Test]
    public async Task Formatted_row_with_a_thick_bottom_border_sets_the_flags()
    {
        using var saved = SaveNew(ws => ws.Row(2).Style.Border.BottomBorder = XLBorderStyleValues.Thick);

        await Assert.That(Flags(saved, 2)).IsEqualTo("thickBot");
        await Assert.That(Flags(saved, 3)).IsEqualTo("thickTop");
    }

    [Test]
    public async Task Removing_a_thick_border_clears_the_flags_the_file_had()
    {
        using var source = ThickBottomBorderSource();
        source.Position = 0;
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            wb.Worksheet(1).Cell("A1").Style.Border.BottomBorder = XLBorderStyleValues.None;
            wb.SaveAs(saved);
        }

        await Assert.That(Flags(saved, 1)).IsEqualTo("");
        await Assert.That(Flags(saved, 2)).IsEqualTo("");
    }

    [Test]
    public async Task Flags_of_a_thick_border_the_file_had_are_kept_after_an_edit()
    {
        using var source = ThickBottomBorderSource();
        source.Position = 0;
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            wb.Worksheet(1).Cell("C5").Value = "edit";
            wb.SaveAs(saved);
        }

        await Assert.That(Flags(saved, 1)).IsEqualTo("thickBot");
        await Assert.That(Flags(saved, 2)).IsEqualTo("thickTop");
    }

    /// <summary>
    /// A1 has a thick bottom border, and the sheet has the flags Excel writes for it: <c>thickBot</c> on
    /// row 1, and an otherwise empty row 2 with <c>thickTop</c>.
    /// </summary>
    private static MemoryStream ThickBottomBorderSource()
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Cell("A1").Value = "text";
            ws.Cell("A1").Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            wb.SaveAs(stream);
        }

        // The save that built the source may have written the flags itself; set them the same either way.
        return stream.RewriteSheet1(xml => Regex.Replace(Regex.Replace(xml, " thick(?:Top|Bot)=\"1\"", ""),
                """<x:row r="2"[^>]*/>""", "")
            .Replace("""<x:row r="1" """, """<x:row r="1" thickBot="1" """, StringComparison.Ordinal)
            .Replace("</x:sheetData>", """<x:row r="2" thickTop="1" /></x:sheetData>""", StringComparison.Ordinal));
    }

    private static MemoryStream SaveNew(Action<IXLWorksheet> edit)
    {
        var saved = new MemoryStream();
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "text";
        edit(ws);
        wb.SaveAs(saved);
        return saved;
    }

    /// <summary>
    /// The thick border flags the saved first sheet gives a row, space separated; empty for a row the
    /// sheet does not have.
    /// </summary>
    private static string Flags(Stream xlsx, int rowNumber)
    {
        xlsx.Position = 0;
        using var doc = SpreadsheetDocument.Open(xlsx, isEditable: false);
        var part = doc.WorkbookPart!.WorksheetParts.Single();
        using var stream = part.GetStream();
        var xml = XDocument.Load(stream);
        var ns = xml.Root!.Name.Namespace;
        var row = xml.Root.Element(ns + "sheetData")!
            .Elements(ns + "row")
            .SingleOrDefault(r => r.Attribute("r")?.Value == rowNumber.ToString());
        return string.Join(" ", new[] { "thickTop", "thickBot" }
            .Where(name => row?.Attribute(name)?.Value is "1" or "true"));
    }
}
