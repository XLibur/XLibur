using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Rows;

/// <summary>
/// Row attributes the file carries that the model has to keep: a formatted row (<c>customFormat</c>,
/// #958) and the thick border flags (<c>thickTop</c>, <c>thickBot</c>, #959).
/// </summary>
public class RowFormatRoundTripTests
{
    private const string DateFormat = "yyyy/mm/dd";

    [Test]
    public async Task Empty_cell_in_a_formatted_row_reads_the_row_style_not_the_column_style()
    {
        using var source = FormattedRowSource();
        using var wb = new XLWorkbook(source);
        var ws = wb.Worksheet(1);

        await Assert.That(ws.Cell("B1").Style.NumberFormat.Format).IsEqualTo(DateFormat);
        await Assert.That(ws.Cell("B2").Style.NumberFormat.NumberFormatId).IsEqualTo(0);
        await Assert.That(ws.Cell("B2").Style.NumberFormat.Format).IsNotEqualTo(DateFormat);
    }

    [Test]
    public async Task Formatted_row_with_the_sheet_style_is_saved_after_an_edit()
    {
        using var source = FormattedRowSource();
        using var saved = LoadAndSave(source, ws => ws.Cell("A1").Value = "edit");

        var row = SavedRow(saved, 1, 2);
        await Assert.That(row).IsNotNull();
        await Assert.That(row!.Attribute("customFormat")?.Value).IsEqualTo("1");
        await Assert.That(row.Attribute("s")?.Value).IsEqualTo("0");

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet(1).Cell("B2").Style.NumberFormat.NumberFormatId).IsEqualTo(0);
    }

    [Test]
    [Arguments(185)]
    [Arguments(259)]
    [Arguments(346)]
    public async Task Formatted_row_of_an_untouched_sheet_is_saved(int rowNumber)
    {
        using var fixture = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\DrawingSizeInPx.xlsx"));
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(fixture))
            wb.SaveAs(saved);

        var row = SavedRow(saved, 7, rowNumber);
        await Assert.That(row).IsNotNull();
        await Assert.That(row!.Attribute("customFormat")?.Value).IsEqualTo("1");
        await Assert.That(row.Attribute("s")).IsNotNull();
    }

    [Test]
    public async Task Thick_border_flags_are_saved_after_an_edit()
    {
        using var source = ThickRowSource();
        using var saved = LoadAndSave(source, ws => ws.Cell("A1").Value = "edit");

        var withCells = SavedRow(saved, 1, 1);
        await Assert.That(withCells!.Attribute("thickBot")?.Value).IsEqualTo("1");
        await Assert.That(withCells.Attribute("thickTop")?.Value).IsEqualTo("1");

        var empty = SavedRow(saved, 1, 3);
        await Assert.That(empty).IsNotNull();
        await Assert.That(empty!.Attribute("thickTop")?.Value).IsEqualTo("1");
        await Assert.That(empty.Attribute("thickBot")).IsNull();
    }

    [Test]
    [Arguments(13, "thickBot")]
    [Arguments(24, "thickTop")]
    [Arguments(26, "thickTop")]
    public async Task Thick_border_flags_of_a_fixture_are_saved_after_an_edit(int rowNumber, string attribute)
    {
        using var fixture = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\LO\xlsx\cell-borders.xlsx"));
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(fixture))
        {
            wb.Worksheet(1).Cell("H1").Value = "edit";
            wb.SaveAs(saved);
        }

        var row = SavedRow(saved, 1, rowNumber);
        await Assert.That(row!.Attribute(attribute)?.Value).IsEqualTo("1");
    }

    /// <summary>
    /// Column B has the date format, A1 holds text, and row 2 is formatted with the sheet style and
    /// has no cells: <c>&lt;row r="2" s="0" customFormat="1"/&gt;</c>.
    /// </summary>
    private static MemoryStream FormattedRowSource()
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Column(2).Style.NumberFormat.Format = DateFormat;
            ws.Cell("A1").Value = "text";
            wb.SaveAs(stream);
        }

        return stream.RewriteSheet1(xml => xml.Replace("</x:sheetData>",
            """<x:row r="2" s="0" customFormat="1" /></x:sheetData>""", StringComparison.Ordinal));
    }

    /// <summary>
    /// Row 1 holds A1 and has both thick border flags, and row 3 has no cells and only
    /// <c>thickTop</c>.
    /// </summary>
    private static MemoryStream ThickRowSource()
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet().Cell("A1").Value = "text";
            wb.SaveAs(stream);
        }

        return stream.RewriteSheet1(xml => xml
            .Replace("""<x:row r="1" """, """<x:row r="1" thickTop="1" thickBot="1" """, StringComparison.Ordinal)
            .Replace("</x:sheetData>", """<x:row r="3" thickTop="1" /></x:sheetData>""", StringComparison.Ordinal));
    }

    private static MemoryStream LoadAndSave(MemoryStream source, Action<IXLWorksheet> edit)
    {
        source.Position = 0;
        using var wb = new XLWorkbook(source);
        edit(wb.Worksheet(1));
        var saved = new MemoryStream();
        wb.SaveAs(saved);
        return saved;
    }

    private static XElement? SavedRow(Stream xlsx, int sheetPosition, int rowNumber)
    {
        xlsx.Position = 0;
        using var doc = SpreadsheetDocument.Open(xlsx, isEditable: false);
        var workbookPart = doc.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .ElementAt(sheetPosition - 1);
        var part = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        using var stream = part.GetStream();
        var xml = XDocument.Load(stream);
        var ns = xml.Root!.Name.Namespace;
        return xml.Root.Element(ns + "sheetData")!
            .Elements(ns + "row")
            .SingleOrDefault(r => r.Attribute("r")?.Value == rowNumber.ToString());
    }
}
