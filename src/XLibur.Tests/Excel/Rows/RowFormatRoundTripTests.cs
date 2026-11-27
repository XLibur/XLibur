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
/// #958) and the thick border flags (<c>thickTop</c>, <c>thickBot</c>, #959), which the save works out
/// from the borders (#1031, <see cref="RowThickBorderFlagTests"/>).
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

    [Test]
    public async Task Styling_an_existing_row_gives_the_cell_where_it_meets_a_styled_column_both_styles()
    {
        using var source = RowWithHeightSource();
        using var saved = LoadAndSave(source, ws =>
        {
            ws.Row(2).Style.Font.Bold = true;
        });

        await Assert.That(SavedCell(saved, 2, "B2")).IsNotNull();

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var style = reloaded.Worksheet(1).Cell("B2").Style;
        await Assert.That(style.Font.Bold).IsTrue();
        await Assert.That(style.NumberFormat.Format).IsEqualTo(DateFormat);
    }

    [Test]
    public async Task Styling_an_existing_column_gives_the_cell_where_it_meets_a_formatted_row_both_styles()
    {
        using var source = BoldRowAndPlainColumnSource();
        using var saved = LoadAndSave(source, ws =>
        {
            ws.Column(2).Style.NumberFormat.Format = DateFormat;
        });

        await Assert.That(SavedCell(saved, 2, "B2")).IsNotNull();

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var style = reloaded.Worksheet(1).Cell("B2").Style;
        await Assert.That(style.Font.Bold).IsTrue();
        await Assert.That(style.NumberFormat.Format).IsEqualTo(DateFormat);
    }

    /// <summary>
    /// Column B has the date format, A1 holds text, and row 2 has a height and no cells, so the row
    /// exists on load without a B2 record.
    /// </summary>
    private static MemoryStream RowWithHeightSource()
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
            """<x:row r="2" ht="20" customHeight="1" /></x:sheetData>""", StringComparison.Ordinal));
    }

    /// <summary>
    /// Column B exists with a width and no style, A1 is bold, and row 2 is formatted with A1's bold
    /// style and has no cells, so B2 has no record on load.
    /// </summary>
    private static MemoryStream BoldRowAndPlainColumnSource()
    {
        var stream = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            ws.Column(2).Width = 20;
            ws.Cell("A1").Value = "text";
            ws.Cell("A1").Style.Font.Bold = true;
            wb.SaveAs(stream);
        }

        return stream.RewriteSheet1(xml =>
        {
            var boldStyle = System.Text.RegularExpressions.Regex.Match(xml, """<x:c r="A1" s="(\d+)" """).Groups[1].Value;
            return xml.Replace("</x:sheetData>",
                $"""<x:row r="2" s="{boldStyle}" customFormat="1" /></x:sheetData>""", StringComparison.Ordinal);
        });
    }

    private static XElement? SavedCell(Stream xlsx, int rowNumber, string reference)
    {
        var row = SavedRow(xlsx, 1, rowNumber);
        return row?.Elements(row.Name.Namespace + "c")
            .SingleOrDefault(c => c.Attribute("r")?.Value == reference);
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
