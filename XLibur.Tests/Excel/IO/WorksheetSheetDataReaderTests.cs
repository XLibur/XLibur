using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using XLibur.Excel.IO;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.IO;

public class WorksheetSheetDataReaderTests
{
    [Test]
    [Arguments("yyyy-MM-dd", XLDataType.DateTime)]
    [Arguments("YYYY-MM-DD", XLDataType.DateTime)]
    [Arguments("Yyyy-Mm-Dd", XLDataType.DateTime)]
    [Arguments("hh:mm:ss", XLDataType.TimeSpan)]
    [Arguments("HH:MM:SS", XLDataType.TimeSpan)]
    [Arguments("#,##0.00", XLDataType.Number)]
    [Arguments("0.00%", XLDataType.Number)]
    [Arguments("mm:ss", XLDataType.TimeSpan)]
    [Arguments("MM:SS", XLDataType.TimeSpan)]
    [Arguments("[Red]0.00", XLDataType.Number)]
    [Arguments("\"Date: \"yyyy-MM-dd", XLDataType.DateTime)]
    [Arguments("[$-409]MMMM D, YYYY", XLDataType.DateTime)]
    public async Task GetDataTypeFromFormat_handles_mixed_case(string format, XLDataType expected)
    {
        var result = WorksheetSheetDataReader.GetDataTypeFromFormat(format);
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments("General")]
    [Arguments("@")]
    [Arguments("")]
    public async Task GetDataTypeFromFormat_returns_null_for_non_numeric_date_formats(string format)
    {
        var result = WorksheetSheetDataReader.GetDataTypeFromFormat(format);
        await Assert.That(result).IsNull();
    }

    [Test]
    public async Task LoadRow_tracks_last_row_so_rows_without_r_attribute_increment_correctly()
    {
        // Create an xlsx where some <row> elements have explicit r attributes and some don't.
        // Row without r should increment from the last known row index.
        using var ms = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }));

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");

            // Row at r=5 with cell A5="First"
            var row5 = new Row(new Cell
            {
                CellReference = "A5",
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text("First"))
            })
            { RowIndex = 5 };

            // Row without RowIndex — should become row 6
            var rowNoIndex1 = new Row(new Cell
            {
                CellReference = "A6",
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text("Second"))
            });

            // Row at r=10 with cell A10="Third"
            var row10 = new Row(new Cell
            {
                CellReference = "A10",
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text("Third"))
            })
            { RowIndex = 10 };

            // Row without RowIndex — should become row 11
            var rowNoIndex2 = new Row(new Cell
            {
                CellReference = "A11",
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text("Fourth"))
            });

            worksheetPart.Worksheet = new Worksheet(new SheetData(row5, rowNoIndex1, row10, rowNoIndex2));
        }

        ms.Position = 0;
        using var wb = new XLWorkbook(ms);
        var ws = wb.Worksheets.First();

        await Assert.That(ws.Cell("A5").GetString()).IsEqualTo("First");
        await Assert.That(ws.Cell("A6").GetString()).IsEqualTo("Second");
        await Assert.That(ws.Cell("A10").GetString()).IsEqualTo("Third");
        await Assert.That(ws.Cell("A11").GetString()).IsEqualTo("Fourth");
    }

    /// <summary>
    /// #558. The attributes of an <c>&lt;f&gt;</c> are read from the reader's character buffer rather
    /// than as strings. That is a different path from <c>XmlReader.Value</c> for a value carrying a
    /// character reference, and it sees the attributes in document order, so each of these forms must
    /// load to the same three formulas.
    /// </summary>
    [Test]
    // What Excel writes.
    [Arguments("t=\"shared\" ref=\"B1:B3\" si=\"0\"", "t=\"shared\" si=\"0\"")]
    // An order XLibur never writes: si before t, and the type after the index.
    [Arguments("si=\"0\" ref=\"B1:B3\" t=\"shared\"", "si=\"0\" t=\"shared\"")]
    // Character references: "shared" and "0" spelled the long way round.
    [Arguments("t=\"s&#104;ared\" ref=\"B1:B3\" si=\"&#48;\"", "t=\"s&#104;ared\" si=\"&#48;\"")]
    // Whitespace round the index, which uint.Parse accepts.
    [Arguments("t=\"shared\" ref=\"B1:B3\" si=\" 0 \"", "t=\"shared\" si=\" 0 \"")]
    public async Task Shared_formula_attributes_load_the_same_whatever_form_they_take(
        string masterAttributes, string memberAttributes)
    {
        using var package = SharedPackage(masterAttributes, memberAttributes);
        await AssertSharedFormulaLoaded(package);
    }

    /// <summary>
    /// A value longer than the reader's scratch buffer arrives in several chunks and is assembled
    /// into a string instead, so the spill path must read it in full.
    /// </summary>
    [Test]
    public async Task A_shared_formula_index_longer_than_the_scratch_buffer_is_read_in_full()
    {
        // Sized from the buffer itself, so raising its capacity cannot quietly take this test off
        // the spill path -- which is the only coverage an attribute value has of it.
        var padded = new string(' ', WorksheetSheetDataReader.ValueBufferLength + 16) + "0";
        using var package = SharedPackage(
            $"t=\"shared\" ref=\"B1:B3\" si=\"{padded}\"", $"t=\"shared\" si=\"{padded}\"");

        await AssertSharedFormulaLoaded(package);
    }

    /// <summary>A shared formula index that no <see cref="uint"/> can hold is refused, not truncated.</summary>
    [Test]
    public async Task A_shared_formula_index_that_no_uint_can_hold_is_refused()
    {
        using var package = SharedPackage("t=\"shared\" ref=\"B1:B3\" si=\"4294967296\"", "t=\"shared\" si=\"0\"");
        package.Position = 0;

        await Assert.That(() => _ = new XLWorkbook(package)).Throws<OverflowException>();
    }

    /// <summary>A formula type the reader does not know is refused rather than loaded as a normal formula.</summary>
    [Test]
    public async Task A_formula_type_the_reader_does_not_know_is_refused()
    {
        using var package = SharedPackage("t=\"bogus\" ref=\"B1:B3\" si=\"0\"", "t=\"shared\" si=\"0\"");
        package.Position = 0;

        await Assert.That(() => _ = new XLWorkbook(package)).Throws<NotSupportedException>();
    }

    /// <summary>
    /// A package whose B1:B3 hold <c>A{row}*10</c> as one shared formula, with the attributes of the
    /// first cell and of the other two written exactly as given.
    /// </summary>
    /// <remarks>
    /// B2 and B3 are saved holding a formula of their own that the rewrite replaces with an empty
    /// shared member, so their loaded text can only be <c>A{row}*10</c> if the shared formula really
    /// drove them. Saving them as <c>A{row}*10</c> in the first place would let a rewrite that
    /// matched nothing pass on the text that was already there.
    /// </remarks>
    private static MemoryStream SharedPackage(string masterAttributes, string memberAttributes)
    {
        var package = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            for (var row = 1; row <= 3; row++)
            {
                ws.Cell(row, 1).Value = row;
                ws.Cell(row, 2).FormulaA1 = row == 1 ? "A1*10" : $"A{row}*999";
            }

            wb.SaveAs(package);
        }

        return package.RewriteSheet1(xml =>
        {
            xml = ReplaceOnce(xml, "<x:f>A1*10</x:f>", $"<x:f {masterAttributes}>A1*10</x:f>");
            xml = ReplaceOnce(xml, "<x:f>A2*999</x:f>", $"<x:f {memberAttributes} />");
            return ReplaceOnce(xml, "<x:f>A3*999</x:f>", $"<x:f {memberAttributes} />");
        });
    }

    private static string ReplaceOnce(string xml, string original, string rewritten)
    {
        if (!xml.Contains(original, StringComparison.Ordinal))
            throw new InvalidOperationException($"'{original}' was not found in the sheet part.");

        return xml.Replace(original, rewritten, StringComparison.Ordinal);
    }

    private static async Task AssertSharedFormulaLoaded(MemoryStream package)
    {
        package.Position = 0;
        using var wb = new XLWorkbook(package);
        var ws = wb.Worksheet("Sheet1");

        for (var row = 1; row <= 3; row++)
            await Assert.That(ws.Cell(row, 2).FormulaA1).IsEqualTo($"A{row}*10");
    }
}
