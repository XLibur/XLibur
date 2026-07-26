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
}
