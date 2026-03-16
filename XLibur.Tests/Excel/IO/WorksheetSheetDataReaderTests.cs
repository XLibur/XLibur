using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using XLibur.Excel;
using XLibur.Excel.IO;

namespace XLibur.Tests.Excel.IO;

[TestFixture]
public class WorksheetSheetDataReaderTests
{
    [TestCase("yyyy-MM-dd", XLDataType.DateTime)]
    [TestCase("YYYY-MM-DD", XLDataType.DateTime)]
    [TestCase("Yyyy-Mm-Dd", XLDataType.DateTime)]
    [TestCase("hh:mm:ss", XLDataType.TimeSpan)]
    [TestCase("HH:MM:SS", XLDataType.TimeSpan)]
    [TestCase("#,##0.00", XLDataType.Number)]
    [TestCase("0.00%", XLDataType.Number)]
    [TestCase("mm:ss", XLDataType.TimeSpan)]
    [TestCase("MM:SS", XLDataType.TimeSpan)]
    [TestCase("[Red]0.00", XLDataType.Number)]
    [TestCase("\"Date: \"yyyy-MM-dd", XLDataType.DateTime)]
    [TestCase("[$-409]MMMM D, YYYY", XLDataType.DateTime)]
    public void GetDataTypeFromFormat_handles_mixed_case(string format, XLDataType expected)
    {
        var result = WorksheetSheetDataReader.GetDataTypeFromFormat(format);
        Assert.That(result, Is.EqualTo(expected));
    }

    [TestCase("General")]
    [TestCase("@")]
    [TestCase("")]
    public void GetDataTypeFromFormat_returns_null_for_non_numeric_date_formats(string format)
    {
        var result = WorksheetSheetDataReader.GetDataTypeFromFormat(format);
        Assert.That(result, Is.Null);
    }

    [Test]
    public void LoadRow_tracks_last_row_so_rows_without_r_attribute_increment_correctly()
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

        Assert.That(ws.Cell("A5").GetString(), Is.EqualTo("First"));
        Assert.That(ws.Cell("A6").GetString(), Is.EqualTo("Second"));
        Assert.That(ws.Cell("A10").GetString(), Is.EqualTo("Third"));
        Assert.That(ws.Cell("A11").GetString(), Is.EqualTo("Fourth"));
    }
}
