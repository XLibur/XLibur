using XLibur.Excel;
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.ConditionalFormats;
using System.Threading.Tasks;
using OpenXmlFormula = DocumentFormat.OpenXml.Spreadsheet.Formula;

namespace XLibur.Tests.Excel.ConditionalFormats;

public class ConditionalFormatCopyTests
{
    [Test]
    public async Task StylesAreCreatedDuringCopy()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");
        var format = ws.Range("A1:A1").AddConditionalFormat();
        format.WhenEquals("=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()).Fill
            .SetBackgroundColor(XLColor.Blue);

        var wb2 = new XLWorkbook();
        var ws2 = wb2.Worksheets.Add("Sheet2");
        ws2.FirstCell().CopyFrom(ws.FirstCell());
        await Assert.That(ws2.ConditionalFormats.First().Style.Fill.BackgroundColor).IsEqualTo(XLColor.Blue); //Added blue style
    }

    [Test]
    public async Task CopyConditionalFormatSingleWorksheet()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");
        var format = ws.Range("A1:A1").AddConditionalFormat();
        format.WhenEquals("=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()).Fill
            .SetBackgroundColor(XLColor.Blue);

        ws.Cell("A1").CopyTo("B2");

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws.ConditionalFormats.First().Ranges.Count).IsEqualTo(2);
        await Assert.That(ws.ConditionalFormats.First().Ranges.First().RangeAddress.ToString()).IsEqualTo("A1:A1");
        await Assert.That(ws.ConditionalFormats.First().Ranges.Last().RangeAddress.ToString()).IsEqualTo("B2:B2");
    }

    [Test]
    public async Task CopyKeepsTheTrailingWhitespaceOfAFormulaLoadedFromAFile()
    {
        // The loader keeps a rule's formula text as the file has it, so its whitespace reaches the copy.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.Worksheets.Add("Sheet1").Range("A1").AddConditionalFormat().WhenEquals("=E1").Fill
                .SetBackgroundColor(XLColor.Blue);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, true))
        {
            var worksheet = doc.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            worksheet.Descendants<OpenXmlFormula>().Single().Text = "E1 ";
        }

        ms.Position = 0;
        using var loaded = new XLWorkbook(ms);
        var ws2 = loaded.Worksheets.Add("Sheet2");
        loaded.Worksheet("Sheet1").Cell("A1").CopyTo(ws2.Cell("B2"));

        await Assert.That(ws2.ConditionalFormats.Single().Values[1].Value).IsEqualTo("F2 ");
    }

    [Test]
    public async Task CopyConditionalFormatSameRange()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");
        var format = ws.Range("A1:C3").AddConditionalFormat();
        format.WhenEquals("=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()).Fill
            .SetBackgroundColor(XLColor.Blue);

        ws.Cell("A1").CopyTo("B2");

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws.ConditionalFormats.First().Ranges.Count).IsEqualTo(1);
        await Assert.That(ws.ConditionalFormats.First().Ranges.First().RangeAddress.ToString()).IsEqualTo("A1:C3");
    }

    [Test]
    public async Task CopyConditionalFormatsDifferentWorksheets()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        var format = ws1.Range("A1:A1").AddConditionalFormat();
        format.WhenEquals("=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()).Fill
            .SetBackgroundColor(XLColor.Blue);
        var ws2 = wb.Worksheets.Add("Sheet2");
        var otherCell = ws2.Cell("B2");

        ws1.Cell("A1").CopyTo(otherCell);

        await Assert.That(ws1.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws2.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws1.ConditionalFormats.First().Ranges.Count).IsEqualTo(1);
        await Assert.That(ws2.ConditionalFormats.First().Ranges.Count).IsEqualTo(1);
        await Assert.That(ws1.ConditionalFormats.First().Ranges.First().Worksheet.Name).IsEqualTo("Sheet1");
        await Assert.That(ws2.ConditionalFormats.First().Ranges.First().Worksheet.Name).IsEqualTo("Sheet2");
        await Assert.That(ws1.ConditionalFormats.First().Ranges.First().RangeAddress.ToString()).IsEqualTo("A1:A1");
        await Assert.That(ws2.ConditionalFormats.First().Ranges.First().RangeAddress.ToString()).IsEqualTo("B2:B2");
    }

    [Test]
    public async Task FullCopyConditionalFormatSameWorksheet()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        var format = (XLConditionalFormat)ws1.Range("A1:A1").AddConditionalFormat();
        format.WhenEquals("=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()).Fill
            .SetBackgroundColor(XLColor.Blue);

        await Assert.That(Action).Throws<InvalidOperationException>();
        return;

        void Action() => format.CopyTo(ws1);
    }

    [Test]
    public async Task FullCopyConditionalFormatDifferentWorksheets()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        var format = (XLConditionalFormat)ws1.Range("A1:C3").AddConditionalFormat();
        format.WhenEquals("=" + format.Ranges.First().FirstCell().CellRight(4).Address.ToStringRelative()).Fill
            .SetBackgroundColor(XLColor.Blue);
        var ws2 = wb.Worksheets.Add("Sheet2");

        format.CopyTo(ws2);

        await Assert.That(ws1.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws2.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws1.ConditionalFormats.First().Ranges.Count).IsEqualTo(1);
        await Assert.That(ws2.ConditionalFormats.First().Ranges.Count).IsEqualTo(1);
        await Assert.That(ws1.ConditionalFormats.First().Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("Sheet1!A1:C3");
        await Assert.That(ws2.ConditionalFormats.First().Ranges.First().RangeAddress.ToString(XLReferenceStyle.A1, true)).IsEqualTo("Sheet2!A1:C3");
    }
}
