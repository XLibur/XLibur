using XLibur.Excel;
using System;
using System.Linq;
using XLibur.Excel.ConditionalFormats;
using System.Threading.Tasks;

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
