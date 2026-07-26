using XLibur.Excel;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Ranges;

public class RangesConsolidationTests
{
    [Test]
    public async Task ConsolidateRangesSameWorksheet()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var ranges = new XLRanges
        {
            ws.Range("A1:E3"),
            ws.Range("A4:B10"),
            ws.Range("E2:F12"),
            ws.Range("C6:I8"),
            ws.Range("G9:G9"),
            ws.Range("C9:D9"),
            ws.Range("H9:H9"),
            ws.Range("I9:I13"),
            ws.Range("C4:D5")
        };

        var consRanges = ranges.Consolidate().ToList();

        await Assert.That(consRanges.Count).IsEqualTo(6);
        await Assert.That(consRanges[0].RangeAddress.ToString()).IsEqualTo("A1:E9");
        await Assert.That(consRanges[1].RangeAddress.ToString()).IsEqualTo("F2:F12");
        await Assert.That(consRanges[2].RangeAddress.ToString()).IsEqualTo("G6:I9");
        await Assert.That(consRanges[3].RangeAddress.ToString()).IsEqualTo("A10:B10");
        await Assert.That(consRanges[4].RangeAddress.ToString()).IsEqualTo("E10:E12");
        await Assert.That(consRanges[5].RangeAddress.ToString()).IsEqualTo("I10:I13");
    }

    [Test]
    public async Task ConsolidateWideRangesSameWorksheet()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var ranges = new XLRanges
        {
            ws.Row(5),
            ws.Row(7),
            ws.Row(6),
            ws.Column("D"),
            ws.Column("F"),
            ws.Column("E")
        };

        var consRanges = ranges.Consolidate()
            .OrderBy(r => r.Worksheet.Name)
            .ThenBy(r => r.RangeAddress.FirstAddress.RowNumber)
            .ThenBy(r => r.RangeAddress.FirstAddress.ColumnNumber)
            .ToList();

        await Assert.That(consRanges.Count).IsEqualTo(3);
        await Assert.That(consRanges[0].RangeAddress.ToString()).IsEqualTo("D:F");
        await Assert.That(consRanges[1].RangeAddress.ToString()).IsEqualTo("A5:C7");
        await Assert.That(consRanges[2].RangeAddress.ToString()).IsEqualTo("G5:XFD7");
    }

    [Test]
    public async Task ConsolidateRangesDifferentWorksheets()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        var ws2 = wb.Worksheets.Add("Sheet2");
        var ranges = new XLRanges
        {
            ws1.Range("A1:E3"),
            ws1.Range("A4:B10"),
            ws1.Range("E2:F12"),
            ws1.Range("C6:I8"),
            ws1.Range("G9:G9"),
            ws2.Row(5),
            ws2.Row(7),
            ws2.Row(6),
            ws2.Column("D"),
            ws2.Column("F"),
            ws2.Column("E"),
            ws1.Range("C9:D9"),
            ws1.Range("H9:H9"),
            ws1.Range("I9:I13"),
            ws1.Range("C4:D5")
        };

        var consRanges = ranges.Consolidate()
            .OrderBy(r => r.Worksheet.Name)
            .ThenBy(r => r.RangeAddress.FirstAddress.RowNumber)
            .ThenBy(r => r.RangeAddress.FirstAddress.ColumnNumber)
            .ToList();

        await Assert.That(consRanges.Count).IsEqualTo(9);
        await Assert.That(consRanges[0].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("Sheet1!$A$1:$E$9");
        await Assert.That(consRanges[1].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("Sheet1!$F$2:$F$12");
        await Assert.That(consRanges[2].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("Sheet1!$G$6:$I$9");
        await Assert.That(consRanges[3].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("Sheet1!$A$10:$B$10");
        await Assert.That(consRanges[4].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("Sheet1!$E$10:$E$12");
        await Assert.That(consRanges[5].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("Sheet1!$I$10:$I$13");

        await Assert.That(consRanges[6].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("Sheet2!$D:$F");
        await Assert.That(consRanges[7].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("Sheet2!$A$5:$C$7");
        await Assert.That(consRanges[8].RangeAddress.ToStringFixed(XLReferenceStyle.Default, true)).IsEqualTo("Sheet2!$G$5:$XFD$7");
    }

    [Test]
    public async Task ConsolidateSparsedRanges()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var ranges = new XLRanges
        {
            ws.Range("A1:C1"),
            ws.Range("E1:G1"),
            ws.Range("A3:C3"),
            ws.Range("E3:G3")
        };

        var consRanges = ranges.Consolidate().ToList();

        await Assert.That(consRanges.Count).IsEqualTo(4);
        await Assert.That(consRanges[0].RangeAddress.ToString()).IsEqualTo("A1:C1");
        await Assert.That(consRanges[1].RangeAddress.ToString()).IsEqualTo("E1:G1");
        await Assert.That(consRanges[2].RangeAddress.ToString()).IsEqualTo("A3:C3");
        await Assert.That(consRanges[3].RangeAddress.ToString()).IsEqualTo("E3:G3");
    }
}
