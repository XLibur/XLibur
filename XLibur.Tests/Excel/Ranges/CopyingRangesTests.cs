using System.Drawing;
using System.Linq;
using XLibur.Excel;
using XLibur.Excel.ConditionalFormats;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Ranges;

public class CopyingRangesTests
{
    [Test]
    public async Task CopyingColumns()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        var column1 = ws.Column(1);
        column1.Cell(1).Style.Fill.SetBackgroundColor(XLColor.Red);
        column1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.FromArgb(1, 1, 1));
        column1.Cell(3).Style.Fill.SetBackgroundColor(XLColor.FromHtml("#CCCCCC"));
        column1.Cell(4).Style.Fill.SetBackgroundColor(XLColor.FromIndex(26));
        column1.Cell(5).Style.Fill.SetBackgroundColor(XLColor.FromColor(Color.MediumSeaGreen));
        column1.Cell(6).Style.Fill.SetBackgroundColor(XLColor.FromName("Blue"));
        column1.Cell(7).Style.Fill.SetBackgroundColor(XLColor.FromTheme(XLThemeColor.Accent3));

        ws.Cell(1, 2).CopyFrom(column1);
        ws.Cell(1, 3).CopyFrom(column1.Column(1, 7));

        var column2 = ws.Column(2);
        await Assert.That(column2.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column2.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromArgb(1, 1, 1));
        await Assert.That(column2.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromHtml("#CCCCCC"));
        await Assert.That(column2.Cell(4).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromIndex(26));
        await Assert.That(column2.Cell(5).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromColor(Color.MediumSeaGreen));
        await Assert.That(column2.Cell(6).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromName("Blue"));
        await Assert.That(column2.Cell(7).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromTheme(XLThemeColor.Accent3));

        var column3 = ws.Column(3);
        await Assert.That(column3.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(column3.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromArgb(1, 1, 1));
        await Assert.That(column3.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromHtml("#CCCCCC"));
        await Assert.That(column3.Cell(4).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromIndex(26));
        await Assert.That(column3.Cell(5).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromColor(Color.MediumSeaGreen));
        await Assert.That(column3.Cell(6).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromName("Blue"));
        await Assert.That(column3.Cell(7).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromTheme(XLThemeColor.Accent3));
    }

    [Test]
    public async Task CopyingRows()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        var row1 = ws.Row(1);
        FillRow(row1);

        ws.Cell(2, 1).CopyFrom(row1);
        ws.Cell(3, 1).CopyFrom(row1.Row(1, 7));

        var row2 = ws.Row(2);
        await Assert.That(row2.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row2.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromArgb(1, 1, 1));
        await Assert.That(row2.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromHtml("#CCCCCC"));
        await Assert.That(row2.Cell(4).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromIndex(26));
        await Assert.That(row2.Cell(5).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromColor(Color.MediumSeaGreen));
        await Assert.That(row2.Cell(6).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromName("Blue"));
        await Assert.That(row2.Cell(7).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromTheme(XLThemeColor.Accent3));

        var row3 = ws.Row(3);
        await Assert.That(row3.Cell(1).Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);
        await Assert.That(row3.Cell(2).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromArgb(1, 1, 1));
        await Assert.That(row3.Cell(3).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromHtml("#CCCCCC"));
        await Assert.That(row3.Cell(4).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromIndex(26));
        await Assert.That(row3.Cell(5).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromColor(Color.MediumSeaGreen));
        await Assert.That(row3.Cell(6).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromName("Blue"));
        await Assert.That(row3.Cell(7).Style.Fill.BackgroundColor).IsEqualTo(XLColor.FromTheme(XLThemeColor.Accent3));

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(3);
        await Assert.That(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B1:B1").Values.Any(v => v.Value.Value == "G1" && v.Value.IsFormula)).IsTrue();
        await Assert.That(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B2:B2").Values.Any(v => v.Value.Value == "G2" && v.Value.IsFormula)).IsTrue();
        await Assert.That(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B3:B3").Values.Any(v => v.Value.Value == "G3" && v.Value.IsFormula)).IsTrue();
    }

    [Test]
    public async Task CopyingConditionalFormats()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet");

        FillRow(ws.Row(1));
        FillRow(ws.Row(2));
        FillRow(ws.Row(3));

        ((XLConditionalFormats)ws.ConditionalFormats).Consolidate();

        ws.Cell(5, 2).CopyFrom(ws.Row(2).Row(1, 7));

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(2);
        await Assert.That(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B1:B3").Values.Any(v => v.Value.Value == "G1" && v.Value.IsFormula)).IsTrue();
        await Assert.That(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "C5:C5").Values.Any(v => v.Value.Value == "H5" && v.Value.IsFormula)).IsTrue();
    }

    [Test]
    public async Task CopyingConditionalFormatsDifferentWorksheets()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        var format = ws1.Range("A1:J2").AddConditionalFormat();

        var address = format.Ranges
            .First()
            .FirstCell()
            .CellRight(4)
            .Address
            .ToStringRelative();

        format.WhenEquals("=" + address)
            .Fill
            .SetBackgroundColor(XLColor.Blue);

        var ws2 = wb.Worksheets.Add("Sheet2");

        ws2.FirstCell().CopyFrom(ws1.Range("B1:B4"));

        await Assert.That(ws2.ConditionalFormats.Count()).IsEqualTo(1);
        await Assert.That(ws2.ConditionalFormats.All(x => x.Ranges.All(s => s.Worksheet == ws2))).IsTrue().Because("A conditional format was created for another worksheet.");
        await Assert.That(ws2.ConditionalFormats
            .Single(x => x.Range.RangeAddress.ToStringRelative() == "A1:A2")
            .Values.Any(v => v.Value.Value == "E1" && v.Value.IsFormula)).IsTrue().Because("The formula has not been transferred correctly.");

        await Assert.That(ws1.ConditionalFormats.First().Ranges.First().Worksheet.Name).IsEqualTo("Sheet1");
        await Assert.That(ws2.ConditionalFormats.First().Ranges.First().Worksheet.Name).IsEqualTo("Sheet2");
        await Assert.That(ws1.ConditionalFormats.First().Ranges.First().RangeAddress.ToString()).IsEqualTo("A1:J2");
        await Assert.That(ws2.ConditionalFormats.First().Ranges.First().RangeAddress.ToString()).IsEqualTo("A1:A2");
    }

    [Test]
    public async Task CopyConditionalFormatColorScaleInRange()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet");

        ws.Row(1).Cell(1).AddConditionalFormat()
            .ColorScale()
            .LowestValue(XLColor.Teal)
            .HighestValue(XLColor.Orange);

        ws.Cell(5, 2).CopyFrom(ws.Range(1, 1, 1, 5));

        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(2);
        await Assert.That(ws.ConditionalFormats.Single(x => x.Range.RangeAddress.ToStringRelative() == "B5:B5").ConditionalFormatType).IsEqualTo(XLConditionalFormatType.ColorScale);
    }

    private static void FillRow(IXLRow row1)
    {
        row1.Cell(1).Style.Fill.SetBackgroundColor(XLColor.Red);
        row1.Cell(2).Style.Fill.SetBackgroundColor(XLColor.FromArgb(1, 1, 1));
        row1.Cell(3).Style.Fill.SetBackgroundColor(XLColor.FromHtml("#CCCCCC"));
        row1.Cell(4).Style.Fill.SetBackgroundColor(XLColor.FromIndex(26));
        row1.Cell(5).Style.Fill.SetBackgroundColor(XLColor.FromColor(Color.MediumSeaGreen));
        row1.Cell(6).Style.Fill.SetBackgroundColor(XLColor.FromName("Blue"));
        row1.Cell(7).Style.Fill.SetBackgroundColor(XLColor.FromTheme(XLThemeColor.Accent3));

        row1.Cell(2).AddConditionalFormat().WhenEquals("=" + row1.FirstCell().CellRight(6).Address.ToStringRelative()).Fill.SetBackgroundColor(XLColor.Blue);
    }
}
