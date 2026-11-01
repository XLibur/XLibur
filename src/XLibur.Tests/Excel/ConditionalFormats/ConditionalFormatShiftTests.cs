using XLibur.Excel;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.ConditionalFormats;

public class ConditionalFormatShiftTests
{
    [Test]
    public async Task CFShiftedOnColumnInsert()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("CFShift");
        ws.Range("A1:A1").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AirForceBlue);
        ws.Range("A2:B2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AliceBlue);
        ws.Range("A3:C3").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Alizarin);
        ws.Range("B4:B6").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Almond);
        ws.Range("C7:D7").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Amaranth);
        ws.Cells("A1:D7").Value = 1;

        ws.Column(2).InsertColumnsAfter(2);
        var cf = ws.ConditionalFormats.ToArray();

        await Assert.That(cf.Length).IsEqualTo(5);
        await Assert.That(cf[0].Range.RangeAddress.ToString()).IsEqualTo("A1:A1");
        await Assert.That(cf[1].Range.RangeAddress.ToString()).IsEqualTo("A2:D2");
        await Assert.That(cf[2].Range.RangeAddress.ToString()).IsEqualTo("A3:E3");
        await Assert.That(cf[3].Range.RangeAddress.ToString()).IsEqualTo("B4:D6");
        await Assert.That(cf[4].Range.RangeAddress.ToString()).IsEqualTo("E7:F7");
    }

    [Test]
    public async Task CFShiftedOnRowInsert()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("CFShift");
        ws.Range("A1:A1").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AirForceBlue);
        ws.Range("B1:B2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AliceBlue);
        ws.Range("C1:C3").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Alizarin);
        ws.Range("D2:F2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Almond);
        ws.Range("G4:G5").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Amaranth);
        ws.Cells("A1:G5").Value = 1;

        ws.Row(2).InsertRowsBelow(2);
        var cf = ws.ConditionalFormats.ToArray();

        await Assert.That(cf.Length).IsEqualTo(5);
        await Assert.That(cf[0].Range.RangeAddress.ToString()).IsEqualTo("A1:A1");
        await Assert.That(cf[1].Range.RangeAddress.ToString()).IsEqualTo("B1:B4");
        await Assert.That(cf[2].Range.RangeAddress.ToString()).IsEqualTo("C1:C5");
        await Assert.That(cf[3].Range.RangeAddress.ToString()).IsEqualTo("D2:F4");
        await Assert.That(cf[4].Range.RangeAddress.ToString()).IsEqualTo("G6:G7");
    }

    [Test]
    public async Task CFShiftedOnColumnDelete()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("CFShift");
        ws.Range("A1:A1").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AirForceBlue);
        ws.Range("A2:B2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AliceBlue);
        ws.Range("A3:C3").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Alizarin);
        ws.Range("B4:B6").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Almond);
        ws.Range("C7:D7").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Amaranth);
        ws.Cells("A1:D7").Value = 1;

        ws.Column(2).Delete();
        var cf = ws.ConditionalFormats.ToArray();

        await Assert.That(cf.Length).IsEqualTo(4);
        await Assert.That(cf[0].Range.RangeAddress.ToString()).IsEqualTo("A1:A1");
        await Assert.That(cf[1].Range.RangeAddress.ToString()).IsEqualTo("A2:A2");
        await Assert.That(cf[2].Range.RangeAddress.ToString()).IsEqualTo("A3:B3");
        await Assert.That(cf[3].Range.RangeAddress.ToString()).IsEqualTo("B7:C7");
    }

    [Test]
    public async Task CFShiftedOnRowDelete()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("CFShift");
        ws.Range("A1:A1").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AirForceBlue);
        ws.Range("B1:B2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.AliceBlue);
        ws.Range("C1:C3").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Alizarin);
        ws.Range("D2:F2").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Almond);
        ws.Range("G4:G5").AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Amaranth);
        ws.Cells("A1:G5").Value = 1;

        ws.Row(2).Delete();
        var cf = ws.ConditionalFormats.ToArray();

        await Assert.That(cf.Length).IsEqualTo(4);
        await Assert.That(cf[0].Range.RangeAddress.ToString()).IsEqualTo("A1:A1");
        await Assert.That(cf[1].Range.RangeAddress.ToString()).IsEqualTo("B1:B1");
        await Assert.That(cf[2].Range.RangeAddress.ToString()).IsEqualTo("C1:C2");
        await Assert.That(cf[3].Range.RangeAddress.ToString()).IsEqualTo("G3:G4");
    }

    [Test]
    public async Task CFShiftedTruncateRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("CFShift");
        ws.AsRange().AddConditionalFormat().WhenGreaterThan(0).Fill.SetBackgroundColor(XLColor.Red);
        var cf = ws.ConditionalFormats.Single();

        ws.Row(2).InsertRowsAbove(1);
        await Assert.That(cf.Range.RangeAddress.IsValid).IsTrue();
        await Assert.That(cf.Range.RangeAddress.ToString()).IsEqualTo($"1:{XLHelper.MaxRowNumber}");

        ws.Column(2).InsertColumnsAfter(1);
        await Assert.That(cf.Range.RangeAddress.IsValid).IsTrue();
        await Assert.That(cf.Range.RangeAddress.ToString()).IsEqualTo($"1:{XLHelper.MaxRowNumber}");
    }
}
