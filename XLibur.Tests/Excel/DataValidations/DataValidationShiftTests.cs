using XLibur.Excel;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.DataValidations;

public class DataValidationShiftTests
{
    [Test]
    public async Task DataValidationShiftedOnColumnInsert()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("DataValidationShift");
        ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("A2:B2").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("A3:C3").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("B4:B6").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("C7:D7").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Cells("A1:D7").Value = 1;

        ws.Column(2).InsertColumnsAfter(2);
        var dv = ws.DataValidations.ToArray();

        await Assert.That(dv.Length).IsEqualTo(5);
        await Assert.That(dv[0].Ranges.Single().RangeAddress.ToString()).IsEqualTo("A1:A1");
        await Assert.That(dv[1].Ranges.Single().RangeAddress.ToString()).IsEqualTo("A2:D2");
        await Assert.That(dv[2].Ranges.Single().RangeAddress.ToString()).IsEqualTo("A3:E3");
        await Assert.That(dv[3].Ranges.Single().RangeAddress.ToString()).IsEqualTo("B4:D6");
        await Assert.That(dv[4].Ranges.Single().RangeAddress.ToString()).IsEqualTo("E7:F7");
    }

    [Test]
    public async Task DataValidationShiftedOnRowInsert()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("DataValidationShift");
        ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("B1:B2").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("C1:C3").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("D2:F2").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("G4:G5").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Cells("A1:G5").Value = 1;

        ws.Row(2).InsertRowsBelow(2);
        var dv = ws.DataValidations.ToArray();

        await Assert.That(dv.Length).IsEqualTo(5);
        await Assert.That(dv[0].Ranges.Single().RangeAddress.ToString()).IsEqualTo("A1:A1");
        await Assert.That(dv[1].Ranges.Single().RangeAddress.ToString()).IsEqualTo("B1:B4");
        await Assert.That(dv[2].Ranges.Single().RangeAddress.ToString()).IsEqualTo("C1:C5");
        await Assert.That(dv[3].Ranges.Single().RangeAddress.ToString()).IsEqualTo("D2:F4");
        await Assert.That(dv[4].Ranges.Single().RangeAddress.ToString()).IsEqualTo("G6:G7");
    }

    [Test]
    public async Task DataValidationShiftedOnColumnDelete()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("DataValidationShift");
        ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("A2:B2").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("A3:C3").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("B4:B6").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("C7:D7").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Cells("A1:D7").Value = 1;

        ws.Column(2).Delete();
        var dv = ws.DataValidations.ToArray();

        await Assert.That(dv.Length).IsEqualTo(4);
        await Assert.That(dv[0].Ranges.Single().RangeAddress.ToString()).IsEqualTo("A1:A1");
        await Assert.That(dv[1].Ranges.Single().RangeAddress.ToString()).IsEqualTo("A2:A2");
        await Assert.That(dv[2].Ranges.Single().RangeAddress.ToString()).IsEqualTo("A3:B3");
        await Assert.That(dv[3].Ranges.Single().RangeAddress.ToString()).IsEqualTo("B7:C7");
    }

    [Test]
    public async Task DataValidationShiftedOnRowDelete()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("DataValidationShift");
        ws.Range("A1:A1").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("B1:B2").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("C1:C3").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("D2:F2").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Range("G4:G5").CreateDataValidation().WholeNumber.Between(0, 1);
        ws.Cells("A1:G5").Value = 1;

        ws.Row(2).Delete();
        var dv = ws.DataValidations.ToArray();

        await Assert.That(dv.Length).IsEqualTo(4);
        await Assert.That(dv[0].Ranges.Single().RangeAddress.ToString()).IsEqualTo("A1:A1");
        await Assert.That(dv[1].Ranges.Single().RangeAddress.ToString()).IsEqualTo("B1:B1");
        await Assert.That(dv[2].Ranges.Single().RangeAddress.ToString()).IsEqualTo("C1:C2");
        await Assert.That(dv[3].Ranges.Single().RangeAddress.ToString()).IsEqualTo("G3:G4");
    }

    [Test]
    public async Task DataValidationShiftedTruncateRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("DataValidationShift");
        ws.AsRange().CreateDataValidation().WholeNumber.Between(0, 1);
        var dv = ws.DataValidations.Single();

        ws.Row(2).InsertRowsAbove(1);
        await Assert.That(dv.Ranges.Single().RangeAddress.IsValid).IsTrue();
        await Assert.That(dv.Ranges.Single().RangeAddress.ToString()).IsEqualTo($"1:{XLHelper.MaxRowNumber}");

        ws.Column(2).InsertColumnsAfter(1);
        await Assert.That(dv.Ranges.Single().RangeAddress.IsValid).IsTrue();
        await Assert.That(dv.Ranges.Single().RangeAddress.ToString()).IsEqualTo($"1:{XLHelper.MaxRowNumber}");
    }
}
