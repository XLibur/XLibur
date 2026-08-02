using XLibur.Excel;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.DataValidations;

public class XLDataValidationsTests
{
    [Test]
    public async Task CannotCreateWithoutWorksheet()
    {
        await Assert.That(() => new XLDataValidations(null!)).Throws<ArgumentNullException>();
    }

    [Test]
    public async Task AddedRangesAreTransferredToTargetSheet()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet();
        var ws2 = wb.AddWorksheet();

        var dv1 = ws1.Range("A1:A3").CreateDataValidation();
        dv1.MinValue = "100";

        var dv2 = ws2.DataValidations.Add(dv1);

        await Assert.That(ws1.DataValidations.Count()).IsEqualTo(1);
        await Assert.That(ws2.DataValidations.Count()).IsEqualTo(1);

        await Assert.That(dv2).IsNotSameReferenceAs(dv1);

        await Assert.That(dv1.Ranges.Single().Worksheet).IsSameReferenceAs(ws1);
        await Assert.That(dv2.Ranges.Single().Worksheet).IsSameReferenceAs(ws2);
    }

    [Test]
    [Arguments("A1:A1", true)]
    [Arguments("A1:A3", true)]
    [Arguments("A1:A4", false)]
    [Arguments("C2:C2", true)]
    [Arguments("C1:C3", true)]
    [Arguments("A1:C3", false)]
    public async Task CanFindDataValidationForRange(string searchAddress, bool expectedResult)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var dv = ws.Range("A1:A3").CreateDataValidation();
        dv.MinValue = "100";
        dv.AddRange(ws.Range("C1:C3"));

        var address = new XLRangeAddress(ws as XLWorksheet, searchAddress);

        var actualResult = ws.DataValidations.TryGet(address, out var foundDv);
        await Assert.That(actualResult).IsEqualTo(expectedResult);
        if (expectedResult)
            await Assert.That(foundDv).IsSameReferenceAs(dv);
        else
            await Assert.That(foundDv).IsNull();
    }

    [Test]
    [Arguments("A1:A1", 1)]
    [Arguments("A1:A3", 1)]
    [Arguments("B1:B4", 0)]
    [Arguments("A1:C3", 1)]
    [Arguments("C2:C3", 1)]
    [Arguments("C2:G6", 2)]
    [Arguments("E2:E3", 0)]
    public async Task CanGetAllDataValidationsForRange(string searchAddress, int expectedCount)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var dv1 = ws.Range("A1:A3").CreateDataValidation();
        dv1.MinValue = "100";
        dv1.AddRange(ws.Range("C1:C3"));

        var dv2 = ws.Range("E4:G6").CreateDataValidation();
        dv2.MinValue = "200";

        var address = new XLRangeAddress(ws as XLWorksheet, searchAddress);

        var actualResult = ws.DataValidations.GetAllInRange(address);

        await Assert.That(actualResult.Count()).IsEqualTo(expectedCount);
    }

    [Test]
    public async Task AddDataValidationSplitsExistingRanges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var dv1 = ws.Ranges("B2:G7,C11:C13").CreateDataValidation();
        dv1.MinValue = "100";

        var dv2 = ws.Range("E4:G6").CreateDataValidation();
        dv2.MinValue = "100";

        await Assert.That(dv1.Ranges.Count()).IsEqualTo(4);
        await Assert.That(string.Join(",", dv1.Ranges.Select(r => r.RangeAddress.ToString()))).IsEqualTo("B2:G3,B4:D6,B7:G7,C11:C13");
    }

    [Test]
    public async Task RemovedRangeExcludedFromIndex()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var dv = ws.Range("A1:A3").CreateDataValidation();
        dv.MinValue = "100";
        var range = ws.Range("C1:C3");
        dv.AddRange(range);

        dv.RemoveRange(range);

        var actualResult = ws.DataValidations.TryGet(range.RangeAddress, out var foundDv);
        await Assert.That(actualResult).IsFalse();
        await Assert.That(foundDv).IsNull();
    }

    [Test]
    public async Task ConsolidatedDataValidationsAreUnsubscribed()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var dv1 = ws.Range("A1:A3").CreateDataValidation();
        dv1.MinValue = "100";
        var dv2 = ws.Range("B1:B3").CreateDataValidation();
        dv2.MinValue = "100";

        (ws.DataValidations as XLDataValidations).Consolidate();
        dv1.AddRange(ws.Range("C1:C3"));
        dv2.AddRange(ws.Range("D1:D3"));

        var consolidatedDv = ws.DataValidations.Single();
        await Assert.That(consolidatedDv).IsSameReferenceAs(dv1);
        await Assert.That(ws.Cell("C1").HasDataValidation).IsTrue();
        await Assert.That(ws.Cell("D1").HasDataValidation).IsFalse();
    }
}
