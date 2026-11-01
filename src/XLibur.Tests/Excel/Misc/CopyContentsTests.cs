using System;
using XLibur.Excel;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Misc;

public class CopyContentsTests
{
    private static void CopyRowAsRange(IXLWorksheet originalSheet, int originalRowNumber, IXLWorksheet destSheet,
        int destRowNumber)
    {
        var destinationRow = destSheet.Row(destRowNumber);
        destinationRow.Clear();

        var originalRow = originalSheet.Row(originalRowNumber);
        var columnNumber = originalRow.LastCellUsed(XLCellsUsedOptions.All)!.Address.ColumnNumber;

        var originalRange = originalSheet.Range(originalRowNumber, 1, originalRowNumber, columnNumber);
        var destRange = destSheet.Range(destRowNumber, 1, destRowNumber, columnNumber);
        originalRange.CopyTo(destRange);
    }

    [Test]
    public async Task CopyConditionalFormatsCount()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().AddConditionalFormat().WhenContains("1").Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.FirstCell().AsRange());
        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(2);
    }

    [Test]
    public async Task CopyConditionalFormatsFixedNum()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "1";
        ws.Cell("B1").Value = "1";
        ws.Cell("A1").AddConditionalFormat().WhenEquals(1).Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
        await Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula))).IsTrue();
        await Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula))).IsTrue();
    }

    [Test]
    public async Task CopyConditionalFormatsFixedString()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "A";
        ws.Cell("B1").Value = "B";
        ws.Cell("A1").AddConditionalFormat().WhenEquals("A").Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
        await Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "A" && !v.Value.IsFormula))).IsTrue();
        await Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "A" && !v.Value.IsFormula))).IsTrue();
    }

    [Test]
    public async Task CopyConditionalFormatsFixedStringNum()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "1";
        ws.Cell("B1").Value = "1";
        ws.Cell("A1").AddConditionalFormat().WhenEquals("1").Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
        await Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula))).IsTrue();
        await Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula))).IsTrue();
    }

    [Test]
    public async Task CopyConditionalFormatsRelative()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "1";
        ws.Cell("B1").Value = "1";
        ws.Cell("A1").AddConditionalFormat().WhenEquals("=B1").Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
        await Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "B1" && v.Value.IsFormula))).IsTrue();
        await Assert.That(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "B2" && v.Value.IsFormula))).IsTrue();
    }

    [Test]
    public async Task TestRowCopyContents()
    {
        var workbook = new XLWorkbook();
        var originalSheet = workbook.Worksheets.Add("original");
        var copyRowSheet = workbook.Worksheets.Add("copy row");
        var copyRowAsRangeSheet = workbook.Worksheets.Add("copy row as range");
        var copyRangeSheet = workbook.Worksheets.Add("copy range");

        originalSheet.Cell("A2").SetValue("test value");
        originalSheet.Range("A2:E2").Merge();

        {
            var originalRange = originalSheet.Range("A2:E2");
            var destinationRange = copyRangeSheet.Range("A2:E2");

            originalRange.CopyTo(destinationRange);
        }
        CopyRowAsRange(originalSheet, 2, copyRowAsRangeSheet, 3);
        {
            var originalRow = originalSheet.Row(2);
            var destinationRow = copyRowSheet.Row(2);
            copyRowSheet.Cell("G2").Value = "must be removed after copy";
            originalRow.CopyTo(destinationRow);
        }
        TestHelper.SaveWorkbook(workbook, "Misc", "CopyRowContents.xlsx");

        await Assert.That(copyRangeSheet.Cell("A2").Value).IsEqualTo((XLCellValue)"test value");
        await Assert.That(copyRowSheet.Cell("A2").Value).IsEqualTo((XLCellValue)"test value");
        await Assert.That(copyRangeSheet.Range("A2:E2").IsMerged()).IsTrue();
        await Assert.That(copyRowAsRangeSheet.Cell("A3").Value).IsEqualTo((XLCellValue)"test value");
        await Assert.That(copyRowAsRangeSheet.Range("A3:E3").IsMerged()).IsTrue();
    }

    [Test]
    public async Task UpdateCellsWorksheetTest()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        ws1.Cell(1, 1).Value = "hello, world.";

        var ws2 = ws1.CopyTo("Sheet2");

        await Assert.That(ws1.FirstCell().Address.Worksheet!.Name).IsEqualTo("Sheet1");
        await Assert.That(ws2.FirstCell().Address.Worksheet!.Name).IsEqualTo("Sheet2");
    }

    [Test]
    public async Task CopyHyperlinksAmongSheets()
    {
        using var wb = new XLWorkbook();
        var source = wb.AddWorksheet();
        var target = wb.AddWorksheet();
        source.Cell("A1")
            .SetValue("link")
            .CreateHyperlink()
            .SetValues("https://example.com", "Test tooltip");

        source.Cell("A1").AsRange().CopyTo(target.Cell("B7"));

        var cell = target.Cell("B7");
        await Assert.That(cell.HasHyperlink).IsTrue();
        await Assert.That(cell.GetHyperlink().IsExternal).IsTrue();
        await Assert.That(cell.GetHyperlink().ExternalAddress).IsEqualTo(new Uri("https://example.com"));
        await Assert.That(cell.GetHyperlink().Tooltip).IsEqualTo("Test tooltip");
    }
}
