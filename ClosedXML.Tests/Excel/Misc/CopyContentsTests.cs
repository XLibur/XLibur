using System;
using ClosedXML.Excel;
using NUnit.Framework;
using System.Linq;

namespace ClosedXML.Tests.Excel.Misc;

[TestFixture]
public class CopyContentsTests
{
    private static void CopyRowAsRange(IXLWorksheet originalSheet, int originalRowNumber, IXLWorksheet destSheet,
        int destRowNumber)
    {
        var destinationRow = destSheet.Row(destRowNumber);
        destinationRow.Clear();

        var originalRow = originalSheet.Row(originalRowNumber);
        var columnNumber = originalRow.LastCellUsed(XLCellsUsedOptions.All).Address.ColumnNumber;

        var originalRange = originalSheet.Range(originalRowNumber, 1, originalRowNumber, columnNumber);
        var destRange = destSheet.Range(destRowNumber, 1, destRowNumber, columnNumber);
        originalRange.CopyTo(destRange);
    }

    [Test]
    public void CopyConditionalFormatsCount()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.FirstCell().AddConditionalFormat().WhenContains("1").Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.FirstCell().AsRange());
        Assert.AreEqual(2, ws.ConditionalFormats.Count());
    }

    [Test]
    public void CopyConditionalFormatsFixedNum()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "1";
        ws.Cell("B1").Value = "1";
        ws.Cell("A1").AddConditionalFormat().WhenEquals(1).Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
        Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)));
        Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)));
    }

    [Test]
    public void CopyConditionalFormatsFixedString()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "A";
        ws.Cell("B1").Value = "B";
        ws.Cell("A1").AddConditionalFormat().WhenEquals("A").Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
        Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "A" && !v.Value.IsFormula)));
        Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "A" && !v.Value.IsFormula)));
    }

    [Test]
    public void CopyConditionalFormatsFixedStringNum()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "1";
        ws.Cell("B1").Value = "1";
        ws.Cell("A1").AddConditionalFormat().WhenEquals("1").Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
        Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)));
        Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)));
    }

    [Test]
    public void CopyConditionalFormatsRelative()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = "1";
        ws.Cell("B1").Value = "1";
        ws.Cell("A1").AddConditionalFormat().WhenEquals("=B1").Fill.SetBackgroundColor(XLColor.Blue);
        ws.Cell("A2").CopyFrom(ws.Cell("A1").AsRange());
        Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "B1" && v.Value.IsFormula)));
        Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "B2" && v.Value.IsFormula)));
    }

    [Test]
    public void TestRowCopyContents()
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
    }

    [Test]
    public void UpdateCellsWorksheetTest()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        ws1.Cell(1, 1).Value = "hello, world.";

        var ws2 = ws1.CopyTo("Sheet2");

        Assert.AreEqual("Sheet1", ws1.FirstCell().Address.Worksheet.Name);
        Assert.AreEqual("Sheet2", ws2.FirstCell().Address.Worksheet.Name);
    }

    [Test]
    public void CopyHyperlinksAmongSheets()
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
        Assert.True(cell.HasHyperlink);
        Assert.True(cell.GetHyperlink().IsExternal);
        Assert.AreEqual(new Uri("https://example.com"), cell.GetHyperlink().ExternalAddress);
        Assert.AreEqual("Test tooltip", cell.GetHyperlink().Tooltip);
    }
}
