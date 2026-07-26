using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Ranges;

public class RangeRowCopyToTests
{
    [Test]
    public async Task CopyTo_Cell_CopiesValuesAndStyles()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell("A1").Value = "Hello";
        ws.Cell("B1").Value = 42;
        ws.Cell("C1").Value = true;
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("B1").Style.Fill.BackgroundColor = XLColor.Red;

        var sourceRow = ws.Range("A1:C1").Row(1);
        var result = sourceRow.CopyTo(ws.Cell("A3"));

        await Assert.That(ws.Cell("A3").Value).IsEqualTo((XLCellValue)"Hello");
        await Assert.That(ws.Cell("B3").Value).IsEqualTo((XLCellValue)42);
        await Assert.That(ws.Cell("C3").Value).IsEqualTo(ExpectedCellValue.From(true));
        await Assert.That(ws.Cell("A3").Style.Font.Bold).IsTrue();
        await Assert.That(ws.Cell("B3").Style.Fill.BackgroundColor).IsEqualTo(XLColor.Red);

        // Verify the returned range row covers the correct address
        await Assert.That(result.RangeAddress.FirstAddress.RowNumber).IsEqualTo(3);
        await Assert.That(result.RangeAddress.FirstAddress.ColumnNumber).IsEqualTo(1);
        await Assert.That(result.RangeAddress.LastAddress.ColumnNumber).IsEqualTo(3);
    }

    [Test]
    public async Task CopyTo_Cell_DoesNotModifySource()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell("A1").Value = "Original";
        ws.Cell("B1").Value = 99;

        var sourceRow = ws.Range("A1:B1").Row(1);
        sourceRow.CopyTo(ws.Cell("A3"));

        // Modify the copy
        ws.Cell("A3").Value = "Modified";

        // Source should be unchanged
        await Assert.That(ws.Cell("A1").Value).IsEqualTo((XLCellValue)"Original");
        await Assert.That(ws.Cell("B1").Value).IsEqualTo((XLCellValue)99);
    }

    [Test]
    public async Task CopyTo_Cell_CrossWorksheet()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Source");
        var ws2 = wb.AddWorksheet("Dest");

        ws1.Cell("A1").Value = "Cross-sheet";
        ws1.Cell("B1").Value = 123;
        ws1.Cell("A1").Style.Font.Italic = true;

        var sourceRow = ws1.Range("A1:B1").Row(1);
        var result = sourceRow.CopyTo(ws2.Cell("C5"));

        await Assert.That(ws2.Cell("C5").Value).IsEqualTo((XLCellValue)"Cross-sheet");
        await Assert.That(ws2.Cell("D5").Value).IsEqualTo((XLCellValue)123);
        await Assert.That(ws2.Cell("C5").Style.Font.Italic).IsTrue();
        await Assert.That(result.Worksheet.Name).IsEqualTo("Dest");
    }

    [Test]
    public async Task CopyTo_RangeBase_CopiesValuesAndStyles()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell("A1").Value = "First";
        ws.Cell("B1").Value = "Second";
        ws.Cell("C1").Value = "Third";
        ws.Cell("B1").Style.Font.Underline = XLFontUnderlineValues.Single;

        var sourceRow = ws.Range("A1:C1").Row(1);
        var targetRange = ws.Range("D5:F5");
        var result = sourceRow.CopyTo(targetRange);

        await Assert.That(ws.Cell("D5").Value).IsEqualTo((XLCellValue)"First");
        await Assert.That(ws.Cell("E5").Value).IsEqualTo((XLCellValue)"Second");
        await Assert.That(ws.Cell("F5").Value).IsEqualTo((XLCellValue)"Third");
        await Assert.That(ws.Cell("E5").Style.Font.Underline).IsEqualTo(XLFontUnderlineValues.Single);

        await Assert.That(result.RangeAddress.FirstAddress.RowNumber).IsEqualTo(5);
    }

    [Test]
    public async Task CopyTo_Cell_ReturnsCorrectRangeRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell("B2").Value = 1;
        ws.Cell("C2").Value = 2;
        ws.Cell("D2").Value = 3;

        var sourceRow = ws.Range("B2:D2").Row(1);
        var result = sourceRow.CopyTo(ws.Cell("E10"));

        // Result should be a range row starting at E10 with 3 cells
        await Assert.That(result.CellCount()).IsEqualTo(3);
        await Assert.That(result.Cell(1).Value).IsEqualTo((XLCellValue)1);
        await Assert.That(result.Cell(2).Value).IsEqualTo((XLCellValue)2);
        await Assert.That(result.Cell(3).Value).IsEqualTo((XLCellValue)3);
    }

    [Test]
    public async Task CopyTo_Cell_CopiesFormulas()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell("A1").Value = 10;
        ws.Cell("B1").FormulaA1 = "=A1*2";

        var sourceRow = ws.Range("A1:B1").Row(1);
        sourceRow.CopyTo(ws.Cell("A3"));

        await Assert.That(ws.Cell("A3").Value).IsEqualTo((XLCellValue)10);
        // Formula should be shifted to reference A3
        await Assert.That(ws.Cell("B3").FormulaA1).IsEqualTo("A3*2");
    }

    [Test]
    public async Task CopyTo_Cell_EmptyRowCopiesWithoutError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var sourceRow = ws.Range("A1:C1").Row(1);
        var result = sourceRow.CopyTo(ws.Cell("A5"));

        await Assert.That(ws.Cell("A5").IsEmpty()).IsTrue();
        await Assert.That(ws.Cell("B5").IsEmpty()).IsTrue();
        await Assert.That(ws.Cell("C5").IsEmpty()).IsTrue();
        await Assert.That(result.RangeAddress.FirstAddress.RowNumber).IsEqualTo(5);
    }
}
