using NUnit.Framework;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

[TestFixture]
public class RangeRowCopyToTests
{
    [Test]
    public void CopyTo_Cell_CopiesValuesAndStyles()
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

        Assert.That(ws.Cell("A3").Value, Is.EqualTo((XLCellValue)"Hello"));
        Assert.That(ws.Cell("B3").Value, Is.EqualTo((XLCellValue)42));
        Assert.That(ws.Cell("C3").Value, Is.EqualTo((XLCellValue)true));
        Assert.That(ws.Cell("A3").Style.Font.Bold, Is.True);
        Assert.That(ws.Cell("B3").Style.Fill.BackgroundColor, Is.EqualTo(XLColor.Red));

        // Verify the returned range row covers the correct address
        Assert.That(result.RangeAddress.FirstAddress.RowNumber, Is.EqualTo(3));
        Assert.That(result.RangeAddress.FirstAddress.ColumnNumber, Is.EqualTo(1));
        Assert.That(result.RangeAddress.LastAddress.ColumnNumber, Is.EqualTo(3));
    }

    [Test]
    public void CopyTo_Cell_DoesNotModifySource()
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
        Assert.That(ws.Cell("A1").Value, Is.EqualTo((XLCellValue)"Original"));
        Assert.That(ws.Cell("B1").Value, Is.EqualTo((XLCellValue)99));
    }

    [Test]
    public void CopyTo_Cell_CrossWorksheet()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Source");
        var ws2 = wb.AddWorksheet("Dest");

        ws1.Cell("A1").Value = "Cross-sheet";
        ws1.Cell("B1").Value = 123;
        ws1.Cell("A1").Style.Font.Italic = true;

        var sourceRow = ws1.Range("A1:B1").Row(1);
        var result = sourceRow.CopyTo(ws2.Cell("C5"));

        Assert.That(ws2.Cell("C5").Value, Is.EqualTo((XLCellValue)"Cross-sheet"));
        Assert.That(ws2.Cell("D5").Value, Is.EqualTo((XLCellValue)123));
        Assert.That(ws2.Cell("C5").Style.Font.Italic, Is.True);
        Assert.That(result.Worksheet.Name, Is.EqualTo("Dest"));
    }

    [Test]
    public void CopyTo_RangeBase_CopiesValuesAndStyles()
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

        Assert.That(ws.Cell("D5").Value, Is.EqualTo((XLCellValue)"First"));
        Assert.That(ws.Cell("E5").Value, Is.EqualTo((XLCellValue)"Second"));
        Assert.That(ws.Cell("F5").Value, Is.EqualTo((XLCellValue)"Third"));
        Assert.That(ws.Cell("E5").Style.Font.Underline, Is.EqualTo(XLFontUnderlineValues.Single));

        Assert.That(result.RangeAddress.FirstAddress.RowNumber, Is.EqualTo(5));
    }

    [Test]
    public void CopyTo_Cell_ReturnsCorrectRangeRow()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell("B2").Value = 1;
        ws.Cell("C2").Value = 2;
        ws.Cell("D2").Value = 3;

        var sourceRow = ws.Range("B2:D2").Row(1);
        var result = sourceRow.CopyTo(ws.Cell("E10"));

        // Result should be a range row starting at E10 with 3 cells
        Assert.That(result.CellCount(), Is.EqualTo(3));
        Assert.That(result.Cell(1).Value, Is.EqualTo((XLCellValue)1));
        Assert.That(result.Cell(2).Value, Is.EqualTo((XLCellValue)2));
        Assert.That(result.Cell(3).Value, Is.EqualTo((XLCellValue)3));
    }

    [Test]
    public void CopyTo_Cell_CopiesFormulas()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell("A1").Value = 10;
        ws.Cell("B1").FormulaA1 = "=A1*2";

        var sourceRow = ws.Range("A1:B1").Row(1);
        sourceRow.CopyTo(ws.Cell("A3"));

        Assert.That(ws.Cell("A3").Value, Is.EqualTo((XLCellValue)10));
        // Formula should be shifted to reference A3
        Assert.That(ws.Cell("B3").FormulaA1, Is.EqualTo("A3*2"));
    }

    [Test]
    public void CopyTo_Cell_EmptyRowCopiesWithoutError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var sourceRow = ws.Range("A1:C1").Row(1);
        var result = sourceRow.CopyTo(ws.Cell("A5"));

        Assert.That(ws.Cell("A5").IsEmpty(), Is.True);
        Assert.That(ws.Cell("B5").IsEmpty(), Is.True);
        Assert.That(ws.Cell("C5").IsEmpty(), Is.True);
        Assert.That(result.RangeAddress.FirstAddress.RowNumber, Is.EqualTo(5));
    }
}
