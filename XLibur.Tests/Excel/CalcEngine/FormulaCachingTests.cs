using XLibur.Excel;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

public class FormulaCachingTests
{
    [Test]
    public async Task StaticCellDoesNotNeedRecalculation()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var cell = sheet.Cell(1, 1);
        cell.Value = "1234567";

        await Assert.That(cell.NeedsRecalculation).IsFalse();
    }

    [Test]
    public async Task EditCellInvalidatesDependentCells()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var cell = sheet.Cell(1, 1);
        var dependentCell = sheet.Cell(2, 1);
        dependentCell.FormulaA1 = "=A1";
        var _ = dependentCell.Value;

        cell.Value = "1234567";

        await Assert.That(dependentCell.NeedsRecalculation).IsTrue();
    }

    [Test]
    public async Task EditFormulaA1InvalidatesDependentCells()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a1 = sheet.Cell("A1");
        var a2 = sheet.Cell("A2");
        var a3 = sheet.Cell("A3");
        var a4 = sheet.Cell("A4");
        a2.FormulaA1 = "=A1*10";
        a3.FormulaA1 = "=A2*10";
        a4.FormulaA1 = "=SUM(A1:A3)";
        a1.Value = 15;

        var res1 = a4.Value;
        a2.FormulaA1 = "=A1*20";
        var res2 = a4.Value;

        await Assert.That(res1).IsEqualTo(15 + 150 + 1500);
        await Assert.That(res2).IsEqualTo(15 + 300 + 3000);
    }

    [Test]
    public async Task EditFormulaR1C1InvalidatesDependentCells()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a1 = sheet.Cell("A1");
        var a2 = sheet.Cell("A2");
        var a3 = sheet.Cell("A3");
        var a4 = sheet.Cell("A4");
        a2.FormulaA1 = "=A1*10";
        a3.FormulaA1 = "=A2*10";
        a4.FormulaA1 = "=SUM(A1:A3)";
        a1.Value = 15;

        var res1 = a4.Value;
        a2.FormulaR1C1 = "=R[-1]C*2";
        var res2 = a4.Value;

        await Assert.That(res1).IsEqualTo(15 + 150 + 1500);
        await Assert.That(res2).IsEqualTo(15 + 30 + 300);
    }

    [Test]
    public async Task InsertRowInvalidatesValues()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a4 = sheet.Cell("A4");
        a4.FormulaA1 = "=COUNTBLANK(A1:A3)";

        await Assert.That(a4.Value).IsEqualTo(3);

        sheet.Row(2).InsertRowsAbove(2);

        await Assert.That(sheet.Cell("A6").Value).IsEqualTo(5);
    }

    [Test]
    public async Task DeleteRowModifiesFormulaAndInvalidatesValues()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var original = sheet.Cell("A4");
        original.FormulaA1 = "=COUNTBLANK(A1:A3)";

        await Assert.That(original.Value).IsEqualTo(3);

        sheet.Row(2).Delete();

        var shifted = sheet.Cell("A3");
        await Assert.That(shifted.FormulaA1).IsEqualTo("COUNTBLANK(A1:A2)");
        await Assert.That(shifted.Value).IsEqualTo(2);
    }

    [Test]
    public async Task ChainedCalculationPreservesIntermediateValues()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a1 = sheet.Cell("A1");
        var a2 = sheet.Cell("A2");
        var a3 = sheet.Cell("A3");
        var a4 = sheet.Cell("A4");
        a2.FormulaA1 = "=A1*10";
        a3.FormulaA1 = "=A2*10";
        a4.FormulaA1 = "=SUM(A1:A3)";

        a1.Value = 15;
        var res = a4.Value;

        await Assert.That(res).IsEqualTo(15 + 150 + 1500);
        await Assert.That(a4.NeedsRecalculation).IsFalse();
        await Assert.That(a3.NeedsRecalculation).IsFalse();
        await Assert.That(a2.NeedsRecalculation).IsFalse();
        await Assert.That(a2.CachedValue).IsEqualTo(150);
        await Assert.That(a3.CachedValue).IsEqualTo(1500);
        await Assert.That(a4.CachedValue).IsEqualTo(15 + 150 + 1500);
    }

    [Test]
    public async Task EditingAffectsDependentCells()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a1 = sheet.Cell("A1");
        var a2 = sheet.Cell("A2");
        var a3 = sheet.Cell("A3");
        var a4 = sheet.Cell("A4");
        a2.FormulaA1 = "=A1*10";
        a3.FormulaA1 = "=A2*10";
        a4.FormulaA1 = "=SUM(A1:A3)";
        a1.Value = 15;

        var res1 = a4.Value;
        a1.Value = 20;
        var res2 = a4.Value;

        await Assert.That(res1).IsEqualTo(15 + 150 + 1500);
        await Assert.That(res2).IsEqualTo(20 + 200 + 2000);
    }

    [Test]
    [Arguments("C4", new[] { "C5" })]
    [Arguments("D4", new string[] { })]
    [Arguments("A1", new[] { "A2", "A3", "A4", "C1", "C2", "C3", "C5" })]
    [Arguments("B2", new[] { "B3", "B4", "C2", "C3", "C5" })]
    [Arguments("C2", new[] { "C5" })]
    public async Task EditingDoesNotAffectNonDependingCells(string changedCell, string[] affectedCells)
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        sheet.Cell("A2").FormulaA1 = "A1+1";
        sheet.Cell("A3").FormulaA1 = "SUM(A1:A2)";
        sheet.Cell("A4").FormulaA1 = "SUM(A1:A3)";
        sheet.Cell("B2").FormulaA1 = "B1+1";
        sheet.Cell("B3").FormulaA1 = "SUM(B1:B2)";
        sheet.Cell("B4").FormulaA1 = "SUM(B1:B3)";
        sheet.Cell("C1").FormulaA1 = "SUM(A1:B1)";
        sheet.Cell("C2").FormulaA1 = "SUM(A2:B2)";
        sheet.Cell("C3").FormulaA1 = "SUM(A3:B3)";
        sheet.Cell("C5").FormulaA1 = "SUM($A$1:$C$4)";
        sheet.RecalculateAllFormulas();
        var allCells = sheet.CellsUsed();

        sheet.Cell(changedCell).Value = 100;
        var modifiedCells = allCells.Where(cell => cell.NeedsRecalculation);

        var xlCells = modifiedCells as IXLCell[] ?? modifiedCells.ToArray();
        await Assert.That(xlCells.Length).IsEqualTo(affectedCells.Length);
        foreach (var cellAddress in affectedCells)
        {
            await Assert.That(xlCells.Any(cell => cell.Address.ToString() == cellAddress)).IsTrue().Because($"Cell {cellAddress} is expected to need recalculation, but it does not");
        }
    }

    [Test]
    public async Task CircularReferenceFailsCalculating()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a1 = sheet.Cell("A1");
        var a2 = sheet.Cell("A2");
        var a3 = sheet.Cell("A3");
        var a4 = sheet.Cell("A4");

        a2.FormulaA1 = "=A1*10";
        a3.FormulaA1 = "=A2*10";
        a4.FormulaA1 = "=A3*10";
        a1.FormulaA1 = "A2+A3+A4";

        var getValueA1 = new Action(() => { _ = a1.Value; });
        var getValueA2 = new Action(() => { _ = a2.Value; });
        var getValueA3 = new Action(() => { _ = a3.Value; });
        var getValueA4 = new Action(() => { _ = a4.Value; });

        await Assert.That(getValueA1).Throws<InvalidOperationException>();
        await Assert.That(getValueA2).Throws<InvalidOperationException>();
        await Assert.That(getValueA3).Throws<InvalidOperationException>();
        await Assert.That(getValueA4).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task CircularReferenceRecalculationNeededDoesNotFail()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.Worksheets.Add("TestSheet");
        var a1 = sheet.Cell("A1");
        var a2 = sheet.Cell("A2");
        var a3 = sheet.Cell("A3");
        var a4 = sheet.Cell("A4");

        a2.FormulaA1 = "=A1*10";
        a3.FormulaA1 = "=A2*10";
        a4.FormulaA1 = "=A3*10";
        var _ = a4.Value;
        a1.FormulaA1 = "=SUM(A2:A4)";

        var recalcNeededA1 = a1.NeedsRecalculation;
        var recalcNeededA2 = a2.NeedsRecalculation;
        var recalcNeededA3 = a3.NeedsRecalculation;
        var recalcNeededA4 = a4.NeedsRecalculation;

        await Assert.That(recalcNeededA1).IsTrue();
        await Assert.That(recalcNeededA2).IsTrue();
        await Assert.That(recalcNeededA3).IsTrue();
        await Assert.That(recalcNeededA4).IsTrue();
    }

    [Test]
    public async Task DeleteWorksheetInvalidatesValues()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.Worksheets.Add("Sheet1");
        var sheet2 = wb.Worksheets.Add("Sheet2");
        var sheet1A1 = sheet1.Cell("A1");
        var sheet2A1 = sheet2.Cell("A1");
        sheet1A1.FormulaA1 = "Sheet2!A1";
        sheet2A1.Value = "TestValue";

        var valueBeforeDeletion = sheet1A1.Value;
        sheet2.Delete();
        var valueAfterDeletion = sheet1A1.Value;

        await Assert.That(valueBeforeDeletion).IsEqualTo("TestValue");
        await Assert.That(valueAfterDeletion).IsEqualTo(XLError.CellReference);
    }

    [Test]
    public async Task CachedValueToExternalWorkbook()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\ExternalLinks\WorkbookWithExternalLink.xlsx"));
        using var wb = new XLWorkbook(stream);
        var ws = wb.Worksheets.First();
        var cell = ws.Cell("B2");
        await Assert.That(cell.NeedsRecalculation).IsFalse();
        await Assert.That(cell.HasFormula).IsTrue();

        // This will fail when we start supporting external links
        await Assert.That(cell.FormulaA1.StartsWith("[1]")).IsTrue();

        await Assert.That(cell.CachedValue).IsEqualTo("hello world");
        await Assert.That(cell.Value).IsEqualTo("hello world");

        await Assert.That(ws.Evaluate("LEN(B2)")).IsEqualTo(11);

        // External file references to evaluate to #REF! instead of throwing
        await Assert.That(wb.RecalculateAllFormulas).ThrowsNothing();
        await Assert.That(cell.Value).IsEqualTo(XLError.CellReference);
    }

    [Test]
    public async Task ChangingValueChangesCachedValue()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test");
        var cell = ws.Cell(1, 1);

        cell.Value = "Hello";
        await Assert.That(cell.CachedValue).IsEqualTo("Hello");

        cell.Value = 74.0;
        await Assert.That(cell.CachedValue).IsEqualTo(74.0);

        cell.Value = new DateTime(2019, 1, 1, 14, 0, 0, DateTimeKind.Unspecified);
        await Assert.That(cell.CachedValue).IsEqualTo(new DateTime(2019, 1, 1, 14, 0, 0, DateTimeKind.Unspecified));
    }
}
