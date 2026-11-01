using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;
/// <summary>
/// Tests that calc engine adjusts its internal state in response to changes of workbook structure.
/// </summary>
internal class CalcEngineListenerTests
{
    [Test]
    public async Task Formulas_dependent_on_specific_sheet_are_dirty_after_sheet_addition()
    {
        using var wb = new XLWorkbook();
        var sutWs = wb.AddWorksheet();
        sutWs.Cell("A1").FormulaA1 = "new!A1";
        await Assert.That(sutWs.Cell("A1").Value).IsEqualTo(XLError.CellReference);

        var newWs = wb.AddWorksheet("new");
        newWs.Cell("A1").Value = 5;

        // Cell contains last calculated value
        await Assert.That(sutWs.Cell("A1").CachedValue).IsEqualTo(XLError.CellReference);

        // But once asked for real value, it calculates it.
        await Assert.That(sutWs.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(sutWs.Cell("A1").Value).IsEqualTo(5.0);
    }

    [Test]
    public async Task Formulas_dependent_on_specific_sheet_are_dirty_after_sheet_deletion()
    {
        using var wb = new XLWorkbook();
        var keptWs = wb.AddWorksheet();
        var deletedWs = wb.AddWorksheet("deleted");

        deletedWs.Cell("A1").Value = 5;
        keptWs.Cell("A1").FormulaA1 = "deleted!A1";
        await Assert.That(keptWs.Cell("A1").Value).IsEqualTo(5.0);

        deletedWs.Delete();

        // Cell contains last calculated value
        await Assert.That(keptWs.Cell("A1").CachedValue).IsEqualTo(5.0);

        // But once asked for real value, it calculates it.
        await Assert.That(keptWs.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(keptWs.Cell("A1").Value).IsEqualTo(XLError.CellReference);
    }

    [Test]
    public async Task Formulas_are_shifted_when_area_is_added_and_cells_shifted_down()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "B1*2";
        ws.Cell("B1").FormulaA1 = "C1*2";
        ws.Cell("C1").FormulaA1 = "1+2";

        ws.RecalculateAllFormulas();

        ws.Range("A1:B1").InsertRowsAbove(2);

        await Assert.That(ws.Cell("A3").Value).IsEqualTo(12.0);
        await Assert.That(ws.Cell("A3").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("B3").NeedsRecalculation).IsFalse();

        // Dependency tree should pick up the change
        ws.Cell("C1").FormulaA1 = "2+2";
        await Assert.That(ws.Cell("A3").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("B3").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("A3").Value).IsEqualTo(16.0);
    }

    [Test]
    public async Task Formulas_are_shifted_when_area_is_added_and_cells_shifted_right()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "A2*2";
        ws.Cell("A2").FormulaA1 = "A3*2";
        ws.Cell("A3").FormulaA1 = "1+2";

        ws.RecalculateAllFormulas();

        ws.Cell("A2").InsertCellsBefore(4);

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(12.0);
        await Assert.That(ws.Cell("E2").NeedsRecalculation).IsFalse();

        // Dependency tree should pick up the change
        ws.Cell("A3").FormulaA1 = "2+2";
        await Assert.That(ws.Cell("E2").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(16.0);
    }

    [Test]
    public async Task Formulas_are_shifted_when_area_is_deleted_and_cells_shifted_up()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A5").FormulaA1 = "1+2";
        ws.Cell("B5").FormulaA1 = "A5*2";
        ws.Cell("C5").FormulaA1 = "B5*2";

        ws.RecalculateAllFormulas();

        ws.Range("B2:C4").Delete(XLShiftDeletedCells.ShiftCellsUp);

        await Assert.That(ws.Cell("C2").Value).IsEqualTo(12.0);
        await Assert.That(ws.Cell("B2").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("A2").NeedsRecalculation).IsFalse();

        // Dependency tree should pick up the change
        ws.Cell("A5").FormulaA1 = "2+2";
        await Assert.That(ws.Cell("B2").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("C2").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("C2").Value).IsEqualTo(16.0);
    }

    [Test]
    public async Task Formulas_are_shifted_when_area_is_deleted_and_cells_shifted_left()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("D1").FormulaA1 = "1+2";
        ws.Cell("E2").FormulaA1 = "D1*2";
        ws.Cell("D3").FormulaA1 = "E2*2";

        ws.RecalculateAllFormulas();

        ws.Range("A1:C5").Delete(XLShiftDeletedCells.ShiftCellsLeft);

        await Assert.That(ws.Cell("A3").Value).IsEqualTo(12.0);
        await Assert.That(ws.Cell("B2").NeedsRecalculation).IsFalse();
        await Assert.That(ws.Cell("A1").NeedsRecalculation).IsFalse();

        // Dependency tree should pick up the change
        ws.Cell("A1").FormulaA1 = "2+2";
        await Assert.That(ws.Cell("B2").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("A3").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("A3").Value).IsEqualTo(16.0);
    }
}
