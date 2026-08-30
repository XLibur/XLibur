using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.Coordinates;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Spec 40 — the dependency-tree walk uses a formula's dirty flag both to mean "needs
/// recalculation" and, internally, "already visited by this walk". Anything other than the walk
/// itself that dirties a formula (<see cref="IXLCell.InvalidateFormula"/>, a sheet rename, a
/// row/column insert, a range move) makes the walk mistake "already dirty for an unrelated
/// reason" for "already visited", stop, and prune everything downstream of that node.
/// </summary>
/// <remarks>
/// Every interference test is paired with a control: the same graph and the same edit, without
/// the interfering operation. The control proves the graph and edit are capable of producing the
/// right answer on their own, so a failure in the interference test can only be attributed to the
/// interference.
/// </remarks>
internal class DirtyPropagationTests
{
    #region Cycle termination (safety net, written and seen green before the fix)

    /// <summary>
    /// The dirty flag used to double as the walk's cycle guard: a node already marked dirty by
    /// the walk itself was not re-enqueued, which is also what stopped the walk looping forever
    /// around a cycle. Separating "visited" from "dirty" must not reopen the cycle as an infinite
    /// loop, and a formula strictly downstream of the cycle must still be reached.
    /// </summary>
    /// <remarks>
    /// A genuine cycle cannot be evaluated to a value through the public API without throwing
    /// (<c>XLCalculationChain</c> detects it and throws once evaluation is attempted), so this
    /// exercises <see cref="DependencyTree.MarkDirty"/> directly, the same seam
    /// <c>DependencyTreeTests</c> already uses for its cycle coverage.
    /// </remarks>
    [Test]
    public async Task Mark_dirty_terminates_on_cycle_and_reaches_tail_without_interference()
    {
        using var wb = new XLWorkbook();
        var tree = new DependencyTree();
        var ws = wb.AddWorksheet();
        tree.AddSheetTree(ws);
        AddFormula(tree, ws, "B1", "=D1 + A1");
        AddFormula(tree, ws, "C1", "=B1");
        AddFormula(tree, ws, "D1", "=C1"); // B1 -> C1 -> D1 -> B1 is a cycle
        AddFormula(tree, ws, "E1", "=D1"); // tail hanging off the cycle

        MarkDirty(tree, ws, "A1");

        await AssertDirty(ws, "B1", "C1", "D1", "E1");
    }

    #endregion

    #region Interference matrix (centrepiece; red before the fix)

    // Chain used by every case below: A1 = 1, B1 = A1 + 1, C1 = B1 + 1, D1 = C1 + 1.
    // Each case dirties C1 (the intermediate node) through something other than the walk, then
    // edits A1 (the root), and asserts the whole chain recalculates.

    [Test]
    public async Task InvalidateFormula_on_an_intermediate_cell_does_not_prune_a_later_edit()
    {
        using var wb = new XLWorkbook();
        var ws = BuildChain(wb);

        ws.Cell("C1").InvalidateFormula();
        ws.Cell("A1").Value = 10;

        await AssertChain(ws, a1: 10, b1: 11, c1: 12, d1: 13);
    }

    [Test]
    public async Task InvalidateFormula_control_without_interference()
    {
        using var wb = new XLWorkbook();
        var ws = BuildChain(wb);

        ws.Cell("A1").Value = 10;

        await AssertChain(ws, a1: 10, b1: 11, c1: 12, d1: 13);
    }

    /// <summary>
    /// A live rename goes through <c>XLWorksheets.Rename</c>, which re-keys the sheet by removing
    /// and re-adding it — the same path a newly added sheet takes — so it also fires
    /// <c>XLCalcEngine.OnAddedSheet</c>, which purges the whole dependency tree and marks every
    /// formula in the workbook dirty. That blanket invalidation is a separate, coarser mechanism
    /// than the one this spec fixes, and it happens to make every formula on the sheet dirty
    /// regardless of whether the rename actually touched its text — so this test is expected to
    /// pass both before and after the fix. <see cref="Formula_marked_dirty_by_a_text_rewrite_does_not_prune_a_later_edit"/>
    /// isolates the walk from that purge and is the test that actually goes red beforehand.
    /// </summary>
    [Test]
    public async Task Sheet_rename_does_not_prune_a_later_edit()
    {
        using var wb = new XLWorkbook();
        var ws = BuildChain(wb, "Data", intermediateReferencesOwnSheetByName: true);

        ws.Name = "Renamed";
        ws.Cell("A1").Value = 10;

        await AssertChain(ws, a1: 10, b1: 11, c1: 12, d1: 13);
    }

    [Test]
    public async Task Sheet_rename_control_without_interference()
    {
        using var wb = new XLWorkbook();
        var ws = BuildChain(wb, "Data", intermediateReferencesOwnSheetByName: true);

        ws.Cell("A1").Value = 10;

        await AssertChain(ws, a1: 10, b1: 11, c1: 12, d1: 13);
    }

    /// <summary>
    /// A row/column insert always goes through <c>XLCalcEngine.OnInsertAreaAndShiftDown</c> /
    /// <c>OnInsertAreaAndShiftRight</c>, which purges the whole dependency tree and marks every
    /// formula in the workbook dirty — the same blanket mechanism <see cref="Sheet_rename_does_not_prune_a_later_edit"/>
    /// documents. It happens to dirty C1 (and B1 and D1) regardless of whether the shift touched
    /// their text, so this test is expected to pass both before and after the fix.
    /// <see cref="Formula_marked_dirty_by_a_text_rewrite_does_not_prune_a_later_edit"/> isolates
    /// the walk from that purge and is the test that actually goes red beforehand.
    /// </summary>
    [Test]
    public async Task Row_insert_shifting_an_intermediate_reference_does_not_prune_a_later_edit()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        // C1 references a far-away helper cell so a row insert well below the chain shifts only
        // that reference, forcing C1's formula text to be rewritten (and the formula marked
        // dirty) without touching A1/B1/D1's text.
        ws.Cell("A1").Value = 1;
        ws.Cell("B1").FormulaA1 = "A1+1";
        ws.Cell("C1").FormulaA1 = "B1+1+Z100";
        ws.Cell("D1").FormulaA1 = "C1+1";
        ws.Cell("Z100").Value = 0;
        await ForceEvaluation(ws);

        ws.Row(50).InsertRowsAbove(1); // shifts Z100 -> Z101, rewrites C1's formula text
        ws.Cell("A1").Value = 10;

        await AssertChain(ws, a1: 10, b1: 11, c1: 12, d1: 13);
    }

    [Test]
    public async Task Row_insert_control_without_interference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell("A1").Value = 1;
        ws.Cell("B1").FormulaA1 = "A1+1";
        ws.Cell("C1").FormulaA1 = "B1+1+Z100";
        ws.Cell("D1").FormulaA1 = "C1+1";
        ws.Cell("Z100").Value = 0;
        await ForceEvaluation(ws);

        ws.Cell("A1").Value = 10;

        await AssertChain(ws, a1: 10, b1: 11, c1: 12, d1: 13);
    }

    /// <summary>
    /// The primitive both a sheet rename (<see cref="XLCellFormula.RenameSheet"/>) and a
    /// row/column insert's reference shift (<see cref="XLCellFormula.UpdateShiftedA1"/>) use to
    /// dirty a formula whose text changed: <see cref="XLCellFormula.MarkExplicitlyDirty"/>,
    /// called directly on the formula, independent of the walk. Exercising it directly, without
    /// going through the live operations, isolates it from the whole-workbook purge those two
    /// operations also happen to trigger today (see <see cref="Sheet_rename_does_not_prune_a_later_edit"/>
    /// and <see cref="Row_insert_shifting_an_intermediate_reference_does_not_prune_a_later_edit"/>),
    /// which would otherwise mask this defect for both of them.
    /// </summary>
    [Test]
    public async Task Formula_marked_dirty_by_a_text_rewrite_does_not_prune_a_later_edit()
    {
        using var wb = new XLWorkbook();
        var ws = BuildChain(wb);

        ((XLCell)ws.Cell("C1")).Formula!.MarkExplicitlyDirty();
        ws.Cell("A1").Value = 10;

        await AssertChain(ws, a1: 10, b1: 11, c1: 12, d1: 13);
    }

    [Test]
    public async Task Formula_marked_dirty_by_a_text_rewrite_control_without_interference()
    {
        using var wb = new XLWorkbook();
        var ws = BuildChain(wb);

        ws.Cell("A1").Value = 10;

        await AssertChain(ws, a1: 10, b1: 11, c1: 12, d1: 13);
    }

    [Test]
    public async Task Range_move_of_an_intermediate_formula_does_not_prune_a_later_edit()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell("A1").Value = 1;
        // D3 is the intermediate node; a transpose of the square range C3:D4 swaps it with its
        // off-diagonal partner C4. E1 is written against that destination address up front, so
        // the chain flows through wherever D3 lands.
        ws.Cell("D3").FormulaA1 = "$A$1+1";
        ws.Cell("E1").FormulaA1 = "C4+1";
        await ForceEvaluation(ws);

        // A square range whose first row number equals its first column number (3 and 3) is
        // used deliberately: XLRange.TransposeRange computes a swapped-off-diagonal cell's new
        // address as Point(col + colOffset, row + rowOffset) instead of the axis-correct
        // Point(col + rowOffset, row + colOffset), so a range whose row/column start numbers
        // differ (e.g. B1:C2) transposes to the wrong cell. That is a pre-existing defect in
        // Transpose, outside spec 40's scope, reported separately; picking equal offsets here
        // keeps this test from depending on it.
        ws.Range("C3:D4").Transpose(XLTransposeOptions.MoveCells);

        ws.Cell("A1").Value = 10;

        await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(10);
        await Assert.That((double)ws.Cell("C4").Value).IsEqualTo(11);
        await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(12);
    }

    [Test]
    public async Task Range_move_control_without_interference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell("A1").Value = 1;
        ws.Cell("C4").FormulaA1 = "$A$1+1";
        ws.Cell("E1").FormulaA1 = "C4+1";
        await ForceEvaluation(ws);

        ws.Cell("A1").Value = 10;

        await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(10);
        await Assert.That((double)ws.Cell("C4").Value).IsEqualTo(11);
        await Assert.That((double)ws.Cell("E1").Value).IsEqualTo(12);
    }

    #endregion

    #region Helpers

    private static IXLWorksheet BuildChain(XLWorkbook wb, string sheetName = "Sheet1", bool intermediateReferencesOwnSheetByName = false)
    {
        var ws = wb.AddWorksheet(sheetName);
        ws.Cell("A1").Value = 1;
        ws.Cell("B1").FormulaA1 = "A1+1";
        ws.Cell("C1").FormulaA1 = intermediateReferencesOwnSheetByName ? $"{sheetName}!B1+1" : "B1+1";
        ws.Cell("D1").FormulaA1 = "C1+1";

        ForceEvaluation(ws).GetAwaiter().GetResult();
        return ws;
    }

    /// <summary>
    /// Reads every used formula cell's value once, so the calc chain and dependency tree exist
    /// and every formula starts clean before a test applies its interference.
    /// </summary>
    private static async Task ForceEvaluation(IXLWorksheet ws)
    {
        foreach (var cell in ws.CellsUsed(c => c.HasFormula))
        {
            _ = cell.Value;
        }

        await Task.CompletedTask;
    }

    private static async Task AssertChain(IXLWorksheet ws, double a1, double b1, double c1, double d1)
    {
        await Assert.That((double)ws.Cell("A1").Value).IsEqualTo(a1);
        await Assert.That((double)ws.Cell("B1").Value).IsEqualTo(b1);
        await Assert.That((double)ws.Cell("C1").Value).IsEqualTo(c1);
        await Assert.That((double)ws.Cell("D1").Value).IsEqualTo(d1);
    }

    private static XLCellFormula AddFormula(DependencyTree tree, IXLWorksheet sheet, string address, string formula)
    {
        var cell = (XLCell)sheet.Cell(address);
        cell.Formula = XLCellFormula.NormalA1(formula);
        cell.Formula.MarkClean(((XLWorksheet)sheet).Workbook);
        var cellArea = new SheetArea(sheet.Name, new Area(cell.SheetPoint, cell.SheetPoint));
        tree.AddFormula(cellArea, cell.Formula, sheet.Workbook);
        return cell.Formula;
    }

    private static void MarkDirty(DependencyTree tree, IXLWorksheet sheet, string range)
    {
        var area = new SheetArea(sheet.Name, Area.Parse(range));
        tree.MarkDirty(area);
    }

    private static async Task AssertDirty(IXLWorksheet sheet, params string[] dirtyRanges)
    {
        var ws = (XLWorksheet)sheet;
        var wb = ws.Workbook;
        foreach (var dirtyRange in dirtyRanges)
        {
            foreach (var dirtyCell in ws.Cells(dirtyRange))
            {
                await Assert.That(dirtyCell.Formula).IsNotNull();
                await Assert.That(dirtyCell.Formula!.IsDirty(wb)).IsTrue();
            }
        }
    }

    #endregion
}
