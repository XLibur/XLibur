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

    #region Helpers

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
