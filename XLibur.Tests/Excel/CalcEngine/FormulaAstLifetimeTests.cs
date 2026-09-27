using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// A cell's formula keeps the tree its first evaluation parsed, so a recalculation does not parse
/// it again (#686). The tree must go whenever the formula's text changes in place, or the cell
/// would go on evaluating the formula it used to hold.
/// </summary>
public class FormulaAstLifetimeTests
{
    [Test]
    public async Task EvaluatingACell_KeepsTheTreeOnItsFormula()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 5;
        var cell = (XLCell)ws.Cell("B1");
        cell.FormulaA1 = "A1*2";

        await Assert.That(cell.Formula!.HasAst).IsFalse();

        await Assert.That(cell.Value.GetNumber()).IsEqualTo(10d);
        await Assert.That(cell.Formula!.HasAst).IsTrue();
    }

    [Test]
    public async Task RecalculatingACell_ReusesTheTree()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 5;
        var cell = (XLCell)ws.Cell("B1");
        cell.FormulaA1 = "A1*2";
        await Assert.That(cell.Value.GetNumber()).IsEqualTo(10d);
        var ast = cell.Formula!.GetAst(wb.CalcEngine);

        ws.Cell("A1").Value = 7;

        await Assert.That(cell.Value.GetNumber()).IsEqualTo(14d);
        await Assert.That(cell.Formula!.GetAst(wb.CalcEngine)).IsSameReferenceAs(ast);
    }

    /// <summary>
    /// A rename rewrites the formula's text in place. The kept tree still names the old sheet, so a
    /// sheet added later under that name would be read instead of the renamed one.
    /// </summary>
    [Test]
    public async Task RenamingAReferencedSheet_DropsTheTree()
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet("Data");
        data.Cell("A1").Value = 5;
        var cell = (XLCell)wb.AddWorksheet("Calc").Cell("A1");
        cell.FormulaA1 = "Data!A1*2";
        await Assert.That(cell.Value.GetNumber()).IsEqualTo(10d);

        data.Name = "Source";
        wb.AddWorksheet("Data").Cell("A1").Value = 100;

        await Assert.That(cell.FormulaA1).IsEqualTo("Source!A1*2");
        await Assert.That(cell.Formula!.HasAst).IsFalse();
        await Assert.That(cell.Value.GetNumber()).IsEqualTo(10d);
    }

    /// <summary>
    /// An array formula is one formula for its whole range, so a shift rewrites its text in place
    /// rather than building a new formula. The kept tree still reads the rows it read before.
    /// </summary>
    [Test]
    public async Task ShiftingAnArrayFormula_DropsTheTree()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 2;
        ws.Range("C1:C2").FormulaArrayA1 = "A1:A2*2";
        await Assert.That(ws.Cell("C2").Value.GetNumber()).IsEqualTo(4d);

        ws.Row(1).InsertRowsAbove(1);

        var cell = (XLCell)ws.Cell("C2");
        await Assert.That(cell.FormulaA1).IsEqualTo("A2:A3*2");
        await Assert.That(cell.Formula!.HasAst).IsFalse();
        await Assert.That(cell.Value.GetNumber()).IsEqualTo(2d);
        await Assert.That(ws.Cell("C3").Value.GetNumber()).IsEqualTo(4d);
    }
}
