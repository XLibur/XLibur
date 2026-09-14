using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// A refused formula is an expected failure (spec 56, Q22). It fails only the cell that holds it and
/// the formulas that depend on it, and blocks no write elsewhere in the workbook (#489).
/// </summary>
public class RefusedAndUnsupportedFormulaTests
{
    /// <summary>
    /// The form the formula bar shows for an external reference. <see cref="IXLCell.FormulaA1"/>
    /// accepts it, and the parser refuses it.
    /// </summary>
    private const string Refused = "'[Book2.xlsx]Sheet1'!A1";

    /// <summary>
    /// #489. A successful evaluation tells the calc engine to build its dependency tree on the next
    /// write. The build parsed every formula in the workbook and threw on the refused one, so every
    /// later write anywhere threw <see cref="ExpressionParseException"/>, after it had been applied.
    /// </summary>
    [Test]
    public async Task Issue489_a_write_after_a_read_does_not_throw()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = Refused;
        ws.Cell("B1").FormulaA1 = "1+1";
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(2);

        ws.Cell("C1").Value = 5;

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(5);
        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<ExpressionParseException>();
    }

    /// <summary>
    /// #489. The tree the write builds still does its job: the write reaches the formulas that
    /// depend on it.
    /// </summary>
    [Test]
    public async Task Issue489_a_write_still_marks_its_dependents_dirty()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = Refused;
        ws.Cell("C1").Value = 1;
        ws.Cell("D1").FormulaA1 = "C1*2";
        await Assert.That(ws.Cell("D1").Value).IsEqualTo(2);

        ws.Cell("C1").Value = 5;

        await Assert.That(ws.Cell("D1").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("D1").Value).IsEqualTo(10);
    }

    /// <summary>
    /// #489. The tree covers the whole workbook, so a refused formula on one sheet blocked writes on
    /// every other.
    /// </summary>
    [Test]
    public async Task Issue489_a_refused_formula_on_another_sheet_does_not_block_a_write()
    {
        using var wb = new XLWorkbook();
        var sheet1 = wb.AddWorksheet("Sheet1");
        var sheet2 = wb.AddWorksheet("Sheet2");
        sheet2.Cell("A1").FormulaA1 = Refused;
        sheet1.Cell("B1").FormulaA1 = "1+1";
        await Assert.That(sheet1.Cell("B1").Value).IsEqualTo(2);

        sheet1.Cell("C1").Value = 5;

        await Assert.That(sheet1.Cell("C1").Value).IsEqualTo(5);
        await Assert.That(() => _ = sheet2.Cell("A1").Value).Throws<ExpressionParseException>();
    }

    /// <summary>
    /// #489. The obvious recovery threw too: setting a value on the refused cell marks the cell dirty
    /// before it clears the formula.
    /// </summary>
    [Test]
    public async Task Issue489_setting_a_value_on_the_refused_cell_recovers_it()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = Refused;
        ws.Cell("B1").FormulaA1 = "1+1";
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(2);

        ws.Cell("A1").Value = 0;

        await Assert.That(ws.Cell("A1").HasFormula).IsFalse();
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(0);
        ws.Cell("C1").Value = 5;
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(5);
    }

    /// <summary>
    /// #489. The writes were half-applied: the formula set on C1 was stored and the value set on C2
    /// was cached, yet each call threw. Both now complete, and the value reaches its dependent.
    /// </summary>
    [Test]
    public async Task Issue489_a_formula_write_and_a_value_write_complete()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = Refused;
        ws.Cell("C2").Value = 1;
        ws.Cell("D2").FormulaA1 = "C2+1";
        await Assert.That(ws.Cell("D2").Value).IsEqualTo(2);

        ws.Cell("C1").FormulaA1 = "2*3";
        ws.Cell("C2").Value = 7;

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(6);
        await Assert.That(ws.Cell("D2").NeedsRecalculation).IsTrue();
        await Assert.That(ws.Cell("D2").Value).IsEqualTo(8);
    }

    /// <summary>
    /// #489, by a second road. Once a recalculation has built the tree, each formula set afterwards
    /// is parsed for its precedents, so a refused formula threw from its own setter, after it had been
    /// stored.
    /// </summary>
    [Test]
    public async Task Issue489_a_refused_formula_can_be_set_after_a_recalculation()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("B1").FormulaA1 = "1+1";
        wb.RecalculateAllFormulas();

        ws.Cell("A1").FormulaA1 = Refused;

        await Assert.That(ws.Cell("A1").FormulaA1).IsEqualTo(Refused);
        ws.Cell("C1").Value = 5;
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(5);
        await Assert.That(() => _ = ws.Cell("A1").Value).Throws<ExpressionParseException>();
    }
}
