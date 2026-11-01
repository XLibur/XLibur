using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine.Exceptions;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// A call of the value of an expression, such as <c>LAMBDA(x,x*2)(3)</c>, which the formula parser
/// reads from version 5.1.0. XLibur keeps such a formula and moves its references, but does not
/// evaluate it until LET and LAMBDA are (#847).
/// </summary>
/// <remarks>
/// <c>LAMBDA</c> reads back as <c>_xlfn.LAMBDA</c>, as every function newer than Excel 2007 does in
/// formula text XLibur writes.
/// </remarks>
public class ExpressionCallTests
{
    [Test]
    public async Task A_formula_that_calls_an_expression_is_kept_through_a_save()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C3").FormulaA1 = "LAMBDA(x,x*A1)(B1)";
        using var saved = new MemoryStream();
        wb.SaveAs(saved);
        using var reloaded = new XLWorkbook(saved);

        await Assert.That(reloaded.Worksheet(1).Cell("C3").FormulaA1).IsEqualTo("_xlfn.LAMBDA(x,x*A1)(B1)");
    }

    [Test]
    public async Task The_references_of_a_call_convert_to_R1C1()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C3").FormulaA1 = "LAMBDA(x,x*A1)(B1)";

        await Assert.That(ws.Cell("C3").FormulaR1C1).IsEqualTo("_xlfn.LAMBDA(x,x*R[-2]C[-2])(R[-2]C[-1])");
    }

    [Test]
    public async Task The_references_of_a_call_move_when_rows_are_inserted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C3").FormulaA1 = "LAMBDA(x,x*A1)(B1)";

        ws.Row(1).InsertRowsAbove(2);

        await Assert.That(ws.Cell("C5").FormulaA1).IsEqualTo("_xlfn.LAMBDA(x,x*A3)(B3)");
    }

    [Test]
    public async Task A_copied_call_is_written_for_its_new_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C3").FormulaA1 = "LAMBDA(x,x*A1)(B1)";

        ws.Cell("C3").CopyTo(ws.Cell("D4"));

        await Assert.That(ws.Cell("D4").FormulaA1).IsEqualTo("_xlfn.LAMBDA(x,x*B2)(C2)");
    }

    [Test]
    public async Task Evaluating_a_call_reports_it_as_not_supported()
    {
        // As for any construct XLibur reads but does not evaluate, such as DDE: an unsupported feature,
        // where before parser 5.1.0 the text was refused as unreadable.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "LAMBDA(x,x*2)(3)";

        await Assert.That(() => ws.Cell("A1").Value).Throws<UnsupportedFeatureException>()
            .WithMessage("Evaluation of a call of the value of an expression is not implemented.");
    }
}
