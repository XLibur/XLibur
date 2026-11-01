using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Exceptions;
using XLibur.Excel.Coordinates;
using XLibur.Tests.Utils;

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

    /// <summary>
    /// A file that holds <c>C1 = LAMBDA(x,x*A1)(B1)</c> with the result 6, as a plain formula or as a
    /// shared formula over C1:C2, the way Excel saves one: <c>_xlfn.</c> and <c>_xlpm.</c> prefixes, a
    /// cached result.
    /// </summary>
    private static MemoryStream FileWithACall(bool shared)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("B1").Value = 3;
        ws.Cell("A2").Value = 4;
        ws.Cell("B2").Value = 5;
        ws.Cell("C1").FormulaA1 = "A1*B1";
        ws.Cell("C2").FormulaA1 = "A2*B2";
        var saved = new MemoryStream();
        wb.SaveAs(saved);

        const string call = "_xlfn.LAMBDA(_xlpm.x,_xlpm.x*A1)(B1)";
        return saved.RewritePart("xl/worksheets/sheet1.xml", xml => shared
            ? xml.Replace("<x:f>A1*B1</x:f>", $"<x:f t=\"shared\" ref=\"C1:C2\" si=\"0\">{call}</x:f><x:v>6</x:v>")
                .Replace("<x:f>A2*B2</x:f>", "<x:f t=\"shared\" si=\"0\"/><x:v>20</x:v>")
            : xml.Replace("<x:f>A1*B1</x:f>", $"<x:f>{call}</x:f><x:v>6</x:v>"));
    }

    [Test]
    [Arguments(false, "A1", "C1")]
    [Arguments(false, "B1", "C1")]
    [Arguments(true, "A1", "C1")]
    [Arguments(true, "B2", "C2")]
    public async Task Changing_a_cell_a_call_reads_makes_its_saved_result_stale(bool shared, string input, string call)
    {
        using var wb = new XLWorkbook(FileWithACall(shared));
        var ws = wb.Worksheet(1);
        await Assert.That(ws.Cell(call).CachedValue.ToString()).IsNotEmpty();

        ws.Cell(input).Value = 7;

        // The call is not evaluated, so a stale result reads as the unsupported feature it is,
        // never as the old value.
        await Assert.That(() => ws.Cell(call).Value).Throws<UnsupportedFeatureException>();
    }

    [Test]
    public async Task Each_cell_of_a_shared_call_reads_its_own_references()
    {
        using var wb = new XLWorkbook(FileWithACall(shared: true));

        await Assert.That(wb.Worksheet(1).Cell("C2").FormulaA1).IsEqualTo("_xlfn.LAMBDA(_xlpm.x,_xlpm.x*A2)(B2)");
    }

    [Test]
    public async Task A_call_whose_argument_names_a_column_with_a_colon_converts_both_ways()
    {
        // The conversion hides the colon from the parser and puts it back in the text it writes.
        const string r1c1 = "LAMBDA(x,x*RC[-1])(Table1[a:b])";
        var origin = new Point(3, 3);

        await Assert.That(FormulaText.TryConvert(r1c1, origin, FormulaNotation.A1, out var a1, out _)).IsTrue();
        await Assert.That(a1).IsEqualTo("LAMBDA(x,x*B3)(Table1[a:b])");
        await Assert.That(FormulaText.TryConvert(a1, origin, FormulaNotation.R1C1, out var back, out _)).IsTrue();
        await Assert.That(back).IsEqualTo(r1c1);
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
