using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// <c>LET</c>: names bound in order, each seen by the values after it and by the body, before any
/// defined name of the same spelling (#847, spec 08 phase 1).
/// </summary>
/// <remarks>
/// Every expected value is what Excel 365 calculated for the same formula, typed into a cell over
/// COM. The workbook it saved is <see cref="Fixture"/>: A1:A3 hold 1, 2 and 3, and the defined
/// names <c>x</c> and <c>Ten</c> both refer to 10.
/// </remarks>
public class LetTests
{
    private const string Fixture = @"Other\ExcelAuthored\LetFormulas.xlsx";

    /// <summary>
    /// The cells of the fixture whose formulas Excel calculated. C14 is left out: its result spills,
    /// and the formula below it blocks the spill. C8, C19 and C22 Excel refused to enter.
    /// </summary>
    public static IEnumerable<string> CalculatedCells() =>
        new[] { 1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 15, 16, 17, 18, 20, 21 }.Select(row => $"C{row}");

    [Test]
    [MethodDataSource(nameof(CalculatedCells))]
    public async Task A_workbook_Excel_wrote_recalculates_to_the_values_Excel_saved(string address)
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var cell = wb.Worksheet(1).Cell(address);
        var cached = cell.CachedValue;

        wb.RecalculateAllFormulas();

        await Assert.That(cell.Value).IsEqualTo(cached).Because(cell.FormulaA1);
    }

    [Test]
    [Arguments("LET(x,2,y,x*3,x+y)", 8)]
    [Arguments("LET(X,1,x+1)", 2)]
    [Arguments("LET(x,1,LET(y,x+1,y*2))", 4)]
    [Arguments("LET(x,1,LET(x,5,x)+x)", 6)]
    [Arguments("LET(r,A1:A3,SUM(r))", 6)]
    [Arguments("LET(r,A1:A3,INDEX(r,2))", 2)]
    [Arguments("LET(r,A1:A3,ROWS(r))", 3)]
    [Arguments("LET(x,1/0,5)", 5)]
    public async Task A_typed_LET_evaluates_as_Excel_does(string formula, double expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 2;
        ws.Cell("A3").Value = 3;
        ws.Cell("C1").FormulaA1 = formula;

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(expected);
    }

    [Test]
    public async Task A_LET_name_comes_before_a_defined_name_only_inside_the_LET()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        wb.DefinedNames.Add("x", "10");
        ws.Cell("C1").FormulaA1 = "SUM(LET(x,1,x),x)";
        ws.Cell("C2").FormulaA1 = "LET(x,1,x)+x";

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(11);
        await Assert.That(ws.Cell("C2").Value).IsEqualTo(11);
    }

    [Test]
    [Arguments("LET(x,y,y,1,x)")]
    [Arguments("LET(x,1,y)")]
    public async Task A_name_no_LET_has_bound_yet_is_not_recognised(string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C1").FormulaA1 = formula;

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.NameNotRecognized);
    }

    [Test]
    public async Task An_error_a_LET_binds_is_its_value_where_it_is_used()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C1").FormulaA1 = "LET(x,1/0,x)";

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task A_range_a_LET_binds_stays_a_reference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C1").FormulaA1 = "LET(r,A1,ISREF(r))";

        await Assert.That(ws.Cell("C1").Value.GetBoolean()).IsTrue();
    }

    [Test]
    public async Task A_LET_whose_result_is_an_array_spills()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 2;
        ws.Cell("A3").Value = 3;
        ws.Cell("C1").SetDynamicFormulaA1("LET(r,A1:A3,r*2)");

        await Assert.That(new[] { ws.Cell("C1").Value, ws.Cell("C2").Value, ws.Cell("C3").Value })
            .IsEquivalentTo(new XLCellValue[] { 2, 4, 6 });
    }

    [Test]
    [Arguments("LET(x,A1,x*2)", "A1")]
    [Arguments("LET(r,A1:A3,SUM(r))", "A2")]
    [Arguments("LET(x,1,y,x+A2,y)", "A2")]
    public async Task Changing_a_cell_a_LET_reads_recalculates_it(string formula, string input)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 2;
        ws.Cell("A3").Value = 3;
        ws.Cell("C1").FormulaA1 = formula;
        var before = ws.Cell("C1").Value;

        ws.Cell(input).Value = 100;

        await Assert.That(ws.Cell("C1").Value).IsNotEqualTo(before);
    }

    /// <summary>
    /// A LET value is intersected as any operand of a legacy formula is: Excel 365 gives 2 for
    /// <c>LET(x,A1:A3*2,SUM(x))</c> in C1 entered as a legacy formula (it shows
    /// <c>@A1:A3*2</c>), and 12 for the same formula entered as a dynamic one.
    /// </summary>
    [Test]
    public async Task A_LET_value_is_intersected_in_a_legacy_formula_and_not_in_a_dynamic_one()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 2;
        ws.Cell("A3").Value = 3;
        ws.Cell("C1").FormulaA1 = "LET(x,A1:A3*2,SUM(x))";
        ws.Cell("D1").SetDynamicFormulaA1("LET(x,A1:A3*2,SUM(x))");

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(2);
        await Assert.That(ws.Cell("D1").Value).IsEqualTo(12);
    }

    /// <summary>
    /// A name bound to a reference can be the operand of a range operator, whose result covers cells
    /// neither end names: Excel 365 recalculates <c>LET(r,A1,SUM(r:A3))</c> from 6 to 24 when A2
    /// becomes 20. The name <c>r</c> is also a column, and <c>r:A3</c> still reads as the name, the
    /// range operator and A3, as a column range would need a column after the colon (#870).
    /// </summary>
    [Test]
    [Arguments("start")]
    [Arguments("r")]
    public async Task A_range_built_from_a_LET_name_recalculates_when_a_cell_inside_it_changes(string name)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 2;
        ws.Cell("A3").Value = 3;
        ws.Cell("C1").FormulaA1 = $"LET({name},A1,SUM({name}:A3))";
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(6);

        ws.Cell("A2").Value = 20;

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(24);
    }

    /// <summary>
    /// A defined name's formula does not see the LET names of the formula that uses it, so its
    /// precedents are its own: here the defined name <c>Inner</c> reads the defined name <c>x</c>,
    /// which is A5, not the LET's <c>x</c>.
    /// </summary>
    [Test]
    public async Task A_defined_name_used_in_a_LET_keeps_its_own_precedents()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        wb.DefinedNames.Add("x", "Sheet1!$A$5");
        wb.DefinedNames.Add("Inner", "x*10");
        ws.Cell("A1").Value = 1;
        ws.Cell("A5").Value = 2;
        ws.Cell("C1").FormulaA1 = "LET(x,A1,Inner+x)";
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(21);

        ws.Cell("A5").Value = 3;

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(31);
    }

    /// <summary>
    /// Each name binds the value its expression has, once: a chain of thirty names each twice the
    /// one before reads the two-element array thirty times, not two to the thirtieth.
    /// </summary>
    [Test]
    public async Task A_chain_of_names_bound_to_arrays_evaluates_each_once()
    {
        var names = Enumerable.Range(0, 31).Select(i => $"n{i}_").ToArray();
        var bindings = string.Join(",", names.Skip(1).Select((name, i) => $"{name},{names[i]}+{names[i]}"));
        var formula = $"LET({names[0]},{{1;2}},{bindings},SUM({names[^1]}))";

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C1").SetDynamicFormulaA1(formula);
        var evaluation = Task.Run(() => ws.Cell("C1").Value);

        await Assert.That(await Task.WhenAny(evaluation, Task.Delay(10_000))).IsSameReferenceAs(evaluation);
        await Assert.That(await evaluation).IsEqualTo(3d * (1L << 30));
    }

    [Test]
    public async Task A_typed_name_that_reads_as_a_cell_is_not_bound_yet()
    {
        // Excel gives 1 for LET(A1,1,A1) and writes the name _xlpm.A1. Typed, the parser reads A1 as
        // a reference where a name goes, so XLibur does not bind it: a known gap (spec 08 Results).
        // Written as a file writes it, the name binds.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C1").FormulaA1 = "LET(A1,1,A1)";
        ws.Cell("C2").FormulaA1 = "_xlfn.LET(_xlpm.A1,5,_xlpm.A1)";

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(ws.Cell("C2").Value).IsEqualTo(5);
    }

    [Test]
    [Arguments("LET(x,1)")]
    [Arguments("LET(x,1,y,2)")]
    [Arguments("LET(1,2,3)")]
    [Arguments("LET(A1:A2,1,2)")]
    public async Task A_LET_Excel_would_refuse_is_a_value_error(string formula)
    {
        // Excel refuses these as they are typed; a file from another producer can still hold one.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C1").FormulaA1 = formula;

        await Assert.That(ws.Cell("C1").Value).IsEqualTo(XLError.IncompatibleValue);
    }
}
