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
