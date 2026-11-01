using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// The functions that take a <c>LAMBDA</c>: <c>MAP</c>, <c>REDUCE</c>, <c>SCAN</c>, <c>BYROW</c>,
/// <c>BYCOL</c> and <c>MAKEARRAY</c>; <c>ISOMITTED</c> and optional parameters; and one of Excel's
/// functions passed by its name, <c>BYROW(D1:E2,SUM)</c> (#847, spec 08 phase 3).
/// </summary>
/// <remarks>
/// Every expected value is what Excel 365 calculated for the same formula, typed into a cell over COM
/// with <c>Formula2</c>. The workbook it saved is <see cref="Fixture"/>: A1:A3 hold 1, 2 and 3, B1:B3
/// hold 10, 20 and 30, and D1:E2 hold {1,2;3,4}. The formulas are in column G, every sixth row.
/// </remarks>
public class LambdaHelperTests
{
    private const string Fixture = @"Other\ExcelAuthored\LambdaHelperFormulas.xlsx";

    /// <summary>
    /// The cells of the fixture whose formulas Excel calculated, and the cells their results spilled
    /// into. Left out: G307, <c>ISOMITTED()</c>, which Excel refused to enter.
    /// </summary>
    public static IEnumerable<string> CalculatedCells()
    {
        foreach (var row in Enumerable.Range(0, 66).Select(i => 1 + 6 * i).Except([307]))
        {
            for (var r = row; r < row + 5; r++)
            {
                foreach (var column in new[] { "G", "H", "I" })
                    yield return $"{column}{r}";
            }
        }
    }

    [Test]
    [MethodDataSource(nameof(CalculatedCells))]
    public async Task A_workbook_Excel_wrote_recalculates_to_the_values_Excel_saved(string address)
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var cell = wb.Worksheet(1).Cell(address);
        var cached = cell.CachedValue;

        wb.RecalculateAllFormulas();

        await Assert.That(cell.Value).IsEqualTo(cached).Because(Formula(wb, address));
    }

    [Test]
    [Arguments("MAP(A1:A3,LAMBDA(v,v*2))", 2, 4, 6)]
    [Arguments("SCAN(0,A1:A3,LAMBDA(a,v,a+v))", 1, 3, 6)]
    [Arguments("BYROW(D1:E2,SUM)", 3, 7, null)]
    [Arguments("MAP(A1:A3,ABS)", 1, 2, 3)]
    public async Task A_typed_helper_spills_what_Excel_calculates(string formula, double first, double second, double? third)
    {
        using var wb = CreateWorkbook();
        var ws = wb.Worksheet(1);
        ws.Cell("G1").SetDynamicFormulaA1(formula);

        await Assert.That(ws.Cell("G1").Value).IsEqualTo(first);
        await Assert.That(ws.Cell("G2").Value).IsEqualTo(second);
        await Assert.That(ws.Cell("G3").Value).IsEqualTo(third is { } t ? t : Blank.Value);
    }

    [Test]
    [Arguments("REDUCE(0,A1:A3,LAMBDA(acc,v,acc+v))", 6)]
    [Arguments("LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(1)", 1)]
    [Arguments("LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(1,2)", 3)]
    [Arguments("LET(f,LAMBDA(a,[b],IF(ISOMITTED(b),a*2,a*b)),f(3)+f(3,4))", 18)]
    [Arguments("LET(f,SUM,f(1,2))", 3)]
    public async Task A_typed_formula_evaluates_as_Excel_does(string formula, double expected)
    {
        using var wb = CreateWorkbook();
        wb.Worksheet(1).Cell("G1").FormulaA1 = formula;

        await Assert.That(wb.Worksheet(1).Cell("G1").Value).IsEqualTo(expected);
    }

    /// <summary>
    /// Excel writes one of its functions passed by its name with the prefix <c>_xleta.</c>, and reads a
    /// name without it as a defined name: <c>_xlfn.BYROW(D1:E2,_xleta.SUM)</c>. It writes the function's
    /// name in its own spelling, so a typed <c>Rate</c> that names nothing is <c>_xleta.RATE</c>.
    /// </summary>
    [Test]
    [Arguments("BYROW(D1:E2,SUM)", "_xlfn.BYROW(D1:E2,_xleta.SUM)")]
    [Arguments("LET(f,SUM,f(1,2))", "_xlfn.LET(_xlpm.f,_xleta.SUM,_xlpm.f(1,2))")]
    [Arguments("LET(x,1,x+Rate)", "_xlfn.LET(_xlpm.x,1,_xlpm.x+_xleta.RATE)")]
    [Arguments("map(A1:A3,abs)", "_xlfn.MAP(A1:A3,_xleta.ABS)")]
    [Arguments("SUM", "_xleta.SUM")]
    [Arguments("ISERROR(DAYS360)", "ISERROR(_xleta.DAYS360)")]
    [Arguments("ISERROR(LOG10)", "ISERROR(LOG10)")]
    [Arguments("ISERROR(SUM)", "ISERROR(_xleta.SUM)")]
    [Arguments("SUM(A1)+ABS(-1)", "SUM(A1)+ABS(-1)")]
    [Arguments("IF(TRUE,\"SUM\",FALSE)", "IF(TRUE,\"SUM\",FALSE)")]
    [Arguments("MAP(A1:A3,LAMBDA(v,SUM(v)))", "_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,SUM(_xlpm.v)))")]
    public async Task A_function_passed_by_its_name_is_written_as_Excel_writes_it(string formula, string written)
    {
        using var wb = CreateWorkbook();
        wb.Worksheet(1).Cell("G1").FormulaA1 = formula;

        await Assert.That(wb.Worksheet(1).Cell("G1").FormulaA1).IsEqualTo(written);
    }

    [Test]
    [Arguments("MAP(A1:A3,LAMBDA(v,v+$B$1))", "B1")]
    [Arguments("REDUCE(0,A1:A3,LAMBDA(a,v,a+v))", "A2")]
    [Arguments("BYROW(D1:E2,SUM)", "D1")]
    public async Task Changing_a_cell_a_helper_reads_recalculates_it(string formula, string input)
    {
        using var wb = CreateWorkbook();
        var ws = wb.Worksheet(1);
        ws.Cell("G1").FormulaA1 = formula;
        var before = ws.Cell("G1").Value;

        ws.Cell(input).Value = 1000;

        await Assert.That(ws.Cell("G1").Value).IsNotEqualTo(before);
    }

    /// <summary>One of Excel's functions passed by its name keeps what it does with an array: ABS spills.</summary>
    [Test]
    public async Task A_function_passed_by_its_name_spills_an_array_in_a_dynamic_formula()
    {
        using var wb = CreateWorkbook();
        var ws = wb.Worksheet(1);
        ws.Cell("G1").SetDynamicFormulaA1("LET(f,ABS,f({-1,-2}))");

        await Assert.That(ws.Cell("G1").Value).IsEqualTo(1);
        await Assert.That(ws.Cell("H1").Value).IsEqualTo(2);
    }

    /// <summary>
    /// An accumulator that is an array is computed at each step, not carried as the operations that
    /// make it: 10,000 lazy additions, one on another, overflowed the stack when the result was read.
    /// </summary>
    [Test]
    public async Task REDUCE_carries_an_array_accumulator_through_many_steps()
    {
        using var wb = CreateWorkbook();
        var ws = wb.Worksheet(1);
        ws.Cell("G1").SetDynamicFormulaA1("REDUCE({0,0},SEQUENCE(10000),LAMBDA(a,v,a+v))");

        await Assert.That(ws.Cell("G1").Value).IsEqualTo(50005000);
        await Assert.That(ws.Cell("H1").Value).IsEqualTo(50005000);
    }

    /// <summary>A defined name whose value is #NAME? is that name, not the function of its spelling.</summary>
    [Test]
    public async Task A_defined_name_with_a_function_spelling_is_the_name_whatever_its_value()
    {
        using var wb = CreateWorkbook();
        wb.DefinedNames.Add("Rate", "Missing");
        wb.Worksheet(1).Cell("G1").FormulaA1 = "MAP(A1:A3,Rate)";

        await Assert.That(wb.Worksheet(1).Cell("G1").Value).IsEqualTo(XLError.NameNotRecognized);
    }

    /// <summary>
    /// A defined name that calls itself by a function's spelling calls itself: its formula is written
    /// with the name, not with the function passed by its name.
    /// </summary>
    [Test]
    public async Task A_defined_name_that_calls_itself_by_a_function_spelling_calls_itself()
    {
        using var wb = CreateWorkbook();
        wb.DefinedNames.Add("Rate", "LAMBDA(n,IF(n=0,1,SUM(MAP(n-1,Rate))*2))");
        // Called as a value: a call by the name, Rate(3), calls Excel's RATE.
        wb.Worksheet(1).Cell("G1").FormulaA1 = "MAP(3,Rate)";

        await Assert.That(wb.DefinedNames.DefinedName("Rate").RefersTo).DoesNotContain("_xleta.");
        await Assert.That(wb.Worksheet(1).Cell("G1").Value).IsEqualTo(8);
    }

    /// <summary>The sheet the formulas were calculated on in Excel.</summary>
    private static XLWorkbook CreateWorkbook()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        for (var row = 1; row <= 3; row++)
        {
            ws.Cell(row, 1).Value = row;
            ws.Cell(row, 2).Value = row * 10;
        }

        ws.Cell("D1").Value = 1;
        ws.Cell("E1").Value = 2;
        ws.Cell("D2").Value = 3;
        ws.Cell("E2").Value = 4;
        return wb;
    }

    /// <summary>The formula of the cell, or of the cell its result spilled from, for a failure message.</summary>
    private static string Formula(XLWorkbook wb, string address)
    {
        var ws = wb.Worksheet(1);
        var cell = ws.Cell(address);
        for (var row = cell.Address.RowNumber; row >= 1; row--)
        {
            var anchor = ws.Cell(row, "G");
            if (anchor.HasFormula)
                return $"{address} of {anchor.FormulaA1}";
        }

        return address;
    }
}
