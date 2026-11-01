using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// <c>LAMBDA</c>: a function value, called where it is written, through a <c>LET</c> name or through a
/// defined name (#847, spec 08 phase 2).
/// </summary>
/// <remarks>
/// Every expected value is what Excel 365 calculated for the same formula, typed into a cell over
/// COM with <c>Formula2</c>. A1:A3 hold 1, 2 and 3.
/// </remarks>
public class LambdaTests
{
    private const string Fixture = @"Other\ExcelAuthored\LambdaFormulas.xlsx";

    /// <summary>
    /// The cells of the fixture whose formulas Excel calculated, every fourth row of column D, so a
    /// result can spill. Left out: D137, D141, D149 and D193, which Excel refused to enter; the
    /// cells of <see cref="A_function_Excel_saved_as_the_result_of_a_cell_loads_and_recalculates_as_CALC"/>; the cells of <see cref="An_error_test_reads_a_function_as_VALUE_where_Excel_reads_no_error"/>;
    /// and D165, <c>LAMBDA(x,ISREF(x))(A1)</c>, which is FALSE because a dynamic formula reads
    /// <c>ISREF(A1)</c> as FALSE, LAMBDA or not.
    /// </summary>
    public static IEnumerable<string> CalculatedCells() =>
        Enumerable.Range(0, 51).Select(i => 1 + 4 * i)
            .Except([137, 141, 149, 193, 1, 81, 89, 17, 21, 121, 181, 189, 165])
            .Select(row => $"D{row}");

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

    /// <summary>
    /// The cells of the fixture whose result is a function. Excel writes <c>#CALC!</c> as it writes
    /// every error newer than Excel 2007: <c>#VALUE!</c>, with value metadata whose rich value names
    /// the error. The load reads the error the metadata names.
    /// </summary>
    [Test]
    [Arguments("D1")]
    [Arguments("D81")]
    [Arguments("D89")]
    public async Task A_function_Excel_saved_as_the_result_of_a_cell_loads_and_recalculates_as_CALC(string address)
    {
        using var wb = TestHelper.LoadWorkbook(Fixture);
        var cell = wb.Worksheet(1).Cell(address);
        await Assert.That(cell.CachedValue).IsEqualTo(XLError.Calc);

        wb.RecalculateAllFormulas();

        await Assert.That(cell.Value).IsEqualTo(XLError.Calc);
    }

    /// <summary>
    /// A known gap: XLibur reads a function as <c>#VALUE!</c> wherever a value is wanted, and the
    /// functions that test for an error see that error. Excel sees no error: <c>ISERROR</c> and
    /// <c>ISERR</c> are FALSE, <c>ERROR.TYPE</c> is <c>#N/A</c>, and <c>IFERROR</c> gives the function,
    /// which is <c>#CALC!</c>. The fixture holds each, in D17, D181, D21 and D189.
    /// </summary>
    [Test]
    [Arguments("ISERROR(LAMBDA(x,x))", true)]
    [Arguments("ISERR(LAMBDA(x,x))", true)]
    [Arguments("ERROR.TYPE(LAMBDA(x,x))", 3)]
    [Arguments("IFERROR(LAMBDA(x,x),1)", 1)]
    public async Task An_error_test_reads_a_function_as_VALUE_where_Excel_reads_no_error(string formula, object xlibur)
    {
        var expected = xlibur is bool logical ? logical : (XLCellValue)(int)xlibur;

        await Assert.That(Evaluate(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("LAMBDA(x,x*2)(3)", 6)]
    [Arguments("_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(3)", 6)]
    [Arguments("LAMBDA(X,x+1)(1)", 2)]
    [Arguments("LAMBDA(x,5)(1/0)", 5)]
    [Arguments("LAMBDA(x,y,x+y)(1,)", 1)]
    [Arguments("LAMBDA(x,x+A1)(1)", 2)]
    [Arguments("ROWS(LAMBDA(x,x)(A1:A3))", 3)]
    [Arguments("LAMBDA(x,x*2)(3)+LAMBDA(x,x*3)(1)", 9)]
    [Arguments("IF(TRUE,LAMBDA(x,x*2))(3)", 6)]
    [Arguments("LAMBDA(x,LAMBDA(y,x+y))(1)(2)", 3)]
    [Arguments("LAMBDA(f,f(3))(LAMBDA(y,y*2))", 6)]
    public async Task A_LAMBDA_called_where_it_is_written_evaluates_its_body(string formula, double expected)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("IF(TRUE,LAMBDA(x,x*2))(3)", 6)]
    [Arguments("IF(FALSE,1,LAMBDA(x,x*3))(3)", 9)]
    public async Task A_LAMBDA_that_IF_picks_in_a_dynamic_formula_is_called(string formula, double expected)
    {
        using var wb = CreateWorkbook();
        var ws = wb.Worksheet(1);
        ws.Cell("D1").SetDynamicFormulaA1(formula);

        await Assert.That(ws.Cell("D1").Value).IsEqualTo(expected);
    }

    [Test]
    [Arguments("LET(f,LAMBDA(x,x*2),f(3))", 6)]
    [Arguments("_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2),_xlpm.f(3))", 6)]
    [Arguments("LET(x,LAMBDA(y,y),5)", 5)]
    [Arguments("LET(y,10,f,LAMBDA(x,x+y),f(1))", 11)]
    [Arguments("LET(Double,LAMBDA(x,x*10),Double(1))", 10)]
    public async Task A_LAMBDA_bound_to_a_LET_name_is_called_by_the_name(string formula, double expected)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(expected);
    }

    [Test]
    public async Task A_LAMBDA_sees_the_names_bound_where_it_is_written_not_where_it_is_called()
    {
        await Assert.That(Evaluate("LET(y,10,f,LAMBDA(x,x+y),LET(y,100,f(1)))")).IsEqualTo(11);
    }

    [Test]
    [Arguments("Double(4)", 8)]
    [Arguments("Double(Double(1))", 4)]
    [Arguments("Factorial(5)", 120)]
    [Arguments("AddB(1)", 101)]
    [Arguments("Apply(Double,5)", 10)]
    public async Task A_defined_name_that_refers_to_a_LAMBDA_is_a_function(string formula, double expected)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("LET(f,LAMBDA(x,x*2),f+1)")]
    [Arguments("SUM(LAMBDA(x,x))")]
    [Arguments("LAMBDA(x,x)=1")]
    [Arguments("LAMBDA(x,x)&\"a\"")]
    public async Task A_function_read_as_a_value_is_VALUE(string formula)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("LAMBDA(x,y,x+y)(1)")]
    [Arguments("LAMBDA(x,x)(1,2)")]
    [Arguments("LAMBDA(x,x)()")]
    [Arguments("LAMBDA(x)(1)")]
    public async Task A_call_with_more_or_fewer_arguments_than_parameters_is_VALUE(string formula)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task An_error_passed_to_a_parameter_the_body_uses_is_the_result()
    {
        await Assert.That(Evaluate("Double(1/0)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task A_range_passed_to_a_parameter_stays_a_reference()
    {
        await Assert.That(Evaluate("LAMBDA(x,ISREF(x))(A1)").GetBoolean()).IsTrue();
        await Assert.That(Evaluate("LAMBDA(x,x)(\"a\")")).IsEqualTo("a");
    }

    [Test]
    [Arguments("LAMBDA(x,x)(A1:A3)", 1, 2, 3)]
    [Arguments("Double(A1:A3)", 2, 4, 6)]
    public async Task A_call_whose_result_is_an_array_spills(string formula, double first, double second, double third)
    {
        using var wb = CreateWorkbook();
        var ws = wb.Worksheet(1);
        ws.Cell("D1").SetDynamicFormulaA1(formula);

        await Assert.That(new[] { ws.Cell("D1").Value, ws.Cell("D2").Value, ws.Cell("D3").Value })
            .IsEquivalentTo(new XLCellValue[] { first, second, third });
    }

    [Test]
    [Arguments("LAMBDA(x,x+A2)(1)", "A2")]
    [Arguments("LET(f,LAMBDA(x,x+A2),f(1))", "A2")]
    [Arguments("Double(A2)", "A2")]
    [Arguments("AddB(1)", "B1")]
    [Arguments("Apply(AddB,1)", "B1")]
    public async Task Changing_a_cell_a_call_reads_recalculates_it(string formula, string input)
    {
        using var wb = CreateWorkbook();
        var ws = wb.Worksheet(1);
        ws.Cell("D1").FormulaA1 = formula;
        var before = ws.Cell("D1").Value;

        ws.Cell(input).Value = 1000;

        await Assert.That(ws.Cell("D1").Value).IsNotEqualTo(before);
    }

    [Test]
    [Arguments("LAMBDA(x,x*2)")]
    [Arguments("LET(f,LAMBDA(x,x*2),f)")]
    [Arguments("LAMBDA(x,x)(LAMBDA(y,y))")]
    [Arguments("Double")]
    public async Task A_function_left_as_the_result_of_a_cell_is_CALC(string formula)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(XLError.Calc);
    }

    [Test]
    [Arguments("Loop(1)")]
    [Arguments("IFERROR(Loop(1),\"stop\")")]
    public async Task A_function_that_calls_itself_for_ever_is_NUM_which_IFERROR_does_not_catch(string formula)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    /// <summary>
    /// <c>SumTo(n)</c> makes n + 1 calls, one inside another. Excel's limit depends on the stack its
    /// formula takes: it calculated <c>SumTo(2000)</c> and not <c>SumTo(2500)</c>. XLibur's is a fixed
    /// 1,024 calls, on any thread: a call that runs out of stack goes on on a thread with more.
    /// </summary>
    [Test]
    public async Task A_function_can_call_itself_1024_calls_deep()
    {
        await Assert.That(Evaluate("SumTo(1023)")).IsEqualTo(523776);
        await Assert.That(Evaluate("SumTo(1024)")).IsEqualTo(XLError.NumberInvalid);
    }

    /// <summary>
    /// Excel writes a defined name's <c>LAMBDA</c> with the prefixes of a cell formula:
    /// <c>_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)</c>.
    /// </summary>
    [Test]
    public async Task A_LAMBDA_typed_into_a_defined_name_is_saved_as_Excel_writes_it()
    {
        using var wb = CreateWorkbook();
        wb.Worksheet(1).Cell("D1").FormulaA1 = "Double(4)";
        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        using var reloaded = new XLWorkbook(saved);
        reloaded.RecalculateAllFormulas();

        await Assert.That(reloaded.DefinedNames.DefinedName("Double").RefersTo).IsEqualTo("_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)");
        await Assert.That(reloaded.Worksheet(1).Cell("D1").Value).IsEqualTo(8);
    }

    private static XLCellValue Evaluate(string formula)
    {
        using var wb = CreateWorkbook();
        var ws = wb.Worksheet(1);
        ws.Cell("D1").FormulaA1 = formula;
        return ws.Cell("D1").Value;
    }

    /// <summary>
    /// The sheet and the defined names the formulas were calculated with in Excel. B1 holds 100.
    /// </summary>
    private static XLWorkbook CreateWorkbook()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 2;
        ws.Cell("A3").Value = 3;
        ws.Cell("B1").Value = 100;
        wb.DefinedNames.Add("Double", "LAMBDA(x,x*2)");
        wb.DefinedNames.Add("Factorial", "LAMBDA(n,IF(n<=1,1,n*Factorial(n-1)))");
        wb.DefinedNames.Add("Loop", "LAMBDA(n,Loop(n+1))");
        wb.DefinedNames.Add("SumTo", "LAMBDA(n,IF(n=0,0,n+SumTo(n-1)))");
        wb.DefinedNames.Add("AddB", "LAMBDA(x,x+Sheet1!$B$1)");
        wb.DefinedNames.Add("Apply", "LAMBDA(f,v,f(v))");
        return wb;
    }
}
