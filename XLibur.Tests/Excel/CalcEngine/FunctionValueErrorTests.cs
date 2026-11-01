using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// A function value, a <c>LAMBDA</c> or one of Excel's functions passed by its name, is no error:
/// <c>ISERROR</c> and <c>ISERR</c> are FALSE, <c>ERROR.TYPE</c> is <c>#N/A</c>, <c>TYPE</c> is 128,
/// and <c>IFERROR</c> and <c>IFNA</c> give the function, which a call can then call (#868).
/// </summary>
/// <remarks>
/// Every expected value is what Excel 365 calculated for the same formula, typed into a cell over COM
/// with <c>Formula2</c>. The workbook it saved is <see cref="Fixture"/>: A1:A3 hold 1, 2 and 3, and
/// the formulas are in column G, every sixth row.
/// </remarks>
public class FunctionValueErrorTests
{
    private const string Fixture = @"Other\ExcelAuthored\FunctionValueErrorTests.xlsx";

    public static IEnumerable<string> CalculatedCells() =>
        Enumerable.Range(0, 15).Select(i => $"G{1 + 6 * i}");

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
    [Arguments("ISERROR(LAMBDA(x,x))", false)]
    [Arguments("ISERR(SUM)", false)]
    [Arguments("ISERROR(1/0)", true)]
    public async Task An_error_test_of_a_typed_function_value_is_as_Excel_gives_it(string formula, bool expected)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("ERROR.TYPE(LAMBDA(x,x))", XLError.NoValueAvailable)]
    [Arguments("IFERROR(LAMBDA(x,x),1)", XLError.Calc)]
    [Arguments("IFNA(LAMBDA(x,x),1)", XLError.Calc)]
    [Arguments("IFERROR(1/0,LAMBDA(x,x))", XLError.Calc)]
    public async Task A_typed_function_value_is_no_error_to_the_functions_that_test_for_one(string formula, XLError expected)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("TYPE(LAMBDA(x,x))", 128)]
    [Arguments("LET(f,IFERROR(LAMBDA(x,x*2),1),f(4))", 8)]
    [Arguments("LET(f,IFNA(LAMBDA(x,x*3),1),f(4))", 12)]
    [Arguments("IFNA(NA(),5)", 5)]
    [Arguments("IFNA(7,5)", 7)]
    public async Task A_typed_formula_with_a_function_value_calculates_as_Excel_does(string formula, double expected)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(expected);
    }

    /// <summary>
    /// A function that reads a function value as a value still does so for each element of an array
    /// formula: only the functions that test for an error see the function whole.
    /// </summary>
    [Test]
    public async Task Any_other_function_passed_a_function_value_in_an_array_formula_gives_VALUE_for_each_element()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("G1").SetDynamicFormulaA1("ROUND({1,2},LAMBDA(x,x))");

        await Assert.That(ws.Cell("G1").Value).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(ws.Cell("H1").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task IFNA_keeps_an_error_other_than_NA()
    {
        await Assert.That(Evaluate("IFNA(1/0,5)")).IsEqualTo(XLError.DivisionByZero);
    }

    private static XLCellValue Evaluate(string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("G1").FormulaA1 = formula;
        return ws.Cell("G1").Value;
    }
}
