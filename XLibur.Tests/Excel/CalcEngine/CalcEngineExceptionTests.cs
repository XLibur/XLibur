using XLibur.Excel;
using XLibur.Excel.CalcEngine.Exceptions;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

public class CalcEngineExceptionTests
{
    [Before(HookType.Class)]
    public static void SetCultureInfo()
    {
        Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
    }

    [Test]
    public async Task InvalidCharNumber()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("CHAR(-2)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr("CHAR(270)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task DivisionByZero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("0/0")).IsEqualTo(XLError.DivisionByZero);
        await Assert.That(new XLWorkbook().AddWorksheet().Evaluate("0/0")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task InvalidFunction()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("XXX(A1:A2)")).IsEqualTo(XLError.NameNotRecognized);

        var ws = new XLWorkbook().AddWorksheet();
        await Assert.That(ws.Evaluate("XXX(A1:A2)")).IsEqualTo(XLError.NameNotRecognized);
    }

    [Test]
    public async Task NestedNameNotRecognizedException()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").SetFormulaA1("=XXX");
        ws.Cell("A2").SetFormulaA1(@"=IFERROR(A1, ""Success"")");

        await Assert.That(ws.Cell("A2").Value).IsEqualTo("Success");
    }

    /// <summary>
    /// Every public entry point into the calc engine must report a missing formula location with
    /// an exception a caller outside the assembly can name.
    /// </summary>
    /// <remarks>
    /// Found by fuzzing (D37). <c>XLFunctionLibrary.TryInvoke</c> already translates the engine's
    /// internal <c>MissingContextException</c> into <see cref="XLNoWorksheetContextException"/>,
    /// and that type's own remarks call itself "the public face of the calc engine's internal
    /// missing-context signal". The three <c>Evaluate</c> overloads did not do the translation, so
    /// they threw the internal type instead — one <c>PublicSurfaceTests</c> asserts must never
    /// become visible. A caller could not catch it by name, and the interface XML documented it by
    /// <c>cref</c> anyway.
    /// </remarks>
    [Test]
    [Arguments("ROW()")]
    [Arguments("COLUMN()")]
    // The fuzzer's own input: A1:B341 lands in VLOOKUP's scalar lookup_value parameter, and
    // implicit intersection needs to know which row the formula is on.
    [Arguments("VLOOKUP(A1:B341,,1,FALSE)")]
    public async Task Worksheet_evaluate_without_a_formula_address_throws_a_public_exception(string expression)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");

        await Assert.That(() => ws.Evaluate(expression)).Throws<XLNoWorksheetContextException>();
    }

    [Test]
    public async Task Workbook_evaluate_without_a_worksheet_throws_a_public_exception()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Data");

        await Assert.That(() => wb.Evaluate("ROW()")).Throws<XLNoWorksheetContextException>();
    }

    [Test]
    public async Task EvaluateExpr_without_a_worksheet_throws_a_public_exception()
    {
        await Assert.That(() => XLWorkbook.EvaluateExpr("ROW()")).Throws<XLNoWorksheetContextException>();
    }

    /// <summary>
    /// Supplying the formula address is what the parameter is for, and it must keep working —
    /// the translation above must not swallow a call that has everything it needs.
    /// </summary>
    [Test]
    public async Task Worksheet_evaluate_with_a_formula_address_still_answers()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");

        await Assert.That(ws.Evaluate("ROW()", "B7")).IsEqualTo(7);
        await Assert.That(ws.Evaluate("COLUMN()", "B7")).IsEqualTo(2);
    }
}
