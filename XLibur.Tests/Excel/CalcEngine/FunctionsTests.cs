using XLibur.Excel;
using System;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

public class FunctionsTests
{
    [Before(HookType.Test)]
    public void Init()
    {
        // Make sure tests run on a deterministic culture
        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
    }

    [Test]
    public async Task Asc()
    {
        Object actual;

        actual = XLWorkbook.EvaluateExpr(@"Asc(""Text"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("Text"));
    }

    [Test]
    public async Task Clean()
    {
        object actual = XLWorkbook.EvaluateExpr($@"Clean(""A{Environment.NewLine}B"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("AB"));
    }

    [Test]
    public async Task Dollar()
    {
        using var wb = new XLWorkbook();
        object actual = wb.Evaluate("DOLLAR(12345.123)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(TestHelper.CurrencySymbol + "12,345.12"));

        actual = wb.Evaluate("DOLLAR(12345.123, 1)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(TestHelper.CurrencySymbol + "12,345.1"));
    }

    [Test]
    [Arguments("A", "A", true)]
    [Arguments("A", "a", false)]
    [Arguments("", "", true)]
    public async Task Exact(string lhs, string rhs, bool result)
    {
        var actual = XLWorkbook.EvaluateExpr($"EXACT(\"{lhs}\", \"{rhs}\")");
        await Assert.That(actual).IsEqualTo(result);
    }

    [Test]
    public async Task Exact_converts_values_to_text()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("EXACT(TRUE, \"true\")")).IsEqualTo(false);
        await Assert.That(XLWorkbook.EvaluateExpr("EXACT(TRUE, \"TRUE\")")).IsEqualTo(true);
        await Assert.That(XLWorkbook.EvaluateExpr("EXACT(1, \"1\")")).IsEqualTo(true);
        await Assert.That(XLWorkbook.EvaluateExpr("EXACT(IF(TRUE,), \"\")")).IsEqualTo(true);

        // Check blank cell
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.Evaluate("EXACT(A1, \"\")")).IsEqualTo(true);
    }

    [Test]
    public async Task Exact_propagates_errors()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("EXACT(#DIV/0!, \"A\")")).IsEqualTo(XLError.DivisionByZero);
        await Assert.That(XLWorkbook.EvaluateExpr("EXACT(\"A\", #DIV/0!)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task Fixed()
    {
        Object actual;

        actual = XLWorkbook.EvaluateExpr("Fixed(12345.123)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("12,345.12"));

        actual = XLWorkbook.EvaluateExpr("Fixed(12345.123, 1)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("12,345.1"));

        actual = XLWorkbook.EvaluateExpr("Fixed(12345.123, 1, TRUE)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("12345.1"));
    }

    [Test]
    public async Task Formula_from_another_sheet()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("ws1");
        ws1.FirstCell().SetValue(1).CellRight().SetFormulaA1("A1 + 1");
        var ws2 = wb.AddWorksheet("ws2");
        ws2.FirstCell().SetFormulaA1("ws1!B1 + 1");
        object v = ws2.FirstCell().Value;
        await Assert.That(v).IsEqualTo(ExpectedCellValue.From(3.0));
    }

    [Test]
    public async Task TextConcat()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 1;
        ws.Cell("B1").Value = 1;
        ws.Cell("B2").Value = 1;

        ws.Cell("C1").FormulaA1 = "\"The total value is: \" & SUM(A1:B2)";

        object r = ws.Cell("C1").Value;
        await Assert.That(r).IsEqualTo(ExpectedCellValue.From("The total value is: 4"));
    }

    [Test]
    public async Task Trim()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("Trim(\"Test    \")")).IsEqualTo("Test");

        //Should not trim non breaking space
        //See http://office.microsoft.com/en-us/excel-help/trim-function-HP010062581.aspx
        await Assert.That(XLWorkbook.EvaluateExpr("Trim(\"Test\u00A0 \")")).IsEqualTo("Test\u00A0");
    }

    [Test]
    public async Task TestEmptyTallyOperations()
    {
        //In these test no values have been set
        var wb = new XLWorkbook();
        wb.Worksheets.Add("TallyTests");
        var cell = wb.Worksheet(1).Cell(1, 1).SetFormulaA1("=MAX(D1,D2)");
        await Assert.That(cell.Value).IsEqualTo(0);
        cell = wb.Worksheet(1).Cell(2, 1).SetFormulaA1("=MIN(D1,D2)");
        await Assert.That(cell.Value).IsEqualTo(0);
        cell = wb.Worksheet(1).Cell(3, 1).SetFormulaA1("=SUM(D1,D2)");
        await Assert.That(cell.Value).IsEqualTo(0);
    }

    [Test]
    public async Task TestOmittedParameters()
    {
        using var wb = new XLWorkbook();
        object value = wb.Evaluate("=IF(TRUE,1)");
        await Assert.That(value).IsEqualTo(ExpectedCellValue.From(1));

        value = wb.Evaluate("=IF(TRUE,1,)");
        await Assert.That(value).IsEqualTo(ExpectedCellValue.From(1));

        value = wb.Evaluate("=ISBLANK(IF(FALSE,1,))");
        await Assert.That(value).IsEqualTo(ExpectedCellValue.From(true));

        value = wb.Evaluate("=IF(FALSE,,2)");
        await Assert.That(value).IsEqualTo(ExpectedCellValue.From(2));
    }

    [Test]
    public async Task TestDefaultExcelFunctionNamespace()
    {
        await Assert.That(() => XLWorkbook.EvaluateExpr("TODAY()")).ThrowsNothing();
        await Assert.That(() => XLWorkbook.EvaluateExpr("_xlfn.TODAY()")).ThrowsNothing();
        await Assert.That((bool)XLWorkbook.EvaluateExpr("_xlfn.TODAY() = TODAY()")).IsTrue();
    }

    [Test]
    [Arguments("=1234%", 12.34)]
    [Arguments("=1234%%", 0.1234)]
    [Arguments("=100+200%", 102.0)]
    [Arguments("=100%+200", 201.0)]
    [Arguments("=(100+200)%", 3.0)]
    [Arguments("=200%^5", 32.0)]
    [Arguments("=200%^400%", 16.0)]
    [Arguments("=SUM(100,200,300)%", 6.0)]
    public async Task PercentOperator(string formula, double expectedResult)
    {
        var res = (double)XLWorkbook.EvaluateExpr(formula);

        await Assert.That(res).IsEqualTo(expectedResult).Within(XLHelper.Epsilon);
    }

    [Test]
    [Arguments("=--1", 1)]
    [Arguments("=++1", 1)]
    [Arguments("=-+-+-1", -1)]
    [Arguments("=2^---2", 0.25)]
    public async Task MultipleUnaryOperators(string formula, double expectedResult)
    {
        var res = (double)XLWorkbook.EvaluateExpr(formula);

        await Assert.That(res).IsEqualTo(expectedResult).Within(XLHelper.Epsilon);
    }

    [Test]
    [Arguments("RIGHT(\"2020\", 2) + 1", 21)]
    [Arguments("LEFT(\"20.2020\", 6) + 1", 21.202)]
    [Arguments("2 + (\"3\" & \"4\")", 36)]
    [Arguments("2 + \"3\" & \"4\"", "54")]
    [Arguments("\"7\" & \"4\"", "74")]
    public async Task TestStringSubExpression(string formula, object expectedResult)
    {
        var actual = XLWorkbook.EvaluateExpr(formula);

        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(expectedResult));
    }

    [Test]
    public async Task Cell_function_is_evaluated_to_reference_error()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = "$B$4(5)";

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(XLError.CellReference);
    }
}
