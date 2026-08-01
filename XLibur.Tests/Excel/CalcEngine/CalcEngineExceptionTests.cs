using XLibur.Excel;
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
}
