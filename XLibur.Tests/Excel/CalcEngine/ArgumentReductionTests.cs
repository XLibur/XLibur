using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Functions registered to take ranges receive their scalar arguments unreduced and read them
/// through <c>AnyValue.TryReduceToNumber/Int/Logical</c> and <c>args.IsOmitted</c>. These tests pin
/// the conversions and the blank handling each family relies on: the text functions read a blank
/// optional int as its default, the dynamic-array functions read it as 0.
/// </summary>
[SetCulture("en-US")]
public class ArgumentReductionTests
{
    [Test]
    public async Task TextFunctions_ReadABlankOptionalIntAsItsDefault()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;

        // A blank instance_num is the default 1, not 0 (which would be an error).
        await Assert.That(ws.Evaluate("TEXTBEFORE(\"a-b-c\", \"-\", )")).IsEqualTo("a");
        await Assert.That(ws.Evaluate("TEXTBEFORE(\"a-b-c\", \"-\", A1)")).IsEqualTo("a-b");
        await Assert.That(ws.Evaluate("TEXTBEFORE(\"a-b-c\", \"-\", 2.9)")).IsEqualTo("a-b");
        await Assert.That(ws.Evaluate("TEXTBEFORE(\"a-b-c\", \"-\", -1.5)")).IsEqualTo("a-b");
        await Assert.That(ws.Evaluate("TEXTBEFORE(\"a-b-c\", \"-\", \"x\")")).IsEqualTo(XLError.IncompatibleValue);

        // A blank match_mode is the default 0, a case-sensitive match.
        await Assert.That(ws.Evaluate("TEXTBEFORE(\"a-B-c\", \"b\", 1, )")).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(ws.Evaluate("TEXTBEFORE(\"a-B-c\", \"b\", 1, 1)")).IsEqualTo("a-");
        await Assert.That(ws.Evaluate("TEXTBEFORE(\"a-B-c\", \"b\", 1, 2)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task TextSplit_TreatsBlankDelimitersAndPaddingAsOmitted()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        await Assert.That(ws.Evaluate("TEXTSPLIT(\"a,b\", , )")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(ws.Evaluate("COLUMNS(TEXTSPLIT(\"a,b\", \",\", ))")).IsEqualTo(2);
        await Assert.That(ws.Evaluate("INDEX(TEXTSPLIT(\"a,b;c\", \",\", \";\", , , \"-\"), 2, 2)")).IsEqualTo("-");
        await Assert.That(ws.Evaluate("INDEX(TEXTSPLIT(\"a,b;c\", \",\", \";\", , , ), 2, 2)")).IsEqualTo(XLError.NoValueAvailable);
    }

    [Test]
    public async Task DynamicArrayFunctions_ReduceIntAndLogicalArguments()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;

        await Assert.That(ws.Evaluate("ROWS(SEQUENCE(A1))")).IsEqualTo(2);
        await Assert.That(ws.Evaluate("SUM(SEQUENCE(2, 2.9))")).IsEqualTo(10);
        await Assert.That(ws.Evaluate("SUM(SEQUENCE(\"2\", TRUE))")).IsEqualTo(3);
        await Assert.That(ws.Evaluate("SUM(SEQUENCE(2, 1, A1, 0.5))")).IsEqualTo(4.5);
        await Assert.That(ws.Evaluate("SEQUENCE(2, \"x\")")).IsEqualTo(XLError.IncompatibleValue);

        await Assert.That(ws.Evaluate("ROWS(UNIQUE({1;1;2}, \"FALSE\"))")).IsEqualTo(2);
        await Assert.That(ws.Evaluate("ROWS(UNIQUE({1;1;2}, , 1))")).IsEqualTo(1);
        await Assert.That(ws.Evaluate("UNIQUE({1;1;2}, \"maybe\")")).IsEqualTo(XLError.IncompatibleValue);

        // An empty placeholder leaves TAKE's row count alone.
        await Assert.That(ws.Evaluate("COLUMNS(TAKE({1,2,3;4,5,6}, , 2))")).IsEqualTo(2);
        await Assert.That(ws.Evaluate("ROWS(TAKE({1,2,3;4,5,6}, , 2))")).IsEqualTo(2);
    }

    [Test]
    public async Task Aggregate_ReducesItsScalarArguments()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 9;
        ws.Cell("A2").Value = 1;
        ws.Cell("B1").Value = 1;
        ws.Cell("B2").Value = 2;
        ws.Cell("B3").Value = 3;

        await Assert.That(ws.Evaluate("AGGREGATE(A1, 0, B1:B3)")).IsEqualTo(6);
        await Assert.That(ws.Evaluate("AGGREGATE(\"4\", 0, B1:B3)")).IsEqualTo(3);
        await Assert.That(ws.Evaluate("AGGREGATE(14, 0, B1:B3, A2)")).IsEqualTo(3);
        await Assert.That(ws.Evaluate("AGGREGATE(\"x\", 0, B1:B3)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Linest_ReadsBlankFlagsAsTheirDefaults()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // const blank = TRUE (fit an intercept), stats blank = FALSE (one row of coefficients).
        await Assert.That(ws.Evaluate("ROWS(LINEST({1,2,4}, {1,2,3}, , ))")).IsEqualTo(1);
        await Assert.That(ws.Evaluate("INDEX(LINEST({2,4,6}, {1,2,3}, , ), 1, 1)")).IsEqualTo(2);
        await Assert.That(ws.Evaluate("ROWS(LINEST({1,2,4}, {1,2,3}, , \"TRUE\"))")).IsEqualTo(5);
        await Assert.That(ws.Evaluate("LINEST({1,2,4}, {1,2,3}, , \"maybe\")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Irr_ReducesTheGuessArgument()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 0.05;

        await Assert.That((double)ws.Evaluate("IRR({-100, 110}, A1)")).IsEqualTo(0.1).Within(1e-9);
        await Assert.That((double)ws.Evaluate("IRR({-100, 110}, \"0.2\")")).IsEqualTo(0.1).Within(1e-9);
        await Assert.That(ws.Evaluate("IRR({-100, 110}, \"x\")")).IsEqualTo(XLError.IncompatibleValue);
    }
}
