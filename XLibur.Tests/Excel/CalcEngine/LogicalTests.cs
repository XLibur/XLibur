using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

public class LogicalTests
{
    private static readonly int[] IfTestDataA = [1, 2, 3];
    private static readonly int[] IfTestDataB = [4, 5, 6];
    private static readonly bool[] IfTestDataC = [true, false, true];

    [Test]
    public async Task And_IsLogicalConjunction()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("AND(TRUE)")).IsEqualTo(true);
        await Assert.That(XLWorkbook.EvaluateExpr("AND(TRUE, TRUE)")).IsEqualTo(true);
        await Assert.That(XLWorkbook.EvaluateExpr("AND(TRUE, TRUE, TRUE)")).IsEqualTo(true);
        await Assert.That(XLWorkbook.EvaluateExpr("AND({TRUE, TRUE}, TRUE)")).IsEqualTo(true);

        await Assert.That(XLWorkbook.EvaluateExpr("AND(FALSE)")).IsEqualTo(false);
        await Assert.That(XLWorkbook.EvaluateExpr("AND(TRUE, FALSE)")).IsEqualTo(false);
        await Assert.That(XLWorkbook.EvaluateExpr("AND({TRUE, FALSE})")).IsEqualTo(false);
        await Assert.That(XLWorkbook.EvaluateExpr("AND(TRUE, {TRUE, FALSE})")).IsEqualTo(false);
    }

    [Test]
    [Arguments("A1")]
    [Arguments("A1:A5")]
    [Arguments("(A1:A5,B1:B5)")]
    public async Task And_NoCollectionValues_Error(string range)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.Evaluate($"AND({range})")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task And_ScalarArgumentsCoercedFromBlankOrTextOrNumber()
    {
        // Blank evaluated to false
        await Assert.That(XLWorkbook.EvaluateExpr("AND(IF(TRUE,,))")).IsEqualTo(false);

        // Number coerced to logical
        await Assert.That(XLWorkbook.EvaluateExpr("AND(0)")).IsEqualTo(false);
        await Assert.That(XLWorkbook.EvaluateExpr("AND(0.1)")).IsEqualTo(true);

        // Text coerced to logical
        await Assert.That(XLWorkbook.EvaluateExpr("AND(\"FALSE\")")).IsEqualTo(false);
        await Assert.That(XLWorkbook.EvaluateExpr("AND(\"TRUE\")")).IsEqualTo(true);
    }

    [Test]
    public async Task And_UnconvertableScalarArgumentsSkipped()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("AND(TRUE,\"z\")")).IsEqualTo(true);
    }

    [Test]
    public async Task And_OnlyLogicalOrNumberElementsOfCollectionUsed()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // 0 is a number and is converted to logical
        ws.Cell("A1").Value = 0;
        await Assert.That(ws.Evaluate("AND(TRUE,A1)")).IsEqualTo(false);

        // false is logical
        ws.Cell("A2").Value = false;
        await Assert.That(ws.Evaluate("AND(TRUE,A2)")).IsEqualTo(false);

        // Text is not converted and thus skipped for evaluation
        ws.Cell("A3").Value = "FALSE";
        await Assert.That(ws.Evaluate("AND(TRUE,A3)")).IsEqualTo(true);

        ws.Cell("A4").Value = "some text";
        await Assert.That(ws.Evaluate("AND(TRUE,A4)")).IsEqualTo(true);
    }

    [Test]
    public async Task If_2_Params_true()
    {
        object actual = XLWorkbook.EvaluateExpr(@"if(1 = 1, ""T"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("T"));
    }

    [Test]
    public async Task If_2_Params_false()
    {
        object actual = XLWorkbook.EvaluateExpr(@"if(1 = 2, ""T"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(false));
    }

    [Test]
    public async Task If_3_Params_true()
    {
        object actual = XLWorkbook.EvaluateExpr(@"if(1 = 1, ""T"", ""F"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("T"));
    }

    [Test]
    public async Task If_3_Params_false()
    {
        object actual = XLWorkbook.EvaluateExpr(@"if(1 = 2, ""T"", ""F"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("F"));
    }

    [Test]
    public async Task If_Comparing_Against_Empty_String()
    {
        object actual = XLWorkbook.EvaluateExpr(@"if(date(2016, 1, 1) = """", ""A"",""B"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("B"));

        actual = XLWorkbook.EvaluateExpr(@"if("""" = date(2016, 1, 1), ""A"",""B"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("B"));

        actual = XLWorkbook.EvaluateExpr(@"if("""" = 123, ""A"",""B"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("B"));

        actual = XLWorkbook.EvaluateExpr(@"if("""" = """", ""A"",""B"")");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From("A"));
    }

    [Test]
    public async Task If_Case_Insensitivity()
    {
        object actual = XLWorkbook.EvaluateExpr(@"IF(""text""=""TEXT"", 1, 2)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(1));
    }

    [Test]
    public async Task If_CanReturnReference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.Evaluate("ISREF(IF(TRUE, A1))")).IsEqualTo(true);
        await Assert.That(ws.Evaluate("ISREF(IF(FALSE,, A1))")).IsEqualTo(true);
    }

    [Test]
    public async Task If_has_scalar_condition_and_range_values()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(IfTestDataA);
        ws.Cell("B1").InsertData(IfTestDataB);
        ws.Cell("C1").InsertData(IfTestDataC);
        for (var row = 1; row <= 4; ++row)
            ws.Cell(row, 4).FormulaA1 = "SUM(IF(C1:C3, A1:A3, B1:B3))";

        // Condition is implicitly intersected because it's a scalar parameter
        await Assert.That(ws.Cell("D1").Value).IsEqualTo(6);
        await Assert.That(ws.Cell("D2").Value).IsEqualTo(15);
        await Assert.That(ws.Cell("D3").Value).IsEqualTo(6);
        await Assert.That(ws.Cell("D4").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task If_ConditionError_ReturnError()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"IF(1/0, ""T"", ""F"")")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task If_ConditionCoercedToLogical()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.Evaluate(@"IF(A1, ""T"", ""F"")")).IsEqualTo("F");

        await Assert.That(ws.Evaluate(@"IF(""TRUE"", ""T"", ""F"")")).IsEqualTo("T");
        await Assert.That(ws.Evaluate(@"IF(""FALSE"", ""T"", ""F"")")).IsEqualTo("F");
        await Assert.That(ws.Evaluate(@"IF(""text"", ""T"", ""F"")")).IsEqualTo(XLError.IncompatibleValue);

        await Assert.That(ws.Evaluate(@"IF(1, ""T"", ""F"")")).IsEqualTo("T");
        await Assert.That(ws.Evaluate(@"IF(0, ""T"", ""F"")")).IsEqualTo("F");
    }

    [Test]
    public async Task If_MissingValues_ReturnBlank()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("ISBLANK(IF(TRUE,,))")).IsEqualTo(true);
        await Assert.That(XLWorkbook.EvaluateExpr("ISBLANK(IF(FALSE,,))")).IsEqualTo(true);
    }

    [Test]
    public async Task IfError_FirstArgumentNonError_ReturnFirstArgument()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("ISBLANK(IFERROR(IF(TRUE,), 5))")).IsEqualTo(true);

        await Assert.That(XLWorkbook.EvaluateExpr("IFERROR(FALSE, 5)")).IsEqualTo(false);
        await Assert.That(XLWorkbook.EvaluateExpr("IFERROR(TRUE, 5)")).IsEqualTo(true);

        await Assert.That(XLWorkbook.EvaluateExpr("IFERROR(0, 5)")).IsEqualTo(0.0);
        await Assert.That(XLWorkbook.EvaluateExpr("IFERROR(-2, 5)")).IsEqualTo(-2.0);

        await Assert.That(XLWorkbook.EvaluateExpr("IFERROR(\"\", 5)")).IsEqualTo(string.Empty);
        await Assert.That(XLWorkbook.EvaluateExpr("IFERROR(\"text\", 5)")).IsEqualTo("text");
    }

    [Test]
    public async Task IfError_FirstArgumentError_ReturnSecondArgument()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("IFERROR(1/0, \"text\")")).IsEqualTo("text");

        await Assert.That(XLWorkbook.EvaluateExpr("IFERROR(#REF!, #NAME?)")).IsEqualTo(XLError.NameNotRecognized);
        await Assert.That(XLWorkbook.EvaluateExpr("IFERROR(#NULL!, TRUE)")).IsEqualTo(true);
        await Assert.That(XLWorkbook.EvaluateExpr("ISBLANK(IFERROR(#VALUE!,IF(TRUE,)))")).IsEqualTo(true);
    }

    [Test]
    public async Task IfError_ReferenceNeverReturned()
    {
        // Unlike IF, IFERROR doesn't return reference
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.Evaluate("ISREF(IFERROR(#VALUE!, A1))")).IsEqualTo(false);
    }

    [Test]
    [Arguments("TRUE", false)]
    [Arguments("FALSE", true)]
    [Arguments("IF(TRUE,,)", true)] // Blank
    [Arguments("0", true)]
    [Arguments("0.1", false)]
    [Arguments("\"true\"", false)]
    [Arguments("\"false\"", true)]
    [Arguments("1/0", XLError.DivisionByZero)]
    public async Task Not(string valueFormula, object expectedResult)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"NOT({valueFormula})")).IsEqualTo(ExpectedCellValue.From(expectedResult));
    }

    [Test]
    public async Task Or_IsLogicalDisjunction()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("OR(TRUE)")).IsEqualTo(true);
        await Assert.That(XLWorkbook.EvaluateExpr("OR(TRUE, TRUE)")).IsEqualTo(true);
        await Assert.That(XLWorkbook.EvaluateExpr("OR(TRUE, FALSE, TRUE)")).IsEqualTo(true);
        await Assert.That(XLWorkbook.EvaluateExpr("OR({FALSE, TRUE}, FALSE)")).IsEqualTo(true);

        await Assert.That(XLWorkbook.EvaluateExpr("OR(FALSE)")).IsEqualTo(false);
        await Assert.That(XLWorkbook.EvaluateExpr("OR(FALSE, FALSE)")).IsEqualTo(false);
        await Assert.That(XLWorkbook.EvaluateExpr("OR({FALSE, FALSE})")).IsEqualTo(false);
        await Assert.That(XLWorkbook.EvaluateExpr("OR(FALSE, {FALSE, FALSE})")).IsEqualTo(false);
    }

    [Test]
    [Arguments("A1")]
    [Arguments("A1:A5")]
    [Arguments("(A1:A5,B1:B5)")]
    public async Task Or_NoCollectionValues_Error(string range)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.Evaluate($"OR({range})")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Or_ScalarArgumentsCoercedFromBlankOrTextOrNumber()
    {
        // Blank evaluated to false
        await Assert.That(XLWorkbook.EvaluateExpr("OR(IF(TRUE,,))")).IsEqualTo(false);

        // Number coerced to logical
        await Assert.That(XLWorkbook.EvaluateExpr("OR(0)")).IsEqualTo(false);
        await Assert.That(XLWorkbook.EvaluateExpr("OR(0.1)")).IsEqualTo(true);

        // Text coerced to logical
        await Assert.That(XLWorkbook.EvaluateExpr("OR(\"FALSE\")")).IsEqualTo(false);
        await Assert.That(XLWorkbook.EvaluateExpr("OR(\"TRUE\")")).IsEqualTo(true);
    }

    [Test]
    public async Task Or_UnconvertableScalarArgumentsSkipped()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("OR(TRUE,\"z\")")).IsEqualTo(true);
    }

    [Test]
    public async Task Or_OnlyLogicalOrNumberElementsOfCollectionUsed()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // 1 is a number and is converted to logical
        ws.Cell("A1").Value = 1;
        await Assert.That(ws.Evaluate("OR(FALSE,A1)")).IsEqualTo(true);

        // false is logical
        ws.Cell("A2").Value = true;
        await Assert.That(ws.Evaluate("OR(FALSE,A2)")).IsEqualTo(true);

        // Text is not converted and thus skipped for evaluation
        ws.Cell("A3").Value = "TRUE";
        await Assert.That(ws.Evaluate("OR(FALSE,A3)")).IsEqualTo(false);

        ws.Cell("A4").Value = "some text";
        await Assert.That(ws.Evaluate("OR(FALSE,A4)")).IsEqualTo(false);
    }
}
