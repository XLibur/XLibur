using XLibur.Excel;
using System;
using XLibur.Extensions;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

[SetCulture("en-US")]
public class InformationTests
{
    [Test]
    [Arguments("A1")] // blank
    [Arguments("TRUE")]
    [Arguments("14.5")]
    [Arguments("\"text\"")]
    public async Task ErrorType_NonErrorsAreNA(string argumentFormula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.Evaluate($"ERROR.TYPE({argumentFormula})")).IsEqualTo(XLError.NoValueAvailable);
    }

    [Test]
    [Arguments("#NULL!", 1)]
    [Arguments("#DIV/0!", 2)]
    [Arguments("#VALUE!", 3)]
    [Arguments("#REF!", 4)]
    [Arguments("#NAME?", 5)]
    [Arguments("#NUM!", 6)]
    [Arguments("#N/A", 7)]
    //[TestCase("#GETTING_DATA", 8)] OLAP Cube not supported
    // #SPILL! (ERROR.TYPE 9) can't be written as a literal — the parser doesn't tokenize it —
    // so it is covered against a real spilled #SPILL! cell in SpillEvaluationTests.
    public async Task ErrorType_ReturnsNumberForError(string error, int expectedNumber)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"ERROR.TYPE({error})")).IsEqualTo(expectedNumber);
    }

    #region IsBlank Tests

    [Test]
    public async Task IsBlank_EmptyCell_True()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var actual = ws.Evaluate("IsBlank(A1)");
        await Assert.That(actual).IsEqualTo(true);
    }

    [Test]
    public async Task IsBlank_NonEmptyCell_False()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "1";
        var actual = ws.Evaluate("IsBlank(A1)");
        await Assert.That(actual).IsEqualTo(false);
    }

    [Test]
    [Arguments("FALSE")]
    [Arguments("0")]
    [Arguments("5")]
    [Arguments("\"\"")]
    [Arguments("\"Hello\"")]
    [Arguments("#DIV/0!")]
    public async Task IsBlank_NonEmptyValue_False(string value)
    {
        var actual = XLWorkbook.EvaluateExpr($"IsBlank({value})");
        await Assert.That(actual).IsEqualTo(false);
    }

    [Test]
    public async Task IsBlank_InlineBlank_True()
    {
        var actual = XLWorkbook.EvaluateExpr("IsBlank(IF(TRUE,,))");
        await Assert.That(actual).IsEqualTo(true);
    }

    #endregion IsBlank Tests

    [Test]
    [Arguments("IF(TRUE,,)")]
    [Arguments("FALSE")]
    [Arguments("0")]
    [Arguments("\"\"")]
    [Arguments("\"text\"")]
    public async Task IsErr_NonErrorValues_False(string valueFormula)
    {
        var actual = XLWorkbook.EvaluateExpr($"IsErr({valueFormula})");
        await Assert.That(actual).IsEqualTo(false);
    }

    [Test]
    [Arguments("#DIV/0!")]
    [Arguments("#NAME?")]
    [Arguments("#NULL!")]
    [Arguments("#NUM!")]
    [Arguments("#REF!")]
    [Arguments("#VALUE!")]
    public async Task IsErr_ErrorsExceptNA_True(string valueFormula)
    {
        var actual = XLWorkbook.EvaluateExpr($"IsErr({valueFormula})");
        await Assert.That(actual).IsEqualTo(true);
    }

    [Test]
    public async Task IsErr_NA_False()
    {
        var actual = XLWorkbook.EvaluateExpr("IsErr(#N/A)");
        await Assert.That(actual).IsEqualTo(false);
    }

    [Test]
    [Arguments("#DIV/0!")]
    [Arguments("#N/A")]
    [Arguments("#NAME?")]
    [Arguments("#NULL!")]
    [Arguments("#NUM!")]
    [Arguments("#REF!")]
    [Arguments("#VALUE!")]
    public async Task IsError_Errors_True(string error)
    {
        var actual = XLWorkbook.EvaluateExpr($"IsError({error})");
        await Assert.That(actual).IsEqualTo(true);
    }

    [Test]
    [Arguments("IF(TRUE,,)")]
    [Arguments("FALSE")]
    [Arguments("0")]
    [Arguments("\"\"")]
    [Arguments("\"text\"")]
    public async Task IsError_NonErrors_False(string valueFormula)
    {
        var actual = XLWorkbook.EvaluateExpr($"IsError({valueFormula})");
        await Assert.That(actual).IsEqualTo(false);
    }

    #region IsEven Tests

    [Test]
    [Arguments("2")]
    [Arguments("\"1 2/2\"")]
    [Arguments("\"4 1/2\"")]
    [Arguments("\"48:30:00\"")]
    [Arguments("\"1900-01-02\"")]
    public async Task IsEven_NumberLikeValue_ConvertedThroughValueSemantic(string valueFormula)
    {
        var actual = XLWorkbook.EvaluateExpr($"IsEven({valueFormula})");
        await Assert.That(actual).IsEqualTo(true);
    }

    [Test]
    public async Task IsEven_NonIntegerValues_TruncatedForEvaluation()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");

        ws.Cell("A1").Value = 4;
        ws.Cell("A2").Value = 0.9;
        ws.Cell("A3").Value = -2.9;

        var actual = ws.Evaluate("=IsEven(A1)");
        await Assert.That(actual).IsEqualTo(true);

        actual = ws.Evaluate("=IsEven(A2)");
        await Assert.That(actual).IsEqualTo(true);

        actual = ws.Evaluate("=IsEven(A3)");
        await Assert.That(actual).IsEqualTo(true);

        actual = ws.Evaluate("=IsEven(A4)");
        await Assert.That(actual).IsEqualTo(true);
    }

    [Test]
    [Skip("Arrays not yet implemented.")]
    public async Task IsEven_Array_ReturnsArray()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("SUM(N(IsEven({\"2.9\";2;1})))")).IsEqualTo(2.0);
    }

    [Test]
    public async Task IsEven_ReferenceToMoreThanOneCell_Error()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell(1, 2).FormulaA1 = "IsEven(A1:A2)";
        await Assert.That(ws.Cell(1, 2).Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("TRUE", XLError.IncompatibleValue)]
    [Arguments("FALSE", XLError.IncompatibleValue)]
    [Arguments("\"\"", XLError.IncompatibleValue)]
    [Arguments("\"test\"", XLError.IncompatibleValue)]
    [Arguments("#DIV/0!", XLError.DivisionByZero)]
    [Arguments("IF(TRUE,,)", XLError.NoValueAvailable)] // Behaves differently from a reference to a blank cell
    public async Task IsEven_NonNumberValues_Error(string valueFormula, XLError expectedError)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"IsEven({valueFormula})")).IsEqualTo(expectedError);
    }

    #endregion IsEven Tests

    #region IsLogical Tests

    [Test]
    [Arguments("TRUE")]
    [Arguments("FALSE")]
    public async Task IsLogical_OnlyLogical_True(string valueFormula)
    {
        var actual = XLWorkbook.EvaluateExpr($"IsLogical({valueFormula})");
        await Assert.That(actual).IsEqualTo(true);
    }

    [Test]
    [Arguments("IF(TRUE,,)")]
    [Arguments("0")]
    [Arguments("1")]
    [Arguments("\"\"")]
    [Arguments("\"text\"")]
    [Arguments("#NAME?")]
    [Arguments("#N/A")]
    [Arguments("#VALUE!")]
    [Arguments("#REF!")]
    public async Task IsLogical_NonLogicalValue_False(string valueFormula)
    {
        var actual = XLWorkbook.EvaluateExpr($"IsLogical({valueFormula})");
        await Assert.That(actual).IsEqualTo(false);
    }

    [Test]
    public async Task IsLogical_ReferenceToLogicalValue_True()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cell("A1").Value = true;

        var actual = ws.Evaluate("IsLogical(A1)");
        await Assert.That(actual).IsEqualTo(true);
    }

    #endregion IsLogical Tests

    [Test]
    public async Task IsNA_NA_True()
    {
        var actual = XLWorkbook.EvaluateExpr("ISNA(#N/A)");
        await Assert.That(actual).IsEqualTo(true);
    }

    [Test]
    [Arguments("IF(TRUE,,)")]
    [Arguments("TRUE")]
    [Arguments("0")]
    [Arguments("\"\"")]
    [Arguments("#REF!")]
    [Arguments("\"#N/A\"")]
    public async Task IsNA_NonNotAvailableValue_False(string valueFormula)
    {
        var actual = XLWorkbook.EvaluateExpr($"ISNA({valueFormula})");
        await Assert.That(actual).IsEqualTo(false);
    }

    #region IsNotText Tests

    [Test]
    public async Task IsNotText_ReferenceToBlankCell_True()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var actual = ws.Evaluate("IsNonText(A1)");
        await Assert.That(actual).IsEqualTo(true);
    }

    [Test]
    [Arguments("")]
    [Arguments("  ")]
    [Arguments("text")]
    public async Task IsNotText_ReferenceToStringCell_False(string text)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = text;
        var actual = ws.Evaluate("IsNonText(A1)");
        await Assert.That(actual).IsEqualTo(false);
    }

    [Test]
    public async Task IsNotText_NonTextValues_True()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");
        ws.Cell("A1").Value = 123; //Double Value
        ws.Cell("A2").Value = DateTime.Now; //Date Value
        ws.Cell("A3").Value = true; //Bool Value
        ws.Cell("A4").Value = XLError.IncompatibleValue; //Error value

        var actual = ws.Evaluate("IsNonText(A1)");
        await Assert.That(actual).IsEqualTo(true);
        actual = ws.Evaluate("IsNonText(A2)");
        await Assert.That(actual).IsEqualTo(true);
        actual = ws.Evaluate("IsNonText(A3)");
        await Assert.That(actual).IsEqualTo(true);
        actual = ws.Evaluate("IsNonText(A4)");
        await Assert.That(actual).IsEqualTo(true);
    }

    #endregion IsNotText Tests

    #region IsNumber Tests

    [Test]
    public async Task IsNumber_Simple_false()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");
        ws.Cell("A1").Value = "asd"; //String Value
        ws.Cell("A2").Value = true; //Bool Value

        var actual = ws.Evaluate("IsNumber(A1)");
        await Assert.That(actual).IsEqualTo(false);
        actual = ws.Evaluate("IsNumber(A2)");
        await Assert.That(actual).IsEqualTo(false);
    }

    [Test]
    public async Task IsNumber_Simple_true()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");
        ws.Cell("A1").Value = 123; //Double Value
        ws.Cell("A2").Value = DateTime.Now; //Date Value
        ws.Cell("A3").Value = new TimeSpan(2, 30, 50); //TimeSpan Value

        var actual = ws.Evaluate("=IsNumber(A1)");
        await Assert.That(actual).IsEqualTo(true);
        actual = ws.Evaluate("=IsNumber(A2)");
        await Assert.That(actual).IsEqualTo(true);
        actual = ws.Evaluate("=IsNumber(A3)");
        await Assert.That(actual).IsEqualTo(true);
    }

    [Test]
    [Arguments("TRUE")]
    [Arguments("FALSE")]
    [Arguments("\"\"")]
    [Arguments("#DIV/0!")]
    [Arguments("#NULL!")]
    [Arguments("#VALUE!")]
    [Arguments("#N/A")]
    public async Task IsNumber_NonNumber_False(string nonNumberValue)
    {
        var actual = XLWorkbook.EvaluateExpr($"IsNumber({nonNumberValue})");
        await Assert.That(actual).IsEqualTo(false);
    }

    #endregion IsNumber Tests

    #region IsOdd Test

    [Test]
    [SetCulture("en-US")]
    [Arguments("1")]
    [Arguments("\"2 3/3\"")]
    [Arguments("\"5 1/3\"")]
    [Arguments("\"25:30:00\"")]
    [Arguments("\"1900-01-03\"")]
    public async Task IsOdd_SingleValue_ConvertedThroughValueSemantic(string valueFormula)
    {
        var actual = XLWorkbook.EvaluateExpr($"IsOdd({valueFormula})");
        await Assert.That(actual).IsEqualTo(true);
    }

    [Test]
    public async Task IsOdd_NonIntegerValues_TruncatedForEvaluation()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");

        ws.Cell("A1").Value = 3;
        ws.Cell("A2").Value = 1.9;
        ws.Cell("A3").Value = -5.9;

        var actual = ws.Evaluate("=IsOdd(A1)");
        await Assert.That(actual).IsEqualTo(true);

        actual = ws.Evaluate("=IsOdd(A2)");
        await Assert.That(actual).IsEqualTo(true);

        actual = ws.Evaluate("=IsOdd(A3)");
        await Assert.That(actual).IsEqualTo(true);

        actual = ws.Evaluate("=IsOdd(A4)");
        await Assert.That(actual).IsEqualTo(false);
    }

    [SetCulture("en-US")]
    [Test]
    [Skip("Arrays not yet implemented.")]
    public async Task IsOdd_Array_ReturnsArray()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("SUM(N(IsOdd({\"3.2\",7,2})))")).IsEqualTo(2.0);
    }

    [Test]
    public async Task IsOdd_ReferenceToMoreThanOneCell_Error()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell(1, 2).FormulaA1 = "IsOdd(A1:A2)";
        await Assert.That(ws.Cell(1, 2).Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("TRUE", XLError.IncompatibleValue)]
    [Arguments("FALSE", XLError.IncompatibleValue)]
    [Arguments("\"\"", XLError.IncompatibleValue)]
    [Arguments("\"test\"", XLError.IncompatibleValue)]
    [Arguments("#DIV/0!", XLError.DivisionByZero)]
    [Arguments("IF(TRUE,,)", XLError.NoValueAvailable)] // Behaves differently from a reference to a blank cell
    public async Task IsOdd_NonNumberValues_Error(string valueFormula, XLError expectedError)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"IsOdd({valueFormula})")).IsEqualTo(expectedError);
    }

    #endregion IsOdd Test

    [Test]
    [Arguments("A1")]
    [Arguments("(A1,A5)")]
    public async Task IsRef_Reference_True(string reference)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");
        ws.Cell("A1").Value = "123";

        ws.Cell("B1").FormulaA1 = $"ISREF({reference})";

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(true);
    }

    [Test]
    [Arguments("IF(TRUE,,)")]
    [Arguments("TRUE")]
    [Arguments("0")]
    [Arguments("\"\"")]
    // [TestCase("{1;2}")] Arrays not yet implemented
    [Arguments("#N/A")]
    [Arguments("#VALUE!")]
    public async Task IsRef_NonReference_False(string nonReference)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet");

        ws.Cell("B1").FormulaA1 = $"ISREF({nonReference})";

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(false);
    }

    #region IsText Tests

    [Test]
    public async Task IsText_BlankCell_False()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B1").FormulaA1 = "ISTEXT(A1)";

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(false);
    }

    [Test]
    [Arguments("0")]
    [Arguments("123")]
    [Arguments("TRUE")]
    [Arguments("#DIV/0!")]
    [Arguments("IF(TRUE,,)")]
    public async Task IsText_NonText_False(string nonText)
    {
        var actual = XLWorkbook.EvaluateExpr($"ISTEXT({nonText})");
        await Assert.That(actual).IsEqualTo(false);
    }

    [Test]
    [Arguments("")]
    [Arguments("abc")]
    public async Task IsText_CellWithText_True(string textValue)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cell("A1").Value = textValue;

        var actual = ws.Evaluate("IsText(A1)");
        await Assert.That(actual).IsEqualTo(true);
    }

    #endregion IsText Tests

    #region N Tests

    [Test]
    public async Task N_Blank_Zero()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var actual = ws.Evaluate("N(A1)");
        await Assert.That(actual).IsEqualTo(0.0);
    }

    [Test]
    public async Task N_Date_SerialNumber()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var testedDate = DateTime.Now;
        ws.Cell("A1").Value = testedDate;
        var actual = ws.Evaluate("N(A1)");
        await Assert.That(actual).IsEqualTo(testedDate.ToSerialDateTime());
    }

    [Test]
    public async Task N_False_Zero()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = false;
        var actual = ws.Evaluate("N(A1)");
        await Assert.That(actual).IsEqualTo(0);
    }

    [Test]
    public async Task N_True_One()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = true;
        var actual = ws.Evaluate("N(A1)");
        await Assert.That(actual).IsEqualTo(1);
    }
    [Test]
    public async Task N_Number_Number()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var testedValue = 123;
        ws.Cell("A1").Value = testedValue;
        var actual = ws.Evaluate("N(A1)");
        await Assert.That(actual).IsEqualTo(testedValue);
    }

    [Test]
    [Arguments("")]
    [Arguments("abc")]
    public async Task N_String_Zero(string text)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = text;
        var actual = ws.Evaluate("N(A1)");
        await Assert.That(actual).IsEqualTo(0);
    }

    [Test]
    [Skip("Array not implemented")]
    public async Task N_Array_ConvertsIndividualItems()
    {
        var actual = XLWorkbook.EvaluateExpr("SUM(N({2,TRUE}))");
        await Assert.That(actual).IsEqualTo(3);
    }

    [Test]
    [Arguments("A1")]
    [Arguments("A1:B1")]
    [Arguments("(A1, B1)")]
    public async Task N_Reference_TakesFirstCellFromFirstArea(string reference)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 5;
        ws.Cell("B1").Value = 10;

        var actual = ws.Evaluate($"SUM(N({reference}))");
        await Assert.That(actual).IsEqualTo(5);
    }

    #endregion N Tests

    [Test]
    [Arguments("IF(TRUE,,)", 1)]
    [Arguments("0", 1)]
    [Arguments("1", 1)]
    [Arguments("-5.2", 1)]
    [Arguments("\"\"", 2)]
    [Arguments("\"text\"", 2)]
    [Arguments("\"1\"", 2)]
    [Arguments("\"TRUE\"", 2)]
    [Arguments("TRUE", 4)]
    [Arguments("FALSE", 4)]
    [Arguments("#DIV/0!", 16)]
    [Arguments("1/0", 16)]
    [Arguments("#N/A", 16)]
    [Arguments("#VALUE!", 16)]
    public async Task Type_NonReferenceScalarValues(string literalValues, double expectedNumber)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").FormulaA1 = $"TYPE({literalValues})";
        await Assert.That(ws.Cell("A1").Value).IsEqualTo(expectedNumber);
    }

    [Test]
    [Skip("Arrays not implemented")]
    [Arguments("{1}")]
    [Arguments("{TRUE,#N/A}")]
    [Arguments("{\"abc\";5}")]
    public async Task Type_Array_HasValue64(string arrayLiteral)
    {
        var actual = XLWorkbook.EvaluateExpr($"TYPE({arrayLiteral})");
        await Assert.That(actual).IsEqualTo(64.0);
    }

    [Test]
    [Arguments("A1:A2")]
    // [TestCase("(A1:A3 A2:B3)")] Not implemented // Intersection results in a 1x2 block
    public async Task Type_ReferenceToNonSingleCell_BehavesLikeArray(string reference)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("C1").FormulaA1 = $"TYPE({reference})";
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(64.0);
    }

    [Test]
    public async Task Type_ReferenceToSingleCell_ReturnsTypeOfCell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "text";

        ws.Cell("C1").FormulaA1 = "TYPE(A1)";
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(2.0);
    }

    [Test]
    public async Task Type_MultiAreaReference_ReturnsError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "text";

        ws.Cell("C1").FormulaA1 = "TYPE((A1,A1))";
        await Assert.That(ws.Cell("C1").Value).IsEqualTo(16.0);
    }
}
