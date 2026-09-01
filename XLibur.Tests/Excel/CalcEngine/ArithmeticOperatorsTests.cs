using XLibur.Excel;
using System;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

public class ArithmeticOperatorsTests
{
    #region Concat text operator

    [Test]
    [Arguments("\"A\" & \"B\"", "AB")]
    [Arguments("\"\" & \"B\"", "B")]
    [Arguments("\"A\" & \"\"", "A")]
    [Arguments("\"\" & \"\"", "")]
    public async Task Concat_ConcatenateText(string formula, object expectedResult)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(ExpectedCellValue.From(expectedResult));
    }

    [Test]
    [Arguments("A1 & \"\"", "")]
    [Arguments("\"\" & A1", "")]
    [Arguments("A1 & A1", "")]
    public async Task Concat_ConcatenateBlank(string formula, object expectedResult)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(ExpectedCellValue.From(expectedResult));
    }

    [Test]
    [Arguments("TRUE & \" to text\"", "TRUE to text")]
    [Arguments("FALSE & \" to text\"", "FALSE to text")]
    [Arguments("true & \" to text\"", "TRUE to text")]
    [Arguments("false & \" to text\"", "FALSE to text")]
    [Arguments("TRUE & FALSE", @"TRUEFALSE")]
    public async Task Concat_ConvertsLogicalToString(string formula, object expectedResult)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(ExpectedCellValue.From(expectedResult));
    }

    [Test]
    [SetCulture("cs-CZ")]
    [Arguments("1 & \" to text\"", "1 to text")]
    [Arguments("1 & 0", "10")]
    [Arguments("1.5 & 0.78", "1,50,78")]
    public async Task Concat_ConvertsNumberToStringUsingCulture(string formula, object expectedResult)
    {
        var wb = new XLWorkbook();
        await Assert.That(wb.Evaluate(formula)).IsEqualTo(ExpectedCellValue.From(expectedResult));
    }

    [Test]
    [Arguments("#DIV/0! & 1", XLError.DivisionByZero)]
    [Arguments("#DIV/0! & \"1\"", XLError.DivisionByZero)]
    [Arguments("#REF! & #DIV/0!", XLError.CellReference)]
    [Arguments("1 & #NAME?", XLError.NameNotRecognized)]
    public async Task Concat_WithErrorAsOperandReturnsTheError(string formula, XLError expectedError)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expectedError);
    }

    #endregion

    #region Unary plus

    [Test]
    [Arguments("+1", 1)]
    [Arguments("+\"1\"", "1")]
    [Arguments("+TRUE", true)]
    [Arguments("+FALSE", false)]
    [Arguments("+#DIV/0!", XLError.DivisionByZero)]
    [Arguments("ISBLANK(+A1)", true)]
    public async Task UnaryPlus_IsNonOpThatKeepsValueAndType(string formula, object expectedValue)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    #endregion

    #region Unary minus

    [Test]
    [Arguments("-1", -1)]
    [Arguments("-125.45", -125.45)]
    [Arguments("-\"1\"", -1)]
    [Arguments("-TRUE", -1)]
    [Arguments("-FALSE", 0)]
    [Arguments("-#DIV/0!", XLError.DivisionByZero)]
    [Arguments("-A1", 0.0)]
    public async Task UnaryMinus_ConvertsArgumentBeforeNegating(string formula, object expectedValue)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    #endregion

    #region Unary minus

    [Test]
    [Arguments("1%", 0.01)]
    [Arguments("100%", 1.0)]
    [Arguments("25.7%", 0.257)]
    [Arguments("125.45%", 1.2545)]
    [Arguments("\"1\"%", 0.01)]
    [Arguments("TRUE%", 0.01)]
    [Arguments("FALSE%", 0)]
    [Arguments("#NAME?%", XLError.NameNotRecognized)]
    [Arguments("(1/0)%", XLError.DivisionByZero)]
    [Arguments("A1%", 0.0)]
    public async Task UnaryPercent_ConvertsArgumentBeforePercentOperator(string formula, object expectedValue)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    #endregion

    #region Exponentiation

    [Test]
    [Arguments("1^1", 1.0)]
    [Arguments("0^0", XLError.NumberInvalid)]
    [Arguments("10^0", 1.0)]
    [Arguments("4^0.5", 2.0)]
    [Arguments("2^0.5", 1.4142135623730951)]
    [Arguments("2^-2", 0.25)]
    [Arguments("\"5\"^\"3\"", 125)]
    [Arguments("5^TRUE", 5)]
    [Arguments("5^FALSE", 1)]
    [Arguments("#VALUE!^1", XLError.IncompatibleValue)]
    [Arguments("1^#REF!", XLError.CellReference)]
    [Arguments("#DIV/0!^#REF!", XLError.DivisionByZero)]
    [Arguments("5^A1", 1.0)]
    [Arguments("A1^4", 0.0)]
    public async Task Exponentiation_CanWorkWithScalars(string formula, object expectedValue)
    {
        var result = Evaluate(formula);
        if (expectedValue is double or int)
            await Assert.That((double)result).IsEqualTo(Convert.ToDouble(expectedValue)).Within(XLHelper.Epsilon);
        else
            await Assert.That(result).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    #endregion

    #region Multiplication

    [Test]
    [Arguments("1+1", 2.0)]
    [Arguments("0*0", 0.0)]
    [Arguments("10*0", 0.0)]
    [Arguments("2*1.5", 3.0)]
    [Arguments("2.5*2.5", 6.25)]
    [Arguments("2*-2", -4)]
    [Arguments("\"5\" * \"3\"", 15)]
    [Arguments("5*TRUE", 5)]
    [Arguments("5*FALSE", 0)]
    [Arguments("#VALUE!*1", XLError.IncompatibleValue)]
    [Arguments("1*#REF!", XLError.CellReference)]
    [Arguments("#DIV/0!*#REF!", XLError.DivisionByZero)]
    [Arguments("10*A1", 0.0)]
    [Arguments("A1*10", 0.0)]
    public async Task Multiplication_CanWorkWithScalars(string formula, object expectedValue)
    {
        var result = Evaluate(formula);
        if (expectedValue is double or int)
            await Assert.That((double)result).IsEqualTo(Convert.ToDouble(expectedValue)).Within(XLHelper.Epsilon);
        else
            await Assert.That(result).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    #endregion

    #region Division

    [Test]
    [Arguments("1/1", 1.0)]
    [Arguments("5/2", 2.5)]
    [Arguments("14.5/2.5", 5.8)]
    [Arguments("10/0", XLError.DivisionByZero)]
    [Arguments("0/0", XLError.DivisionByZero)]
    [Arguments("2.5/-0.5", -5)]
    [Arguments("\"10\" / \"4\"", 2.5)]
    [Arguments("5/TRUE", 5)]
    [Arguments("5/FALSE", XLError.DivisionByZero)]
    [Arguments("#VALUE!/1", XLError.IncompatibleValue)]
    [Arguments("1/#REF!", XLError.CellReference)]
    [Arguments("#DIV/0!/#REF!", XLError.DivisionByZero)]
    [Arguments("A1/5", 0.0)]
    [Arguments("5/A1", XLError.DivisionByZero)]
    public async Task Division_CanWorkWithScalars(string formula, object expectedValue)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    #endregion

    #region Addition

    [Test]
    [Arguments("1+1", 2.0)]
    [Arguments("5+2.5", 7.5)]
    [Arguments("10+0", 10.0)]
    [Arguments("\"10\" + \"4\"", 14.0)]
    [Arguments("5+TRUE", 6.0)]
    [Arguments("5+FALSE", 5.0)]
    [Arguments("#VALUE! + 1", XLError.IncompatibleValue)]
    [Arguments("1 + #REF!", XLError.CellReference)]
    [Arguments("#DIV/0! + #REF!", XLError.DivisionByZero)]
    [Arguments("A1 + 7", 7)]
    public async Task Addition_CanWorkWithScalars(string formula, object expectedValue)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    #endregion

    #region Subtraction

    [Test]
    [Arguments("1-1", 0.0)]
    [Arguments("2.5-7.8", -5.3)]
    [Arguments("10-0", 10.0)]
    [Arguments("\"10\" - \"4\"", 6.0)]
    [Arguments("5-TRUE", 4.0)]
    [Arguments("5-FALSE", 5.0)]
    [Arguments("#VALUE! - 1", XLError.IncompatibleValue)]
    [Arguments("1 - #REF!", XLError.CellReference)]
    [Arguments("#DIV/0! - #REF!", XLError.DivisionByZero)]
    [Arguments("A1 - 5", -5)]
    public async Task Subtraction_CanWorkWithScalars(string formula, object expectedValue)
    {
        await Assert.That(Evaluate(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    #endregion

    #region Array Operations

    [Test]
    public async Task ArraysOperation_BinaryOperationBetweenAreaReferenceAndSingleCellReferenceShouldWork()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Test1");
        ws.Cell("A1").Value = new DateTime(2021, 1, 15, 0, 0, 0, DateTimeKind.Unspecified);
        ws.Cell("A2").Value = new DateTime(2021, 1, 10, 0, 0, 0, DateTimeKind.Unspecified);
        ws.Cell("B1").Value = new DateTime(2021, 1, 5, 0, 0, 0, DateTimeKind.Unspecified);
        await Assert.That(ws.Evaluate("MIN(A1:A2-B1)")).IsEqualTo(5);
    }

    [Test]
    public async Task ArraysOperation_MultiAreaReferencesArgumentResultsInScalarError()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cells("A1:A2").Value = 1;
        await Assert.That(ws.Evaluate("(A1:A1,A1:A2)+1")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(ws.Evaluate("TYPE((A1:A1,A1:A2)+1)")).IsEqualTo(16); // The result is a scalar error, not an array of errors
    }

    /// <summary>
    /// An operand that is a reference should be reduced by implicit intersection before the
    /// operator sees it, the way a scalar function's arguments already are.
    /// </summary>
    /// <remarks>
    /// D38, open. With <c>B1 = 100</c> and <c>B3 = 5</c>, a formula at <c>C3</c> intersects
    /// <c>B:B</c> to <c>B3</c>, so <c>A1+B:B</c> is <c>42 + 5 = 47</c>. XLibur answers
    /// <c>142</c> — it takes the array's first element, <c>B1</c> — and builds the whole
    /// 1,048,576-element array to do it, at roughly 600 ms per full column. The fuzzer found this
    /// as a 366-second timeout on <c>VLO…+A.+UU:B</c>, a 16-byte input spanning 566 columns.
    ///
    /// <para>
    /// A bare <c>B:B</c> at <c>C3</c> already answers <c>5</c> in no measurable time, so the
    /// engine has the right behaviour and does not reach it from an operand. The two-line fix —
    /// intersect both operands in <c>AnyValue.BinaryOperation</c> — fails six tests, of which
    /// <c>MIN(A1:A2-B1)</c> is the instructive one: inside a range-accepting function the operand
    /// must stay an array. See the comment at that call site before attempting it.
    /// </para>
    /// </remarks>
    [Test]
    [Arguments("A1+B:B", 47.0)]
    [Arguments("A1+B1:B10", 47.0)]
    [Arguments("A1*B:B", 210.0)]
    public async Task ImplicitIntersection_AppliesToAnOperandOfABinaryOperator(string formula, double expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = 42;
        ws.Cell("B1").Value = 100;
        ws.Cell("B3").Value = 5;

        await Assert.That(ws.Evaluate(formula, "C3")).IsEqualTo(expected);
    }

    /// <summary>
    /// Every operator kind intersects a reference operand, not only the arithmetic ones.
    /// </summary>
    /// <remarks>
    /// D38. Each expectation was read off Excel 16.0 with the formula in <c>C3</c> and
    /// <c>A1 = 42</c>, <c>B1 = 100</c>, <c>B2 = 7</c>, <c>B3 = 5</c>; Excel stores them as
    /// <c>=A1&amp;@B1:B3</c>, <c>=-@B1:B3</c> and so on.
    /// </remarks>
    [Test]
    [Arguments("A1&B1:B3", "425")]
    [Arguments("A1>B1:B3", true)]
    [Arguments("A1<B1:B3", false)]
    [Arguments("-B1:B3", -5.0)]
    [Arguments("B1:B3%", 0.05)]
    [Arguments("B1:B3+B1:B3", 10.0)]
    public async Task ImplicitIntersection_AppliesToEveryOperatorKind(string formula, object expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = 42;
        ws.Cell("B1").Value = 100;
        ws.Cell("B2").Value = 7;
        ws.Cell("B3").Value = 5;

        await Assert.That(ws.Evaluate(formula, "C3")).IsEqualTo(XLCellValue.FromObject(expected));
    }

    /// <summary>
    /// An operand the formula's own cell does not line up with is <c>#VALUE!</c>, not the operand's
    /// first cell.
    /// </summary>
    /// <remarks>
    /// D38. Excel intersects a column vector on the formula's row and a row vector on its column,
    /// and answers <c>#VALUE!</c> when the formula lies outside that span — all three rows below
    /// were read off Excel 16.0. XLibur used to answer with element <c>[0,0]</c> instead, so
    /// <c>A1+B1:B2</c> in <c>C3</c> was <c>142</c>.
    /// </remarks>
    [Test]
    [Arguments("A1+B1:B2", "C3")] // column vector, formula's row is below it
    [Arguments("A1+B1:B3", "C5")] // column vector, formula's row is below it
    [Arguments("A1+B1:D1", "G3")] // row vector, formula's column is right of it
    public async Task ImplicitIntersection_OperandThatDoesNotSpanTheFormulaIsAnError(string formula, string address)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = 42;
        ws.Cell("B1").Value = 100;
        ws.Cell("B2").Value = 7;
        ws.Cell("B3").Value = 5;

        await Assert.That(ws.Evaluate(formula, address)).IsEqualTo(XLError.IncompatibleValue);
    }

    /// <summary>
    /// A function call inside an operand does not leave the following operator in array context.
    /// </summary>
    /// <remarks>
    /// D38. The visitor clears the intersection flag while a function's arguments are evaluated and
    /// has to put it back afterwards, not assume it was off. Excel stores this as
    /// <c>=A1+SUM(B1:B3)*@B1:B3</c> and answers <c>42 + 112 * 5 = 602</c>; a visitor that failed to
    /// restore would take <c>B1</c> and answer <c>11242</c>.
    /// </remarks>
    [Test]
    [Arguments("A1+SUM(B1:B3)*B1:B3", 602.0)]
    [Arguments("SUM(B1:B3)+B1:B3", 117.0)]
    public async Task ImplicitIntersection_SurvivesAFunctionCallInsideTheExpression(string formula, double expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = 42;
        ws.Cell("B1").Value = 100;
        ws.Cell("B2").Value = 7;
        ws.Cell("B3").Value = 5;

        await Assert.That(ws.Evaluate(formula, "C3")).IsEqualTo(expected);
    }

    /// <summary>
    /// The three places that keep array semantics: a dynamic-array formula, an expression with no
    /// cell to intersect against, and a function's argument.
    /// </summary>
    /// <remarks>
    /// D38. The middle one is forced — implicit intersection needs a row and a column, and
    /// <c>Evaluate(expression)</c> supplies neither. The last one is narrower than Excel, which
    /// intersects inside <c>SUM</c> too and would answer <c>5</c> to the third case; closing that
    /// gap needs a per-parameter array/range classification XLibur does not have. See spec 53.
    /// </remarks>
    [Test]
    public async Task ImplicitIntersection_DoesNotApplyWithoutAScalarContext()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = 42;
        ws.Cell("B1").Value = 100;
        ws.Cell("B2").Value = 7;
        ws.Cell("B3").Value = 5;

        // No formula address: nothing to intersect against, so the array survives and reduces.
        await Assert.That(ws.Evaluate("A1+B1:B3")).IsEqualTo(142.0);

        // Inside a function argument the enclosing function supplies the context.
        await Assert.That(ws.Evaluate("SUM(B1:B3+0)", "C3")).IsEqualTo(112.0);

        // A dynamic-array formula spills, and each element keeps its own operand.
        ws.Cell("D3").SetDynamicFormulaA1("A1+B1:B3");
        await Assert.That(ws.Cell("D3").Value).IsEqualTo(142.0);
        await Assert.That(ws.Cell("D4").Value).IsEqualTo(49.0);
        await Assert.That(ws.Cell("D5").Value).IsEqualTo(47.0);
    }

    [Test]
    public async Task ArrayOperation_SameSizeArrayPerformsOperationIndividually()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("SUM({1,2,3;4,5,6} + {6,5,4;3,2,1})")).IsEqualTo(6 * 7);
        await Assert.That(XLWorkbook.EvaluateExpr("COLUMNS({1,2} + \"A\")")).IsEqualTo(2);
    }

    [Test]
    public async Task ArrayOperation_ArrayPlusScalarUpscalesScalarToSizeOfArray()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("SUM({1,1,1;1,1,1} * 3)")).IsEqualTo(18);
        await Assert.That(XLWorkbook.EvaluateExpr("SUM(6 / {2,2,2;3,3,3})")).IsEqualTo(15);
    }

    [Test]
    public async Task ArrayOperation_RowOnlyArrayIsRepeatedToHaveSameNumberOfRowsAsOtherArray()
    {
        // {3,2} is scaled to {3,2;3,2} of second array
        await Assert.That(XLWorkbook.EvaluateExpr("SUM({3,2}+{1,1;1,1})")).IsEqualTo(14);
        await Assert.That(XLWorkbook.EvaluateExpr("SUM({1,1;1,1}+{3,2})")).IsEqualTo(14);
    }

    [Test]
    public async Task ArrayOperation_ColumnOnlyArrayIsRepeatedToHaveSameNumberOfColumnsAsOtherArray()
    {
        // {3;2} is scaled to {3,3;2,2} of second array
        await Assert.That(XLWorkbook.EvaluateExpr("SUM({3;2}*{1,1;2,3})")).IsEqualTo(16);
        await Assert.That(XLWorkbook.EvaluateExpr("SUM({1,1;2,3}*{3;2})")).IsEqualTo(16);
    }

    [Test]
    public async Task ArrayOperation_1x1ArrayIsScaledToOtherArray()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("SUM({2}*{1,2;3,4})")).IsEqualTo(20);
        await Assert.That(XLWorkbook.EvaluateExpr("SUM({1,2;3,4}*{2})")).IsEqualTo(20);
    }

    [Test]
    public async Task ArrayOperation_DifferentSizedArraysAreUpscaledToContainingSize()
    {
        // The extra value are #N/A + value, i.e. #N/A, thus the whole sum is #N/A
        await Assert.That(XLWorkbook.EvaluateExpr("SUM({1,2;3,4;5,6}+{1,2,3;4,5,6})")).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(XLWorkbook.EvaluateExpr("ROWS({1,2;3,4;5,6}+{1,2,3;4,5,6})")).IsEqualTo(3);
        await Assert.That(XLWorkbook.EvaluateExpr("COLUMNS({1,2;3,4;5,6}+{1,2,3;4,5,6})")).IsEqualTo(3);
    }

    #endregion

    private static XLCellValue Evaluate(string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        return ws.Evaluate(formula);
    }
}
