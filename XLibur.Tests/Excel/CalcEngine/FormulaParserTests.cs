using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;
/// <summary>
/// Tests that verify that we can parse formulas and evaluate them. Take a look at XLParser ExcelFormulaGrammar.cs and each rule + its transformation into Abstract Syntax Tree is checked here.
/// </summary>
[SetCulture("en-US")]
public class FormulaParserTests
{
    #region Start.Rule

    [Test]
    [Arguments]
    public async Task Formula_string_can_starting_with_an_equal_sign()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("=1")).IsEqualTo(1);
    }

    [Test]
    [Arguments]
    public async Task Formula_string_can_omit_starting_equal_sign()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("1")).IsEqualTo(1);
    }

    [Test]
    [Arguments]
    public async Task Root_formula_string_can_be_union_without_parenthesis()
    {
        // The root of a formula string is pretty much the only place where reference union can be without parenthesis. Elsewhere it must have
        // parentheses to avoid misusing union op (coma) with a separation of arguments in a function call.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A3").Value = 3;

        // Union reference can't be implicitly intersected with Z100, so it returns #VALUE!
        await Assert.That(ws.Evaluate("=A1,A3", "Z100")).IsEqualTo(XLError.IncompatibleValue);
    }

    #endregion

    #region Formula.Rule

    [Test]
    [Arguments]
    public async Task Formula_can_be_reference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Text";
        await Assert.That(ws.Evaluate("=A1")).IsEqualTo("Text");
    }

    [Test]
    [Arguments("=1", 1)]
    [Arguments("=\"text\"", "text")]
    [Arguments("=TRUE", true)]
    public async Task Formula_can_be_constant(string formula, object expectedValue)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    [Test]
    [Arguments("=SUM(1,2)", 3)]
    [Arguments("=2+3", 5)]
    [Arguments("=-3", -3)]
    [Arguments("=150%", 1.5)]
    public async Task Formula_can_be_function_call(string formula, object expectedValue)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    [Test]
    [Arguments]
    public async Task Formula_can_be_constant_array()
    {
        // 1 is determined through implicit intersection (first element)
        await Assert.That(XLWorkbook.EvaluateExpr("={1,2,3;4,5,6}")).IsEqualTo(1);
    }

    [Test]
    [Arguments("=(1)", 1)]
    [Arguments("=(\"text\")", "text")]
    public async Task Formula_can_be_another_formula_in_parenthesis(string formula, object expectedValue)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }
    #endregion

    #region Constant.Rule
    [Test]
    [Arguments("=1", 1)] // int
    [Arguments("=1.5", 1.5)]  // double
    [Arguments("=1.23e2", 123)]
    [Arguments("=1.23e-1", 0.123)]
    [Arguments("=1.23e+3", 1230)]
    [Arguments("=032399977109", 32399977109)] // long
    [Arguments("=9223372036854775808", 9223372036854775808)] // BigInteger (long value + 1)
    public async Task Constant_can_be_number(string formula, double expectedNumber)
    {
        // Irony returns number as an object of various types, e.g. int or double
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expectedNumber);
    }

    [Test]
    [Arguments("=\"text\"", "text")]
    [Arguments("=\"first line\nsecond line\"", "first line\nsecond line")]
    [Arguments("=\"we'll\"", "we'll")]
    [Arguments("=\"use two double quote \"\" to nest quotes\"", "use two double quote \" to nest quotes")]
    public async Task Constant_can_be_text(string formula, string expectedText)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expectedText);
    }

    [Test]
    [Arguments("=TRUE", true)]
    [Arguments("=FALSE", false)]
    [Arguments("=tRuE", true)]
    public async Task Constant_can_be_bool(string formula, bool expectedBool)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expectedBool);
    }

    // #REF! is converted by a different rule, so it is not here.
    [Test]
    [Arguments("#VALUE!", XLError.IncompatibleValue)]
    [Arguments("#DIV/0!", XLError.DivisionByZero)]
    [Arguments("#NAME?", XLError.NameNotRecognized)]
    [Arguments("#N/A", XLError.NoValueAvailable)]
    [Arguments("#NULL!", XLError.NullValue)]
    [Arguments("#NUM!", XLError.NumberInvalid)]
    public async Task Constant_can_be_error(string formula, object expectedError)
    {
        var error = (XLError)XLWorkbook.EvaluateExpr(formula);
        await Assert.That(error).IsEqualTo(ExpectedCellValue.From(expectedError));
    }
    #endregion

    // Function call from XLParser is anything that takes arguments and uses some transformation (e.g. addition, excel function, unary operation..)
    #region FunctionCall.Rule

    [Test]
    [Arguments("=COS(0)", 1)]
    [Arguments("=SUM(1,2,3)", 6)]
    public async Task FunctionCall_can_be_excel_predefined_function(string formula, object expectedValue)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    [Test]
    [Arguments("=+1", 1)]
    [Arguments("=-1", -1)]
    //        [TestCase("=@A1", 1)]
    public async Task FunctionCall_can_be_unary_prefix_operation(string formula, object expectedValue)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    [Test]
    [Arguments("=75%", 0.75)]
    public async Task FunctionCall_can_be_unary_postfix_operation(string formula, object expectedValue)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    [Test]
    [Arguments("=2^3", 8)]
    [Arguments("=4^1.5", 8)]
    [Arguments("=3*2", 6)]
    [Arguments("=6/2", 3)]
    [Arguments("=3/2", 1.5)]
    [Arguments("=1+2", 3)]
    [Arguments("=3-5", -2)]
    [Arguments(@"=""A"" & ""B""", "AB")]
    [Arguments("=2>1", true)]
    [Arguments("=1>2", false)]
    [Arguments("=5=5", true)]
    [Arguments("=1=2", false)]
    [Arguments("=1<2", true)]
    [Arguments("=2<1", false)]
    [Arguments("=2<>1", true)]
    [Arguments("=3<>3", false)]
    [Arguments("=2>=1", true)]
    [Arguments("=2>=2", true)]
    [Arguments("=1>=2", false)]
    [Arguments("=1<=2", true)]
    [Arguments("=1<=1", true)]
    [Arguments("=2<=1", false)]
    public async Task FunctionCall_can_be_binary_infix_operation(string formula, object expectedValue)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }
    #endregion

    #region Argument.Rule

    [Test]
    [Arguments("=PMT(0,1,1000,,1)", -1000)]
    public async Task Empty_arguments_are_passed_to_function(string formula, object expectedValue)
    {
        var result = XLWorkbook.EvaluateExpr(formula);
        if (expectedValue is double or int)
            await Assert.That((double)result).IsEqualTo(Convert.ToDouble(expectedValue)).Within(XLHelper.Epsilon);
        else
            await Assert.That(result).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    #endregion

    #region Reference.Rule

    [Test]
    [Arguments("=A1", 1)]
    [Arguments("=TestRangeName", 5)]
    //        [TestCase("=UndefinedRangeName", Error.NameNotRecognized)]
    public async Task Reference_can_be_reference_item(string formula, object expectedValue)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 5;
        ws.Range("A2:A2").AddToNamed("TestRangeName");

        await Assert.That(ws.Evaluate(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    [Test]
    [Arguments]
    public async Task Reference_can_be_reference_function_call()
    {
        // XLParser considers a limited subset of predefined functions (IF, CHOOSE, INDEX...) to be different from other predefined function because they can return reference.
        await Assert.That(XLWorkbook.EvaluateExpr("=IF(FALSE,1,2)")).IsEqualTo(2);
    }

    [Test]
    [Arguments]
    public async Task Reference_can_be_another_reference_in_parenthesis()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;

        await Assert.That(ws.Evaluate("=(A1)")).IsEqualTo(1);
    }

    [Test]
    [Arguments]
    public async Task Reference_can_be_reference_item_with_prefix()
    {
        using var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Sheet1");
        var ws2 = wb.AddWorksheet("Sheet2");
        ws2.Cell("A1").Value = 1;

        await Assert.That(ws1.Evaluate("=Sheet2!  A1")).IsEqualTo(1);
    }

    [Test]
    [Arguments]
    [Skip("ClosedXML.Parser can't tokenize a dynamic data exchange reference.")]
    public async Task Reference_can_be_dynamic_data_exchange()
    {
        await AssertCanParseButNotEvaluate("=Sdemo123|tik!'id1?req?AAPL_STK_SMART_USD_~/'", "Evaluation of dynamic data exchange is not implemented.");
    }

    #endregion

    #region ReferenceFunctionCall.Rule

    [Test]
    [Arguments]
    public async Task Reference_function_call_can_be_binary_range_of_two_references()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Binary range of two references can't be implicitly intersected with Z100, so it returns #VALUE!
        await Assert.That(ws.Evaluate("A1:A3:C2", "Z100")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments]
    public async Task Reference_function_call_can_be_intersection_of_two_references()
    {
        await AssertCanParseButNotEvaluate("=A1:A3 A2:B2", "Evaluation of range intersection operator is not implemented.");
    }

    [Test]
    [Arguments]
    public async Task Reference_function_call_can_be_union_in_parenthesis()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Union reference can't be implicitly intersected with Z100, so it returns #VALUE!
        await Assert.That(ws.Evaluate("=(A1:A3,A2:B2,B1:B4)", "Z100")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments]
    public async Task Reference_function_call_can_be_reference_function()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("=IF(TRUE,1,2)")).IsEqualTo(1);
    }

    [Test]
    [Arguments]
    public async Task Reference_function_call_can_be_reference_with_spill_range_operator()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // A1 is not a spill anchor, so the spill-range operator resolves to a #REF! error
        // (it parses and evaluates rather than throwing).
        await Assert.That(ws.Evaluate("=A1#")).IsEqualTo(XLError.CellReference);
    }

    #endregion

    #region RefFunctionName.Rule

    [Test]
    [Arguments("=IF(FALSE,1,2)", 2)]
    // [TestCase("=CHOOSE(2,\"A\",\"B\",73)", "B")] Not implemented
    public async Task Ref_function_name_can_be_excel_ref_conditional_function(string formula, object expectedValue)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    [Test]
    [Arguments("=INDEX(A1:B2,1,2)", "Lemons")]
    //[TestCase("=OFFSET(C4,-1,-2)", "Pears")] Not implemented
    [Arguments("=INDIRECT(\"A2\")", "Bananas")]
    public async Task Ref_function_name_can_be_excel_ref_function(string formula, object expectedValue)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = "Apples";
        ws.Cell("B1").Value = "Lemons";
        ws.Cell("A2").Value = "Bananas";
        ws.Cell("B2").Value = "Pears";
        await Assert.That(ws.Evaluate(formula)).IsEqualTo(ExpectedCellValue.From(expectedValue));
    }

    #endregion

    #region ReferenceItem.Rule
    // Reference item is transient and is thus inside the reference

    [Test]
    [Arguments]
    public async Task Reference_item_can_be_cell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;

        await Assert.That(ws.Evaluate("=A1")).IsEqualTo(1);
    }

    [Test]
    [Arguments("TestRange")]
    [Arguments("A1A1")]
    public async Task Reference_item_can_be_named_range(string rangeName)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:C4").SetValue(1).AddToNamed(rangeName);

        await Assert.That(ws.Evaluate($"=SUM({rangeName})")).IsEqualTo(12);
    }

    [Test]
    [Arguments]
    public async Task Reference_item_can_be_vertical_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:C4").SetValue(1);

        await Assert.That(ws.Evaluate("=SUM(A:B)")).IsEqualTo(8);
    }

    [Test]
    [Arguments]
    public async Task Reference_item_can_be_horizontal_range()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("A1:C4").SetValue(1);

        await Assert.That(ws.Evaluate("=SUM(2:2)")).IsEqualTo(3);
    }

    [Test]
    [Arguments]
    public async Task Reference_item_can_be_ref_error()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("#REF!")).IsEqualTo(XLError.CellReference);
    }

    [Test]
    [Arguments]
    public async Task Reference_item_can_be_user_defined_function_call()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("CustomFunction(1)")).IsEqualTo(XLError.NameNotRecognized);
    }

    [Test]
    [Arguments]
    public async Task Reference_item_can_be_structured_reference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertTable([new { Amount = 1 }, new { Amount = 2 }]);

        await Assert.That(ws.Evaluate("SUM(Table1[#Data])")).IsEqualTo(3);
    }

    #endregion

    #region ConstantArray.Rule

    [Test]
    public async Task Const_array_must_have_same_number_of_columns()
    {
        var calcEngine = new XLCalcEngine(CultureInfo.InvariantCulture);
        var ex = await Assert.That(() => calcEngine.Parse("{1;2,3}")).Throws<ExpressionParseException>()!;
        await Assert.That(ex.Message).Contains("Rows of an array don't have same size.");
    }

    [Test]
    public async Task Const_array_cant_contain_implicit_intersection_operator()
    {
        // XLParser allows @ for number through 'PrefixOp + Number'
        var calcEngine = new XLCalcEngine(CultureInfo.InvariantCulture);
        var ex = await Assert.That(() => calcEngine.Parse("{@1}")).Throws<ExpressionParseException>()!;
        await Assert.That(ex.Message).Contains("Unexpected token INTERSECT.");
    }

    [Test]
    [MethodDataSource(nameof(ArrayCases))]
    public async Task Const_array_can_have_only_scalars(string formula, object expected)
    {
        var expectedArray = (ConstArray)expected;
        var calcEngine = new XLCalcEngine(CultureInfo.InvariantCulture);

        var ast = calcEngine.Parse(formula);

        var actual = ((ArrayNode)ast.AstRoot).Value;
        await Assert.That(actual.Width).IsEqualTo(expectedArray.Width);
        await Assert.That(actual.Height).IsEqualTo(expectedArray.Height);
        for (var row = 0; row < actual.Height; ++row)
        {
            for (var col = 0; col < actual.Width; ++col)
            {
                var actualElement = actual[row, col];
                var expectedElement = expectedArray[row, col];
                await Assert.That(actualElement).IsEqualTo(expectedElement);
            }
        }
    }

    public static IEnumerable<object[]> ArrayCases
    {
        get
        {
            yield return
            [
                "{1}",
                new ConstArray(new ScalarValue[,] { { 1 } })
            ];
            yield return
            [
                "{#REF!}",
                new ConstArray(new ScalarValue[,] { { XLError.CellReference } })
            ];
            yield return
            [
                "{1,2,3,4}",
                new ConstArray(new ScalarValue[,] { { 1, 2, 3, 4 } })
            ];
            yield return
            [
                "{1,2;3,4}",
                new ConstArray(new ScalarValue[,] { { 1, 2}, { 3, 4 } })
            ];
            yield return
            [
                "{+1,#REF!,\"Text\";FALSE,#DIV/0!,-1.5}",
                new ConstArray(new ScalarValue[,] { { 1, XLError.CellReference, "Text" }, { false, XLError.DivisionByZero, -1.5 } })
            ];
        }
    }

    #endregion

    #region Prefix.Rule

    // No quotes
    [Test]
    [Arguments("=Sheet5!A1", "Sheet5")]
    [Arguments("=Test_sheet!A1", "Test_sheet")]
    // Sheet with quotes
    [Arguments("='Test Sheet'!A1", "Test Sheet")]
    [Arguments("='Test-Sheet'!A1", "Test-Sheet")]
    [Arguments("='^%>;-+'!A1", "^%>;-+")]
    // Sheet can be named as #REF! error, but sheet reference must be escaped
    [Arguments("='#REF'!A1", "#REF")]
    public async Task Prefix_can_be_sheet_token(string formula, string sheetName)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet(sheetName);
        ws.Cell("A1").Value = 5;
        await Assert.That(ws.Evaluate(formula)).IsEqualTo(5);
    }

    [Test]
    [Arguments("=Sheet1:Sheet5!A1")]
    [Arguments("=Jan:Dec!A1")]
    public async Task Prefix_can_be_sheets_for_3d_reference(string formula)
    {
        await AssertCanParseAndEvaluateToRefError(formula);
    }

    [Test]
    [Arguments("=[1]Sheet4!A1")]
    [Arguments("=INDEX([1]Data!$A$2:$N$1494,MIN(IF([1]Data!$A$2:$N$1494=A2,ROW([1]Data!$A$2:$N$1494)-ROW([1]Data!$A$2)+1)),MATCH(A2,INDEX([1]Data!$A$2:$N$1494,MIN(IF([1]Data!$A$2:$N$1494=A2,ROW([1]Data!$A$2:$N$1494)-ROW([1]Data!$A$2)+1)),0),0)+2)")]
    public async Task Prefix_can_be_file_and_sheet_token(string formula)
    {
        await AssertCanParseAndEvaluateToRefError(formula);
    }

    [Test]
    [Arguments("='[asdf.xlsx]Sec'!A1")]
    [Arguments("='[workbook.xlsx]Sheet1'!A1")]
    public async Task External_workbook_reference_with_filename_does_not_throw(string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").SetFormulaA1(formula);
        await Assert.That(ws.Cell("A1").FormulaA1).IsEqualTo(formula.TrimStart('='));
    }

    #endregion

    private static async Task AssertCanParseButNotEvaluate(string formula, string notSupportedMessage)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var calcEngine = new XLCalcEngine(CultureInfo.InvariantCulture);
        _ = calcEngine.Parse(formula);
        await Assert.That(() => ws.Evaluate(formula, "A1")).Throws<Exception>();
    }

    private static async Task AssertCanParseAndEvaluateToRefError(string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var calcEngine = new XLCalcEngine(CultureInfo.InvariantCulture);
        _ = calcEngine.Parse(formula);
        await Assert.That(ws.Evaluate(formula, "A1")).IsEqualTo(XLError.CellReference);
    }
}
