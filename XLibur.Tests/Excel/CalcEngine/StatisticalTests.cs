
using XLibur.Excel;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

public class StatisticalTests
{
    private const double tolerance = 1e-6;
    private static XLWorkbook workbook;

    [Test]
    public async Task Average()
    {
        double value;
        value = (double)workbook.Evaluate("AVERAGE(-27.5,93.93,64.51,-70.56)");
        await Assert.That(value).IsEqualTo(15.095).Within(tolerance);

        var ws = workbook.Worksheets.First();
        value = (double)ws.Evaluate("AVERAGE(G3:G45)");
        await Assert.That(value).IsEqualTo(49.3255814).Within(tolerance);

        // Column D contains only strings - no average, because non-number types are skipped
        await Assert.That(ws.Evaluate("AVERAGE(D3:D45)")).IsEqualTo(XLError.DivisionByZero);

        // Non-numbers in array are skipped instead of being converted
        await Assert.That(ws.Evaluate("AVERAGE({FALSE, TRUE, \"1\", \"0 0/2\", -1})")).IsEqualTo(-1);

        // Blank value in references are skipped
        ws.Cell("Z1").Value = Blank.Value;
        await Assert.That(ws.Evaluate("AVERAGE(Z1,1)")).IsEqualTo(1);

        await AssertScalarToNumberConversion("AVERAGE", 0.5);
        await AssertAnyErrorIsPropagated("AVERAGE");
    }

    [Test]
    public async Task AverageA()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Examples from specification
        ws.Cell("E1").Value = Blank.Value;
        await Assert.That(ws.Evaluate("AVERAGEA(10, E1)")).IsEqualTo(10);
        ws.Cell("E2").Value = true;
        await Assert.That(ws.Evaluate("AVERAGEA(10, E2)")).IsEqualTo(5.5);
        ws.Cell("E3").Value = false;
        await Assert.That(ws.Evaluate("AVERAGEA(10, E3)")).IsEqualTo(5);

        // Array logical arguments are ignored
        await Assert.That(workbook.Evaluate("AVERAGEA({2,TRUE,TRUE,FALSE,FALSE})")).IsEqualTo(2);

        // Array text arguments are counted as zero (4+2+0+0)/4
        await Assert.That(workbook.Evaluate("AVERAGEA({4, 2, \"hello\", \"10\" })")).IsEqualTo(1.5);

        // Reference argument only counts logical as 0/1, text as 0 and ignores blanks.
        ws.Cell("Z1").Value = Blank.Value; // Not counted
        ws.Cell("Z2").Value = true; // 1
        ws.Cell("Z3").Value = "100"; // 0
        ws.Cell("Z4").Value = "hello"; // 0
        ws.Cell("Z5").Value = 0; // 0
        ws.Cell("Z6").Value = 4; // 4
        await Assert.That((double)ws.Evaluate("AVERAGEA(Z1:Z6)")).IsEqualTo(1);

        await AssertScalarToNumberConversion("AVERAGEA", 0.5);
        await AssertAnyErrorIsPropagated("AVERAGEA");
    }

    [Test]
    [Arguments(6, 10, 0.5, 0.205078125)]
    [Arguments(4, 20, 0.2, 0.2181994)] // p different than 0.5
    [Arguments(0, 5, 0.2, 0.32768)] // 0 out of 5 successes
    [Arguments(0, 0, 0.2, 1)] // 0 out of 0 successes
    [Arguments(1, 1, 0, 0)]
    [Arguments(1, 1, 1, 1)]
    [Arguments(2, 4, 0.5, 0.375)]
    [Arguments(2.9, 4.9, 0.5, 0.375)] // Attempts are floored
    public async Task BinomDist_calculates_non_cumulative_binomial_distribution(double k, double n, double p, double expected)
    {
        var kString = k.ToInvariantString();
        var nString = n.ToInvariantString();
        var pString = p.ToInvariantString();
        var result = (double)XLWorkbook.EvaluateExpr($"BINOMDIST({kString}, {nString}, {pString}, FALSE)");
        await Assert.That(result).IsEqualTo(expected).Within(tolerance);
    }

    [Test]
    [Arguments(6, 10, 0.5, 0.828125)]
    [Arguments(2, 7, 0.3, 0.6470695)]
    [Arguments(0, 7, 0.3, 0.0823543)]
    [Arguments(0, 0, 0.3, 1)]
    [Arguments(0, 0, 1, 1)]
    [Arguments(2, 4, 0.5, 0.6875)]
    [Arguments(2.9, 4.9, 0.5, 0.6875)] // Values are floored
    public async Task BinomDist_calculates_cumulative_binomial_distribution(double k, double n, double p, double expected)
    {
        var kString = k.ToInvariantString();
        var nString = n.ToInvariantString();
        var pString = p.ToInvariantString();
        var result = (double)XLWorkbook.EvaluateExpr($"BINOMDIST({kString}, {nString}, {pString}, TRUE)");
        await Assert.That(result).IsEqualTo(expected).Within(tolerance);
    }

    [Test]
    [Arguments(5, 4, 0.5)] // Five successes out of 4 attempts
    [Arguments(-1, 4, 0.5)] // Negative successes
    [Arguments(0, -1, 0.5)] // Negative attempts
    [Arguments(2, 4, -0.1)] // p < 0
    [Arguments(2, 4, 1.1)] // p > 1
    [Arguments(1E+300, 2E+300, 0.5)] // Too large values
    public async Task BinomDist_returns_num_error_on_invalid_calculations(double k, double n, double p)
    {
        var kString = k.ToInvariantString();
        var nString = n.ToInvariantString();
        var pString = p.ToInvariantString();
        var result = XLWorkbook.EvaluateExpr($"BINOMDIST({kString}, {nString}, {pString}, FALSE)");
        await Assert.That(result).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Count()
    {
        var ws = workbook.Worksheets.First();
        XLCellValue value;
        value = ws.Evaluate("COUNT(D3:D45)");
        await Assert.That(value).IsEqualTo(0);

        value = ws.Evaluate("COUNT(G3:G45)");
        await Assert.That(value).IsEqualTo(43);

        value = ws.Evaluate("COUNT(G:G)");
        await Assert.That(value).IsEqualTo(43);

        value = workbook.Evaluate("COUNT(Data!G:G)");
        await Assert.That(value).IsEqualTo(43);

        // Scalar blank, logical and text is counted as numbers
        await Assert.That(ws.Evaluate("COUNT(IF(TRUE,,),TRUE, FALSE, \"1\")")).IsEqualTo(4);

        // Non-number values in arrays are not counted as numbers.
        await Assert.That(ws.Evaluate("COUNT({TRUE,FALSE,\"1\"})")).IsEqualTo(0);

        // Text is not counted as number.
        await Assert.That(ws.Evaluate("COUNT(\"Hello\")")).IsEqualTo(0);

        // Blank cells are not counted as numbers
        ws.Cell("Z1").Value = Blank.Value;
        await Assert.That(ws.Evaluate("COUNT(Z1)")).IsEqualTo(0);

        // Scalar errors are not propagated
        await Assert.That(ws.Evaluate("COUNT(1, #NULL!)")).IsEqualTo(1);

        // Array errors are not propagated
        await Assert.That(ws.Evaluate("COUNT({1, #NULL!})")).IsEqualTo(1);

        // Reference errors are not propagated
        ws.Cell("Z1").Value = XLError.NullValue;
        await Assert.That(ws.Evaluate("COUNT(Z1)")).IsEqualTo(0);
    }

    [Test]
    public async Task CountA()
    {
        var ws = workbook.Worksheets.First();
        var value = ws.Evaluate("COUNTA(D3:D45)");
        await Assert.That(value).IsEqualTo(43);

        value = ws.Evaluate("COUNTA(G3:G45)");
        await Assert.That(value).IsEqualTo(43);

        value = ws.Evaluate("COUNTA(G:G)");
        await Assert.That(value).IsEqualTo(44);

        value = workbook.Evaluate("COUNTA(Data!G:G)");
        await Assert.That(value).IsEqualTo(44);
    }

    [Test]
    public async Task CountA_counts_non_blank_values()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").Value = 39790;
        ws.Cell("A3").Value = 0;
        ws.Cell("A4").Value = 22.24;
        ws.Cell("A5").Value = "Text";
        ws.Cell("A6").Value = false;
        ws.Cell("A7").Value = true;
        ws.Cell("A8").Value = XLError.DivisionByZero;
        ws.Cell("A9").FormulaA1 = "COUNTA(A1:B8)";
        await Assert.That(ws.Cell("A9").Value).IsEqualTo(7);
    }

    [Test]
    public async Task CountA_on_examples_from_spec()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("COUNTA(1,2,3,4,5)")).IsEqualTo(5);
        await Assert.That(XLWorkbook.EvaluateExpr("COUNTA(1,2,3,4,5)")).IsEqualTo(5);
        await Assert.That(XLWorkbook.EvaluateExpr("COUNTA({1,2,3,4,5},6,\"7\")")).IsEqualTo(7);

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("E2").Value = true;
        await Assert.That(ws.Evaluate("COUNTA(10, E1)")).IsEqualTo(1);
        await Assert.That(ws.Evaluate("COUNTA(10, E2)")).IsEqualTo(2);
    }

    [Test]
    public async Task CountA_accepts_union_references()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A2").Value = 7;
        ws.Cell("B5").Value = false;
        await Assert.That(ws.Evaluate("COUNTA((A1:A4,B4:B7))")).IsEqualTo(2);
    }

    [Test]
    public async Task CountA_doesnt_count_single_blank_cell_reference()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.Evaluate("COUNTA(A1)")).IsEqualTo(0);
    }

    [Test]
    public async Task CountA_counts_blank_argument()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("COUNTA(IF(TRUE,,))")).IsEqualTo(1);
    }

    [Test]
    public async Task CountA_counts_error_arguments()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("COUNTA(#NULL!, #DIV/0!, #VALUE!, #REF!, #NAME?, #NUM!, #N/A)")).IsEqualTo(7);
    }

    [Test]
    public async Task CountA_counts_empty_string()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = string.Empty;
        await Assert.That(ws.Evaluate("COUNTA(A1, \"\")")).IsEqualTo(2);
    }

    [Test]
    public async Task CountBlank()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").Value = 0;
        ws.Cell("A3").Value = 1;
        ws.Cell("A4").Value = false;
        ws.Cell("A5").Value = true;
        ws.Cell("A6").Value = "";
        ws.Cell("A7").Value = "Text";
        ws.Cell("A8").Value = XLError.DivisionByZero;

        // Blank and empty text value is counted as blank
        await Assert.That(ws.Evaluate("COUNTBLANK(A1)")).IsEqualTo(1);
        await Assert.That(ws.Cell("A6").Value).IsEqualTo(string.Empty);
        await Assert.That(ws.Evaluate("COUNTBLANK(A6)")).IsEqualTo(1);

        // Anything else isn't counted as blank
        await Assert.That(ws.Evaluate("COUNTBLANK(A1:A8)")).IsEqualTo(2);

        await Assert.That(ws.Evaluate("COUNTBLANK(A:XFD)")).IsEqualTo(17179869178d);

        // Check that all others argument types. The Excel grammar doesn't allow that,
        // so use IF workaround for that.
        await Assert.That(ws.Evaluate("COUNTBLANK(IF(TRUE,))")).IsEqualTo(XLError.IncompatibleValue); // Blank
        await Assert.That(ws.Evaluate("COUNTBLANK(IF(TRUE,FALSE))")).IsEqualTo(XLError.IncompatibleValue); // Logical
        await Assert.That(ws.Evaluate("COUNTBLANK(IF(TRUE,1))")).IsEqualTo(XLError.IncompatibleValue); // Number
        await Assert.That(ws.Evaluate("COUNTBLANK(IF(TRUE,\"\"))")).IsEqualTo(XLError.IncompatibleValue); // Text
        await Assert.That(ws.Evaluate("COUNTBLANK(IF(TRUE,#DIV/0!))")).IsEqualTo(XLError.DivisionByZero); // Error
        await Assert.That(ws.Evaluate("COUNTBLANK(IF(TRUE,{1}))")).IsEqualTo(XLError.IncompatibleValue); // Array
    }

    [Test]
    public async Task CountIf()
    {
        var ws = workbook.Worksheets.First();
        XLCellValue value;
        value = ws.Evaluate(@"=COUNTIF(D3:D45,""Central"")");
        await Assert.That(value).IsEqualTo(24);

        value = ws.Evaluate(@"=COUNTIF(D:D,""Central"")");
        await Assert.That(value).IsEqualTo(24);

        value = workbook.Evaluate(@"=COUNTIF(Data!D:D,""Central"")");
        await Assert.That(value).IsEqualTo(24);
    }

    [Test]
    [Arguments(@"=COUNTIF(Data!E:E, ""J*"")", 13)]
    [Arguments(@"=COUNTIF(Data!E:E, ""*i*"")", 21)]
    [Arguments(@"=COUNTIF(Data!E:E, ""*in*"")", 9)]
    [Arguments(@"=COUNTIF(Data!E:E, ""*i*l"")", 9)]
    [Arguments(@"=COUNTIF(Data!E:E, ""*i?e*"")", 9)]
    [Arguments(@"=COUNTIF(Data!E:E, ""*o??s*"")", 10)]
    [Arguments(@"=COUNTIF(Data!X1:X1000, """")", 1000)]
    [Arguments(@"=COUNTIF(Data!E1:E44, """")", 1)]
    public async Task CountIf_ConditionWithWildcards(string formula, int expectedResult)
    {
        var ws = workbook.Worksheets.First();

        var value = ws.Evaluate(formula);
        await Assert.That(value).IsEqualTo(expectedResult);
    }

    [Test]
    [Arguments("=COUNTIF(A1:A10, 1)", 1)]
    [Arguments("=COUNTIF(A1:A10, 2.0)", 1)]
    [Arguments(@"=COUNTIF(A1:A10, ""3"")", 2)]
    [Arguments("=COUNTIF(A1:A10, 3)", 2)]
    [Arguments("=COUNTIF(A1:A10, 43831)", 1)]
    [Arguments("=COUNTIF(A1:A10, DATE(2020, 1, 1))", 1)]
    [Arguments("=COUNTIF(A1:A10, TRUE)", 1)]
    public async Task CountIf_MixedData(string formula, int expected)
    {
        // We follow to Excel's convention.
        // Excel treats 1 and TRUE as unequal, but 3 and "3" as equal
        // LibreOffice Calc handles some SUMIF and COUNTIF differently, e.g. it treats 1 and TRUE as equal, but 3 and "3" differently
        var ws = workbook.Worksheet("MixedData");
        await Assert.That(ws.Evaluate(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("x", @"=COUNTIF(A1:A1, ""?"")", 1)]
    [Arguments("x", @"=COUNTIF(A1:A1, ""~?"")", 0)]
    [Arguments("?", @"=COUNTIF(A1:A1, ""~?"")", 1)]
    [Arguments("~?", @"=COUNTIF(A1:A1, ""~?"")", 0)]
    [Arguments("~?", @"=COUNTIF(A1:A1, ""~~~?"")", 1)]
    [Arguments("?", @"=COUNTIF(A1:A1, ""~~?"")", 0)]
    [Arguments("~?", @"=COUNTIF(A1:A1, ""~~?"")", 1)]
    [Arguments("~x", @"=COUNTIF(A1:A1, ""~~?"")", 1)]
    [Arguments("*", @"=COUNTIF(A1:A1, ""~*"")", 1)]
    [Arguments("~*", @"=COUNTIF(A1:A1, ""~*"")", 0)]
    [Arguments("~*", @"=COUNTIF(A1:A1, ""~~~*"")", 1)]
    [Arguments("*", @"=COUNTIF(A1:A1, ""~~*"")", 0)]
    [Arguments("~*", @"=COUNTIF(A1:A1, ""~~*"")", 1)]
    [Arguments("~x", @"=COUNTIF(A1:A1, ""~~*"")", 1)]
    [Arguments("~xyz", @"=COUNTIF(A1:A1, ""~~*"")", 1)]
    public async Task CountIf_MoreWildcards(string cellContent, string formula, int expectedResult)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell(1, 1).Value = cellContent;

        await Assert.That((double)ws.Evaluate(formula)).IsEqualTo(expectedResult);
    }

    [Test]
    [Arguments("=COUNTIFS(B1:D1, \"=Yes\")", 1)]
    [Arguments("=COUNTIFS(B1:B4, \"=Yes\", C1:C4, \"=Yes\")", 2)]
    [Arguments("=COUNTIFS(B4:D4, \"=Yes\", B2:D2, \"=Yes\")", 1)]
    public async Task CountIfs_ReferenceExample1FromExcelDocumentations(
        string formula,
        int expectedOutcome)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell(1, 1).Value = "Davidoski";
        ws.Cell(1, 2).Value = "Yes";
        ws.Cell(1, 3).Value = "No";
        ws.Cell(1, 4).Value = "No";

        ws.Cell(2, 1).Value = "Burke";
        ws.Cell(2, 2).Value = "Yes";
        ws.Cell(2, 3).Value = "Yes";
        ws.Cell(2, 4).Value = "No";

        ws.Cell(3, 1).Value = "Sundaram";
        ws.Cell(3, 2).Value = "Yes";
        ws.Cell(3, 3).Value = "Yes";
        ws.Cell(3, 4).Value = "Yes";

        ws.Cell(4, 1).Value = "Levitan";
        ws.Cell(4, 2).Value = "No";
        ws.Cell(4, 3).Value = "Yes";
        ws.Cell(4, 4).Value = "Yes";

        await Assert.That(ws.Evaluate(formula)).IsEqualTo(expectedOutcome);
    }

    [Test]
    public async Task CountIfs_SingleCondition()
    {
        var ws = workbook.Worksheets.First();
        XLCellValue value;
        value = ws.Evaluate(@"=COUNTIFS(D3:D45,""Central"")");
        await Assert.That(value).IsEqualTo(24);

        value = ws.Evaluate(@"=COUNTIFS(D:D,""Central"")");
        await Assert.That(value).IsEqualTo(24);

        value = workbook.Evaluate(@"=COUNTIFS(Data!D:D,""Central"")");
        await Assert.That(value).IsEqualTo(24);
    }

    [Test]
    [Arguments(@"=COUNTIFS(Data!E:E, ""J*"")", 13)]
    [Arguments(@"=COUNTIFS(Data!E:E, ""*i*"")", 21)]
    [Arguments(@"=COUNTIFS(Data!E:E, ""*in*"")", 9)]
    [Arguments(@"=COUNTIFS(Data!E:E, ""*i*l"")", 9)]
    [Arguments(@"=COUNTIFS(Data!E:E, ""*i?e*"")", 9)]
    [Arguments(@"=COUNTIFS(Data!E:E, ""*o??s*"")", 10)]
    [Arguments(@"=COUNTIFS(Data!X1:X1000, """")", 1000)]
    [Arguments(@"=COUNTIFS(Data!E1:E44, """")", 1)]
    public async Task CountIfs_SingleConditionWithWildcards(string formula, int expectedResult)
    {
        var ws = workbook.Worksheets.First();

        var value = ws.Evaluate(formula);
        await Assert.That(value).IsEqualTo(expectedResult);
    }

    [Test]
    [Arguments("COUNTIFS(H1:I3, 1, D1:F2, 2)")]
    [Arguments("COUNTIFS(A:B, \"A*\", C:C, \">2\")")]
    public async Task CountIfs_returns_error_when_areas_dimensions_are_different(string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.Evaluate(formula)).IsEqualTo(XLError.IncompatibleValue);
    }

    [After(HookType.Class)]
    public static void Dispose()
    {
        workbook.Dispose();
    }

    [Test]
    [Arguments("H3:H45", 7.51126069234216)]
    [Arguments("H:H", 7.51126069234216)]
    [Arguments("Data!H:H", 7.51126069234216)]
    [Arguments("H3:H10", 5.26214814727941)]
    [Arguments("H3:H20", 7.01281435054797)]
    [Arguments("H3:H30", 7.00137389296182)]
    [Arguments("H3:H3", 1.99)]
    [Arguments("H10:H20", 8.37855107505682)]
    [Arguments("H15:H20", 15.8927310267677)]
    [Arguments("H20:H30", 7.14321227391814)]
    public async Task Geomean_calculation(string sourceValue, double expected)
    {
        await Assert.That((double)workbook.Worksheets.First().Evaluate($"GEOMEAN({sourceValue})")).IsEqualTo(expected).Within(1e-12);
    }

    [Test]
    [Arguments("D3:D45", XLError.NumberInvalid)]
    [Arguments("-1, 0, 3", XLError.NumberInvalid)]
    [Arguments("0", XLError.NumberInvalid)]
    public async Task Geomean_IncorrectCases(string sourceValue, XLError expected)
    {
        var ws = workbook.Worksheets.First();
        await Assert.That((XLError)ws.Evaluate($"GEOMEAN({sourceValue})")).IsEqualTo(expected);
    }

    [Test]
    public async Task Geomean()
    {
        // Example from the specification
        await Assert.That((double)XLWorkbook.EvaluateExpr("GEOMEAN(10.5,5.3,2.9)")).IsEqualTo(5.4444547024966).Within(1e-8);
        await Assert.That((double)XLWorkbook.EvaluateExpr("GEOMEAN(10.5,{5.3,2.9},\"12\")")).IsEqualTo(6.6337805880630).Within(1e-8);

        // GEOMEAN isn't limited by double scale, i.e. it doesn't use naive algorithm for large number.
        await Assert.That((double)XLWorkbook.EvaluateExpr("GEOMEAN(1E+307, 1E+307)")).IsEqualTo(1.0000000000000231E+307d).Within(1e-8);

        // Scalar blank is counted as a 0
        await Assert.That(XLWorkbook.EvaluateExpr("GEOMEAN(IF(TRUE,), 1)")).IsEqualTo(XLError.NumberInvalid);

        // Scalar logical and text is converted to numbers
        await Assert.That((double)XLWorkbook.EvaluateExpr("GEOMEAN(TRUE, \"5\")")).IsEqualTo(2.236067977).Within(1e-8);

        // Non-number values in arrays are ignored.
        await Assert.That((double)XLWorkbook.EvaluateExpr("GEOMEAN({TRUE, FALSE, \"1\", 7}, 5)")).IsEqualTo(5.916079783).Within(1e-8);

        // Scalar non-number text causes an error due to conversion.
        await Assert.That(XLWorkbook.EvaluateExpr("GEOMEAN(\"Hello\", 5)")).IsEqualTo(XLError.IncompatibleValue);

        // Reference non-number arguments are ignored
        var ws = workbook.Worksheets.First();
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = "1";
        ws.Cell("Z3").Value = "hello";
        ws.Cell("Z4").Value = false;
        ws.Cell("Z5").Value = true;
        ws.Cell("Z6").Value = 5;
        await Assert.That((double)ws.Evaluate("GEOMEAN(Z1:Z6)")).IsEqualTo(5).Within(1e-8);

        await AssertAnyErrorIsPropagated("GEOMEAN");
    }

    [Before(HookType.Test)]
    public void Init()
    {
        // Make sure tests run on a deterministic culture
        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        workbook = SetupWorkbook();
    }

    [Test]
    [Arguments("H3:H45", 94145.5271162791)]
    [Arguments("H:H", 94145.5271162791)]
    [Arguments("Data!H:H", 94145.5271162791)]
    [Arguments("H3:H10", 411.5)]
    [Arguments("H3:H20", 13604.2067611111)]
    [Arguments("H3:H30", 14231.0694)]
    [Arguments("H3:H3", 0)]
    [Arguments("H10:H20", 12713.7600909091)]
    [Arguments("H15:H20", 10827.2200833333)]
    [Arguments("H20:H30", 477.132272727273)]
    public async Task DevSq(string sourceValue, double expected)
    {
        await Assert.That((double)workbook.Worksheets.First().Evaluate($"DEVSQ({sourceValue})")).IsEqualTo(expected).Within(1e-10);
    }

    [Test]
    [Arguments("D3:D45", XLError.NumberInvalid)]
    public async Task Devsq_IncorrectCases(string sourceValue, XLError expected)
    {
        var ws = workbook.Worksheets.First();
        await Assert.That((XLError)ws.Evaluate($"DEVSQ({sourceValue})")).IsEqualTo(expected);
    }

    [Test]
    public async Task Devsq_is_calculated_from_numbers()
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr("DEVSQ(5.6, 8.2, 9.2)")).IsEqualTo(6.90666666666666).Within(1e-10);
        await Assert.That((double)XLWorkbook.EvaluateExpr("DEVSQ({ 5.6, 8.2, 9.2})")).IsEqualTo(6.90666666666666).Within(1e-10);

        // Array logical arguments are ignored
        await Assert.That(workbook.Evaluate("DEVSQ({2,TRUE,TRUE,FALSE,FALSE})")).IsEqualTo(0);
        await Assert.That((double)workbook.Evaluate("DEVSQ({2, 1, 1, 0, 0})")).IsEqualTo(2.8).Within(1e-10);

        // Array text arguments are ignored
        await Assert.That(workbook.Evaluate("DEVSQ({4, 2, \"hello\", \"10\" })")).IsEqualTo(2);

        // Non-numerical reference values are ignored.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = Blank.Value; // Ignored
        ws.Cell("A2").Value = true; // Ignored
        ws.Cell("A3").Value = "100"; // Ignored
        ws.Cell("A4").Value = "hello"; // Ignored
        ws.Cell("A5").Value = 2; // Included
        ws.Cell("A6").Value = 4; // Included
        await Assert.That(ws.Evaluate("DEVSQ(A1:A6)")).IsEqualTo(2);

        await AssertScalarToNumberConversion("DEVSQ", 0.5);
        await AssertAnyErrorIsPropagated("DEVSQ");
    }

    [Test]
    [Arguments(0, 0)]
    [Arguments(0.2, 0.202732554054082)]
    [Arguments(0.25, 0.255412811882995)]
    [Arguments(0.3296001056, 0.342379555936801)]
    [Arguments(-0.36, -0.37688590118819)]
    [Arguments(-0.000003, -0.00000299999999998981)]
    [Arguments(-0.063453535345348, -0.0635389037459617)]
    [Arguments(0.559015883901589171354964, 0.631400600322212)]
    [Arguments(0.2691496, 0.275946780611959)]
    [Arguments(-0.10674142, -0.107149608461448)]
    public async Task Fisher(double sourceValue, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"FISHER({sourceValue})")).IsEqualTo(expected).Within(1e-12);
    }

    [Test]
    [Arguments("\"asdf\"", XLError.IncompatibleValue)]
    [Arguments("5", XLError.NumberInvalid)]
    [Arguments("-1", XLError.NumberInvalid)]
    [Arguments("1", XLError.NumberInvalid)]
    public async Task Fisher_IncorrectCases(string sourceValue, XLError expected)
    {
        await Assert.That((XLError)XLWorkbook.EvaluateExpr($"FISHER({sourceValue})")).IsEqualTo(expected);
    }

    [Test]
    [Arguments(0.05464, 60, 1.960041187)] // From Microsoft documentation
    [Arguments(0.05, 10, 2.228138852)]
    [Arguments(0.1, 30, 1.697260887)]
    [Arguments(0.5, 1, 1.0)]
    [Arguments(0.01, 5, 4.032142984)]
    [Arguments(0.9, 100, 0.125980882)] // T.INV.2S(0.9, 100) = T.INV(0.55, 100)
    public async Task TInv_returns_two_tailed_inverse(double probability, double df, double expected)
    {
        var result = (double)XLWorkbook.EvaluateExpr($"TINV({probability.ToInvariantString()}, {df.ToInvariantString()})");
        await Assert.That(result).IsEqualTo(expected).Within(1e-6);
    }

    [Test]
    [Arguments(0, 10, XLError.NumberInvalid)] // probability = 0
    [Arguments(1, 10, XLError.NumberInvalid)] // probability = 1
    [Arguments(-0.1, 10, XLError.NumberInvalid)] // probability < 0
    [Arguments(1.1, 10, XLError.NumberInvalid)] // probability > 1
    [Arguments(0.5, 0, XLError.NumberInvalid)] // df < 1
    [Arguments(0.5, 0.5, XLError.NumberInvalid)] // df rounds down to 0
    public async Task TInv_returns_error_on_invalid_input(double probability, double df, XLError expected)
    {
        await Assert.That((XLError)XLWorkbook.EvaluateExpr($"TINV({probability.ToInvariantString()}, {df.ToInvariantString()})")).IsEqualTo(expected);
    }

    [Test]
    [Arguments(0.05464, 60, 1.960041187)]
    [Arguments(0.05, 10, 2.228138852)]
    public async Task TInv2T_returns_same_as_TInv(double probability, double df, double expected)
    {
        // T.INV.2T is the modern equivalent of TINV
        var result = (double)XLWorkbook.EvaluateExpr($"T.INV.2T({probability.ToInvariantString()}, {df.ToInvariantString()})");
        await Assert.That(result).IsEqualTo(expected).Within(1e-6);
    }

    [Test]
    [Arguments(0.5, 10, 0.0)] // Median is 0 for symmetric distribution
    [Arguments(0.975, 60, 2.000297822)] // T.INV(0.975, 60) — TINV(0.05, 60)
    [Arguments(0.025, 60, -2.000297822)] // Negative tail
    [Arguments(0.95, 10, 1.812461123)]
    [Arguments(0.05, 10, -1.812461123)]
    public async Task TInv_one_tailed_returns_left_inverse(double probability, double df, double expected)
    {
        var result = (double)XLWorkbook.EvaluateExpr($"T.INV({probability.ToInvariantString()}, {df.ToInvariantString()})");
        await Assert.That(result).IsEqualTo(expected).Within(1e-6);
    }

    [Test]
    [Arguments(0, 10, XLError.NumberInvalid)]
    [Arguments(1, 10, XLError.NumberInvalid)]
    [Arguments(0.5, 0, XLError.NumberInvalid)]
    public async Task TInv_one_tailed_returns_error_on_invalid_input(double probability, double df, XLError expected)
    {
        await Assert.That((XLError)XLWorkbook.EvaluateExpr($"T.INV({probability.ToInvariantString()}, {df.ToInvariantString()})")).IsEqualTo(expected);
    }

    [Test]
    public async Task TInv_user_example()
    {
        // The exact example from the issue
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sample Sheet");

        ws.Cell("A1").Value = "0.05464";
        ws.Cell("A2").Value = "60";
        ws.Cell("B1").FormulaA1 = "=TINV(A1,A2)";

        await Assert.That((double)ws.Cell("B1").Value).IsEqualTo(1.960041187).Within(1e-6);
    }

    [Test]
    public async Task Max()
    {
        var ws = workbook.Worksheets.First();
        XLCellValue value;
        value = ws.Evaluate("=MAX(D3:D45)");
        await Assert.That(value).IsEqualTo(0);

        value = ws.Evaluate("=MAX(G3:G45)");
        await Assert.That(value).IsEqualTo(96);

        value = ws.Evaluate("=MAX(G:G)");
        await Assert.That(value).IsEqualTo(96);

        value = workbook.Evaluate("=MAX(Data!G:G)");
        await Assert.That(value).IsEqualTo(96);

        // Although in most cases blank cells are considered 0, MAX just ignores them.
        value = workbook.Evaluate("MAX(-10, Data!X:Z)");
        await Assert.That(value).IsEqualTo(-10);

        // Arrays - numbers are used
        value = workbook.Evaluate("MAX(-10, { -6, -5, 7 })");
        await Assert.That(value).IsEqualTo(7);

        // Arrays - non-number and non-error values are skipped.
        value = workbook.Evaluate(@"MAX(-10, { TRUE, FALSE, ""100"" })");
        await Assert.That(value).IsEqualTo(-10);

        // Reference argument ignores everything but number.
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = true;
        ws.Cell("Z3").Value = "100";
        ws.Cell("Z4").Value = "hello";
        ws.Cell("Z5").Value = -4;
        await Assert.That(ws.Evaluate("MAX(Z1:Z5)")).IsEqualTo(-4);

        await AssertScalarToNumberConversion("MAX", 1);
        await AssertAnyErrorIsPropagated("MAX");
    }

    [Test]
    public async Task MaxA()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Examples from specification
        await Assert.That(ws.Evaluate("MAXA(10.4,-3.5,12.6)")).IsEqualTo(12.6);
        await Assert.That(ws.Evaluate("MAXA(10.4,{-3.5,12.6})")).IsEqualTo(12.6);
        await Assert.That(ws.Evaluate("MAXA({\"ABC\",TRUE})")).IsEqualTo(0);
        ws.Cell("B3").Value = Blank.Value;
        await Assert.That(ws.Evaluate("MAX(-10,-12,-15,B3)")).IsEqualTo(-10);
        ws.Cell("B3").Value = 0;
        await Assert.That(ws.Evaluate("MAXA(-10,-12,-15,B3)")).IsEqualTo(0);

        // Array logical arguments are ignored
        await Assert.That(workbook.Evaluate("MAXA({-2, TRUE, TRUE, FALSE, FALSE})")).IsEqualTo(-2);

        // Array text arguments are ignored
        await Assert.That(workbook.Evaluate("MAXA({-4, -2, \"hello\", \"10\" })")).IsEqualTo(-2);

        // Reference argument only counts logical as 0/1, text as 0 and ignores blanks.
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").Value = true;
        ws.Cell("A3").Value = "100";
        ws.Cell("A4").Value = "hello";
        ws.Cell("A5").Value = -4;
        await Assert.That(ws.Evaluate("MAXA(A1:A5)")).IsEqualTo(1);
        await Assert.That(ws.Evaluate("MAXA(A3:A5)")).IsEqualTo(0);

        await AssertScalarToNumberConversion("MAXA", 1);
        await AssertAnyErrorIsPropagated("MAXA");
    }

    [Test]
    public async Task Median_with_area_without_numeric_values_returns_error()
    {
        var ws = workbook.Worksheets.First();

        // Column D contains names of regions
        await Assert.That(ws.Evaluate("MEDIAN(D3:D45)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Median_EvenCountOfCellRange_ReturnsAverageOfTwoElementsInMiddleOfSortedList()
    {
        // Arrange
        var ws = workbook.Worksheets.First();

        // Act
        var value = (double)ws.Evaluate("MEDIAN(I3:I10)");

        // Assert
        await Assert.That(value).IsEqualTo(244.225).Within(tolerance);
    }

    [Test]
    public async Task Median_EvenCountOfManualNumbers_ReturnsAverageOfTwoElementsInMiddleOfSortedList()
    {
        // Act
        var value = (double)workbook.Evaluate("MEDIAN(-27.5,93.93,64.51,-70.56)");

        // Assert
        await Assert.That(value).IsEqualTo(18.505).Within(tolerance);
    }

    [Test]
    public async Task Median_OddCountOfCellRange_ReturnsElementInMiddleOfSortedList()
    {
        // Arrange
        var ws = workbook.Worksheets.First();

        // Act
        var value = (double)ws.Evaluate("MEDIAN(I3:I11)");

        // Assert
        await Assert.That(value).IsEqualTo(189.05).Within(tolerance);
    }

    [Test]
    public async Task Median_OddCountOfManualNumbers_ReturnsElementInMiddleOfSortedList()
    {
        // Act
        var value = (double)workbook.Evaluate("MEDIAN(-27.5,93.93,64.51,-70.56,101.65)");

        // Assert
        await Assert.That(value).IsEqualTo(64.51).Within(tolerance);
    }

    [Test]
    public async Task Median_uses_only_numbers()
    {
        var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Examples from specification
        await Assert.That(ws.Evaluate("MEDIAN(10, 20)")).IsEqualTo(15);
        await Assert.That(ws.Evaluate("MEDIAN(-3.5, 1.4, 6.9, -4.5)")).IsEqualTo(-1.05);
        await Assert.That(ws.Evaluate("MEDIAN({ -3.5,1.4,6.9},-4.5)")).IsEqualTo(-1.05);

        // Reference with no value will return error
        ws.Cell("A1").Value = Blank.Value;
        await Assert.That(ws.Evaluate("MEDIAN(A1)")).IsEqualTo(XLError.NumberInvalid);

        // Array non-number values are ignored
        await Assert.That(ws.Evaluate("MEDIAN({7, TRUE,FALSE,\"1\"})")).IsEqualTo(7);

        // Only numbers are used from reference, rest is ignored
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").Value = true;
        ws.Cell("A3").Value = "100";
        ws.Cell("A4").Value = "hello";
        ws.Cell("A5").Value = 0;
        ws.Cell("A6").Value = 4;
        ws.Cell("A7").Value = 5;
        await Assert.That(ws.Evaluate("MEDIAN(A1:A7)")).IsEqualTo(4);

        await AssertScalarToNumberConversion("MEDIAN", 0.5);
        await AssertAnyErrorIsPropagated("MEDIAN");
    }

    [Test]
    public async Task Min()
    {
        var ws = workbook.Worksheets.First();
        await Assert.That(ws.Evaluate("MIN(D3:D45)")).IsEqualTo(0);
        await Assert.That(ws.Evaluate("MIN(G3:G45)")).IsEqualTo(2);
        await Assert.That(ws.Evaluate("MIN(G:G)")).IsEqualTo(2);
        await Assert.That(workbook.Evaluate("MIN(Data!G:G)")).IsEqualTo(2);

        // Array non-number arguments are ignored
        await Assert.That(workbook.Evaluate("MIN({5, TRUE, FALSE, \"1\", \"hello\"})")).IsEqualTo(5);

        // Reference non-number arguments are ignored
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = "1";
        ws.Cell("Z3").Value = "hello";
        ws.Cell("Z4").Value = false;
        ws.Cell("Z5").Value = true;
        ws.Cell("Z6").Value = 5;
        await Assert.That(ws.Evaluate("MIN(Z1:Z6)")).IsEqualTo(5);

        // If there is no value, return 0
        await Assert.That(ws.Evaluate("MIN({\"hello\"})")).IsEqualTo(0);

        await AssertScalarToNumberConversion("MIN", 0);
        await AssertAnyErrorIsPropagated("MIN");
    }

    [Test]
    public async Task MinA()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Examples from specification
        await Assert.That(ws.Evaluate("MINA(10.4, -3.5, 12.6)")).IsEqualTo(-3.5);
        await Assert.That(ws.Evaluate("MINA(10.4, {-3.5, 12.6})")).IsEqualTo(-3.5);
        await Assert.That(ws.Evaluate("MINA({\"ABC\", TRUE})")).IsEqualTo(0);
        ws.Cell("B3").Value = Blank.Value;
        await Assert.That(ws.Evaluate("MINA(10, 12, 15, B3)")).IsEqualTo(10);
        ws.Cell("B3").Value = "Text";
        await Assert.That(ws.Evaluate("MINA(10, 12, 15, B3)")).IsEqualTo(0);

        // Blanks in references are ignored and when MINA doesn't have any values, it returns 0
        ws.Cell("A1").Value = Blank.Value;
        await Assert.That(ws.Evaluate("MINA(A1)")).IsEqualTo(0);

        // Array logical arguments are ignored
        await Assert.That(wb.Evaluate("MINA({2, TRUE, TRUE, FALSE, FALSE})")).IsEqualTo(2);

        // Array text arguments are ignored
        await Assert.That(wb.Evaluate("MINA({4, 2, \"hello\", \"1\"})")).IsEqualTo(2);

        // Reference argument only counts logical as 0/1, text as 0 and ignores blanks.
        ws.Cell("A1").Value = Blank.Value; // Ignores
        ws.Cell("A2").Value = true; // Includes
        ws.Cell("A3").Value = "100"; // Considers 0
        ws.Cell("A4").Value = "hello"; // Considers 0
        ws.Cell("A5").Value = -4; // Included
        await Assert.That(ws.Evaluate("MINA(A1:A2)")).IsEqualTo(1);
        await Assert.That(ws.Evaluate("MINA(A1:A3)")).IsEqualTo(0);
        await Assert.That(ws.Evaluate("MINA(A1:A5)")).IsEqualTo(-4);

        await AssertScalarToNumberConversion("MINA", 0);
        await AssertAnyErrorIsPropagated("MINA");
    }

    [Test]
    public async Task StDev()
    {
        var ws = workbook.Worksheets.First();

        // Only non-convertible text in D column, thus less than 2 samples will return error
        await Assert.That(ws.Evaluate("STDEV(D3:D45)")).IsEqualTo(XLError.DivisionByZero);

        // Calculate StDev from numeric values (reference contains only numbers)
        var value = (double)ws.Evaluate("STDEV(H3:H45)");
        await Assert.That(value).IsEqualTo(47.34511769).Within(tolerance).Within(tolerance);

        // Ignores text values in the H column and only uses numeric ones, same as reference with only number
        value = (double)ws.Evaluate("STDEV(H:H)");
        await Assert.That(value).IsEqualTo(47.34511769).Within(tolerance).Within(tolerance);

        value = (double)workbook.Evaluate("STDEV(Data!H:H)");
        await Assert.That(value).IsEqualTo(47.34511769).Within(tolerance).Within(tolerance);

        // Need at least two values, otherwise returns error
        await Assert.That(workbook.Evaluate("STDEV(1)")).IsEqualTo(XLError.DivisionByZero);
        await Assert.That(workbook.Evaluate("STDEV(0, 0)")).IsEqualTo(0);

        // Array non-number arguments are ignored
        await Assert.That((double)workbook.Evaluate("STDEV({0, 1, \"Hello\", FALSE, TRUE})")).IsEqualTo(0.707106781).Within(tolerance).Within(tolerance);

        // Reference argument only uses number, ignores blanks, logical and text
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = true;
        ws.Cell("Z3").Value = "100";
        ws.Cell("Z4").Value = "hello";
        ws.Cell("Z5").Value = 0;
        ws.Cell("Z6").Value = 1;
        await Assert.That((double)ws.Evaluate("STDEV(Z1:Z6)")).IsEqualTo(0.707106781).Within(tolerance).Within(tolerance);

        await AssertScalarToNumberConversion("STDEV", 0.707106781);
        await AssertAnyErrorIsPropagated("STDEV");
    }

    [Test]
    public async Task StDevA()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Example from specification
        await Assert.That((double)ws.Evaluate("STDEVA(123, 134, 143, 173, 112, 109)")).IsEqualTo(23.72902583).Within(tolerance);

        // Array non-number arguments are ignored
        await Assert.That((double)ws.Evaluate("STDEVA({0, 1, \"9\", \"Hello\", FALSE, TRUE})")).IsEqualTo(0.707106781).Within(tolerance);

        // Reference argument ignores blanks, uses numbers, logical and text as zero
        ws.Cell("A1").Value = Blank.Value; // Ignore
        ws.Cell("A2").Value = true; // Include
        ws.Cell("A3").Value = ""; // Consider 0
        ws.Cell("A4").Value = "100"; // Consider 0
        ws.Cell("A5").Value = "hello"; // Consider 0
        ws.Cell("A6").Value = 5;
        ws.Cell("A7").Value = 7;
        await Assert.That((double)ws.Evaluate("STDEVA(A1:A7)")).IsEqualTo(3.060501048).Within(tolerance);

        // Need at least one sample, otherwise returns error (text in array is ignored)
        await Assert.That(ws.Evaluate("STDEVA({\"hello\"})")).IsEqualTo(XLError.DivisionByZero);

        await AssertScalarToNumberConversion("STDEVA", 0.707106781);
        await AssertAnyErrorIsPropagated("STDEVA");
    }

    [Test]
    public async Task StDevP()
    {
        var ws = workbook.Worksheets.First();

        // Example from specification
        await Assert.That((double)ws.Evaluate("STDEVP(123, 134, 143, 173, 112, 109)")).IsEqualTo(21.66153785).Within(tolerance);

        // Column D contains only region names (non-convertible text), thus reference contains less than 1 sample that is required
        await Assert.That(ws.Evaluate("STDEVP(D3:D45)")).IsEqualTo(XLError.DivisionByZero);

        // Calculate StDevP from numeric values (reference contains only numbers)
        await Assert.That((double)ws.Evaluate("STDEVP(H3:H45)")).IsEqualTo(46.79135458).Within(tolerance);

        // StDevP ignores text values/blanks in the H column and only uses numeric ones, the result is same as the reference above that contains only numbers
        await Assert.That((double)ws.Evaluate("STDEVP(H:H)")).IsEqualTo(46.79135458).Within(tolerance);

        await Assert.That((double)workbook.Evaluate("STDEVP(Data!H:H)")).IsEqualTo(46.79135458).Within(tolerance);

        // If sample size is 0, return error
        await Assert.That(workbook.Evaluate("STDEVP({TRUE})")).IsEqualTo(XLError.DivisionByZero);
        await Assert.That(workbook.Evaluate("STDEVP(100)")).IsEqualTo(0);

        // Array non-number arguments are ignored
        await Assert.That(workbook.Evaluate("STDEVP({0, 1, \"Hello\", FALSE, TRUE})")).IsEqualTo(0.5);

        // Reference argument only uses numbers, ignores blanks, logical and text
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = true;
        ws.Cell("Z3").Value = "100";
        ws.Cell("Z4").Value = "hello";
        ws.Cell("Z5").Value = 0;
        ws.Cell("Z6").Value = 1;
        await Assert.That(ws.Evaluate("STDEVP(Z1:Z6)")).IsEqualTo(0.5);

        await AssertScalarToNumberConversion("STDEVP", 0.5);
        await AssertAnyErrorIsPropagated("STDEVP");
    }

    [Test]
    public async Task StDevPA()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Example from specification
        await Assert.That((double)ws.Evaluate("STDEVPA(123, 134, 143, 173, 112, 109)")).IsEqualTo(21.66153785).Within(tolerance);

        // Array non-number arguments are ignored
        await Assert.That((double)ws.Evaluate("STDEVPA({0, 1, \"9\", \"Hello\", FALSE, TRUE})")).IsEqualTo(0.5).Within(tolerance);

        // Reference argument ignores blanks, uses numbers, logical and text as zero
        ws.Cell("A1").Value = Blank.Value; // Ignore
        ws.Cell("A2").Value = true; // Include
        ws.Cell("A3").Value = ""; // Consider 0
        ws.Cell("A4").Value = "100"; // Consider 0
        ws.Cell("A5").Value = "hello"; // Consider 0
        ws.Cell("A6").Value = 5;
        ws.Cell("A7").Value = 7;
        await Assert.That((double)ws.Evaluate("STDEVPA(A1:A7)")).IsEqualTo(2.793842436).Within(tolerance);

        // Need at least one sample, otherwise returns error (text in array is ignored)
        await Assert.That(ws.Evaluate("STDEVPA({\"hello\"})")).IsEqualTo(XLError.DivisionByZero);

        await AssertScalarToNumberConversion("STDEVPA", 0.5);
        await AssertAnyErrorIsPropagated("STDEVPA");
    }

    [Test]
    [Arguments("=SUMIF(A1:A10, 1, A1:A10)", 1)]
    [Arguments("=SUMIF(A1:A10, 2.0, A1:A10)", 2)]
    [Arguments("=SUMIF(A1:A10, 3, A1:A10)", 3)]
    [Arguments(@"=SUMIF(A1:A10, ""3"", A1:A10)", 3)]
    [Arguments("=SUMIF(A1:A10, 43831, A1:A10)", 43831)]
    [Arguments("=SUMIF(A1:A10, DATE(2020, 1, 1), A1:A10)", 43831)]
    [Arguments("=SUMIF(A1:A10, TRUE, A1:A10)", 0)]
    public async Task SumIf_MixedData(string formula, double expected)
    {
        // We follow to Excel's convention.
        // Excel treats 1 and TRUE as unequal, but 3 and "3" as equal
        // LibreOffice Calc handles some SUMIF and COUNTIF differently, e.g. it treats 1 and TRUE as equal, but 3 and "3" differently
        var ws = workbook.Worksheet("MixedData");
        await Assert.That(ws.Evaluate(formula)).IsEqualTo(expected);
    }

    [Test]
    public async Task SumIf_specification_examples()
    {
        // Test examples from specification.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 3;
        ws.Cell("B1").Value = 10;
        ws.Cell("C1").Value = 7;
        ws.Cell("D1").Value = 10;

        await Assert.That(ws.Evaluate("SUMIF(A1:D1,\"=10\")")).IsEqualTo(20);
        await Assert.That(ws.Evaluate("SUMIF(A1:D1,\">5\")")).IsEqualTo(27);
        await Assert.That(ws.Evaluate("SUMIF(A1:D1,\"<>10\")")).IsEqualTo(10);

        ws.Cell("A2").Value = "apples";
        ws.Cell("B2").Value = "melons";
        ws.Cell("C2").Value = 10;
        ws.Cell("D2").Value = 15;
        await Assert.That(ws.Evaluate("SUMIF(A2:B2,\"*es\",C2:D2)")).IsEqualTo(10);
    }

    [Test]
    [Arguments("COUNT(G:I,G:G,H:I)", 258d)]
    [Property("Description", "COUNT overlapping columns")]
    [Arguments("COUNT(6:8,6:6,7:8)", 30d)]
    [Property("Description", "COUNT overlapping rows")]
    [Arguments("COUNTBLANK(H:J)", 3145640d)]
    [Property("Description", "COUNTBLANK columns")]
    [Arguments("COUNTBLANK(7:9)", 49128d)]
    [Property("Description", "COUNTBLANK rows")]
    [Arguments("COUNT(1:1048576)", 216d)]
    [Property("Description", "COUNT worksheet")]
    [Arguments("COUNTBLANK(1:1048576)", 17179868831d)]
    [Property("Description", "COUNTBLANK worksheet")]
    [Arguments("SUM(H:J)", 20501.15d)]
    [Property("Description", "SUM columns")]
    [Arguments("SUM(4:5)", 85366.12d)]
    [Property("Description", "SUM rows")]
    [Arguments("SUMIF(G:G,50,H:H)", 24.98d)]
    [Property("Description", "SUMIF columns")]
    [Arguments("SUMIF(G23:G52,\"\",H3:H32)", 53.24d)]
    [Property("Description", "SUMIF ranges")]
    [Arguments("SUMIFS(H:H,G:G,50,I:I,\">900\")", 19.99d)]
    [Property("Description", "SUMIFS columns")]
    public async Task TallySkipsEmptyCells(string formulaA1, double expectedResult)
    {
        using var wb = SetupWorkbook();
        var ws = wb.Worksheets.First();
        //Let's pre-initialize cells we need so they didn't affect the result
        ws.Range("A1:J45").Style.Fill.BackgroundColor = XLColor.Amber;
        ws.Cell("ZZ1000").Value = 1;

        var actualResult = (double)ws.Evaluate(formulaA1);

        await Assert.That(actualResult).IsEqualTo(expectedResult).Within(tolerance);
    }

    [Test]
    public async Task Var()
    {
        var ws = workbook.Worksheets.First();

        // Example from specification
        await Assert.That(ws.Evaluate("VAR(1202,1220,1323,1254,1302)")).IsEqualTo(2683.2);

        // Only non-convertible text in D column, thus less than 2 samples.
        await Assert.That(ws.Evaluate("VAR(D3:D45)")).IsEqualTo(XLError.DivisionByZero);

        // Calculate VAR from numeric values (reference contains only numbers)
        await Assert.That((double)ws.Evaluate("VAR(H3:H45)")).IsEqualTo(2241.560169).Within(tolerance);

        // Ignores text values in the H column and only uses numeric ones, same as reference with only number
        await Assert.That((double)ws.Evaluate("VAR(H:H)")).IsEqualTo(2241.560169).Within(tolerance);
        await Assert.That((double)workbook.Evaluate("VAR(Data!H:H)")).IsEqualTo(2241.560169).Within(tolerance);

        // Need at least two samples, otherwise returns error
        await Assert.That(workbook.Evaluate("VAR({\"hello\"})")).IsEqualTo(XLError.DivisionByZero);
        await Assert.That(workbook.Evaluate("VAR(5)")).IsEqualTo(XLError.DivisionByZero);
        await Assert.That(workbook.Evaluate("VAR(5, 6)")).IsEqualTo(0.5);

        // Array non-number arguments are ignored
        await Assert.That(workbook.Evaluate("VAR({0, 1, \"Hello\", FALSE, TRUE})")).IsEqualTo(0.5);

        // Reference argument only uses number, ignores blanks, logical and text
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = true;
        ws.Cell("Z3").Value = "100";
        ws.Cell("Z4").Value = "hello";
        ws.Cell("Z5").Value = 0;
        ws.Cell("Z6").Value = 1;
        await Assert.That(ws.Evaluate("VAR(Z1:Z6)")).IsEqualTo(0.5);

        await AssertScalarToNumberConversion("VAR", 0.5);
        await AssertAnyErrorIsPropagated("VAR");
    }

    [Test]
    public async Task VarA()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Example from specification
        await Assert.That(ws.Evaluate("VARA(1202, 1220, 1323, 1254, 1302)")).IsEqualTo(2683.2);

        // Array non-number arguments are ignored
        await Assert.That(ws.Evaluate("VARA({5, 7, \"9\", \"Hello\", FALSE, TRUE})")).IsEqualTo(2);

        // Reference argument ignores blanks, uses numbers, logical and text as zero
        ws.Cell("A1").Value = Blank.Value; // Ignore
        ws.Cell("A2").Value = true; // Include
        ws.Cell("A3").Value = ""; // Consider 0
        ws.Cell("A4").Value = "100"; // Consider 0
        ws.Cell("A5").Value = "hello"; // Consider 0
        ws.Cell("A6").Value = 5;
        ws.Cell("A7").Value = 7;
        await Assert.That((double)ws.Evaluate("VARA(A1:A7)")).IsEqualTo(9.366666667).Within(tolerance);

        // Need at least one sample, otherwise returns error (text in array is ignored)
        await Assert.That(ws.Evaluate("VARA({\"hello\"})")).IsEqualTo(XLError.DivisionByZero);

        await AssertScalarToNumberConversion("VARA", 0.5);
        await AssertAnyErrorIsPropagated("VARA");
    }

    [Test]
    public async Task VarP()
    {
        var ws = workbook.Worksheets.First();

        // Example from specification
        await Assert.That((double)ws.Evaluate("VARP(1202,1220,1323,1254,1302)")).IsEqualTo(2146.56).Within(tolerance);

        // Only non-convertible text in D column, thus less than 1 sample.
        await Assert.That(ws.Evaluate("VARP(D3:D45)")).IsEqualTo(XLError.DivisionByZero);

        // Calculate VARP from numeric values (reference contains only numbers)
        await Assert.That((double)ws.Evaluate("VARP(H3:H45)")).IsEqualTo(2189.430863).Within(tolerance);

        // Ignores text values in the H column and only uses numeric ones, same as reference with only number
        await Assert.That((double)ws.Evaluate("VARP(H:H)")).IsEqualTo(2189.430863).Within(tolerance);
        await Assert.That((double)workbook.Evaluate("VARP(Data!H:H)")).IsEqualTo(2189.430863).Within(tolerance);

        // Need at least one sample, otherwise returns error
        await Assert.That(workbook.Evaluate("VARP({\"hello\"})")).IsEqualTo(XLError.DivisionByZero);
        await Assert.That(workbook.Evaluate("VARP(5)")).IsEqualTo(0);

        // Array non-number arguments are ignored
        await Assert.That(workbook.Evaluate("VARP({0, 1, \"Hello\", FALSE, TRUE})")).IsEqualTo(0.25);

        // Reference argument only uses number, ignores blanks, logical and text
        ws.Cell("Z1").Value = Blank.Value;
        ws.Cell("Z2").Value = true;
        ws.Cell("Z3").Value = "100";
        ws.Cell("Z4").Value = "hello";
        ws.Cell("Z5").Value = 0;
        ws.Cell("Z6").Value = 1;
        await Assert.That(ws.Evaluate("VARP(Z1:Z6)")).IsEqualTo(0.25);

        await AssertScalarToNumberConversion("VARP", 0.25);
        await AssertAnyErrorIsPropagated("VARP");
    }

    [Test]
    public async Task VarPA()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Example from specification
        await Assert.That(ws.Evaluate("VARPA(1202, 1220, 1323, 1254, 1302)")).IsEqualTo(2146.56);

        // Array non-number arguments are ignored
        await Assert.That(ws.Evaluate("VARPA({5, 7, \"9\", \"Hello\", FALSE, TRUE})")).IsEqualTo(1);

        // Reference argument ignores blanks, uses numbers, logical and text as zero
        ws.Cell("A1").Value = Blank.Value; // Ignore
        ws.Cell("A2").Value = true; // Include
        ws.Cell("A3").Value = ""; // Consider 0
        ws.Cell("A4").Value = "100"; // Consider 0
        ws.Cell("A5").Value = "hello"; // Consider 0
        ws.Cell("A6").Value = 5;
        ws.Cell("A7").Value = 7;
        await Assert.That((double)ws.Evaluate("VARPA(A1:A7)")).IsEqualTo(7.805555556).Within(tolerance);

        // Need at least one sample, otherwise returns error (text in array is ignored)
        await Assert.That(ws.Evaluate("VARPA({\"hello\"})")).IsEqualTo(XLError.DivisionByZero);

        await AssertScalarToNumberConversion("VARPA", 0.25);
        await AssertAnyErrorIsPropagated("VARPA");
    }

    [Test]
    public async Task Large()
    {
        var ws = workbook.Worksheet("Data");
        var value = ws.Evaluate("LARGE(G1:G45, 1)");
        await Assert.That(value).IsEqualTo(96);

        value = ws.Evaluate("LARGE(G1:G45, 7)");
        await Assert.That(value).IsEqualTo(87);

        value = ws.Evaluate("LARGE(G1:G45, 0)");
        await Assert.That(value).IsEqualTo(XLError.NumberInvalid);

        value = ws.Evaluate("LARGE(G1:G45, -1)");
        await Assert.That(value).IsEqualTo(XLError.NumberInvalid);

        value = ws.Evaluate("LARGE(G1:G45,\"test\")");
        await Assert.That(value).IsEqualTo(XLError.IncompatibleValue);

        value = ws.Evaluate("LARGE(C:C,7)");
        await Assert.That(value).IsEqualTo(42623);

        value = ws.Evaluate("LARGE(D:D,7)");
        await Assert.That(value).IsEqualTo(XLError.NumberInvalid);

        ws = workbook.Worksheet("MixedData");

        value = ws.Evaluate("LARGE(A1:A7,6)");
        await Assert.That(value).IsEqualTo(XLError.NumberInvalid);

        // Ignores non-numbers.
        value = ws.Evaluate("LARGE(A1:A7,5)");
        await Assert.That(value).IsEqualTo(1);

        // Accepts non-area references.
        value = ws.Evaluate("LARGE((A1:A2,A4:A6),2)");
        await Assert.That(value).IsEqualTo(3);

        // Errors are returned.
        value = ws.Evaluate("LARGE({ 1, 2, #N/A }, 1)");
        await Assert.That(value).IsEqualTo(XLError.NoValueAvailable);

        // Uses ceiling logic for number (1.1 -> 2) + can use arrays.
        value = ws.Evaluate("LARGE({ 1, 2 }, 1.1)");
        await Assert.That(value).IsEqualTo(1);

        // If a scalar number-like value supplied, it is converted to number.
        value = ws.Evaluate("LARGE(\"1 1/2\", 1)");
        await Assert.That(value).IsEqualTo(1.5);

        // When the scalar can't be converted, return conversion error.
        value = ws.Evaluate("LARGE(\"test\", 1)");
        await Assert.That(value).IsEqualTo(XLError.IncompatibleValue);
    }

    private static XLWorkbook SetupWorkbook()
    {
        var wb = new XLWorkbook();
        var ws1 = wb.AddWorksheet("Data");
        var data = new object[]
        {
            new {Id=1, OrderDate = DateTime.Parse("2015-01-06"), Region = "East", Rep = "Jones", Item = "Pencil", Units = 95, UnitCost = 1.99, Total = 189.05 },
            new {Id=2, OrderDate = DateTime.Parse("2015-01-23"), Region = "Central", Rep = "Kivell", Item = "Binder", Units = 50, UnitCost = 19.99, Total = 999.5},
            new {Id=3, OrderDate = DateTime.Parse("2015-02-09"), Region = "Central", Rep = "Jardine", Item = "Pencil", Units = 36, UnitCost = 4.99, Total = 179.64},
            new {Id=4, OrderDate = DateTime.Parse("2015-02-26"), Region = "Central", Rep = "Gill", Item = "Pen", Units = 27, UnitCost = 19.99, Total = 539.73},
            new {Id=5, OrderDate = DateTime.Parse("2015-03-15"), Region = "West", Rep = "Sorvino", Item = "Pencil", Units = 56, UnitCost = 2.99, Total = 167.44},
            new {Id=6, OrderDate = DateTime.Parse("2015-04-01"), Region = "East", Rep = "Jones", Item = "Binder", Units = 60, UnitCost = 4.99, Total = 299.4},
            new {Id=7, OrderDate = DateTime.Parse("2015-04-18"), Region = "Central", Rep = "Andrews", Item = "Pencil", Units = 75, UnitCost = 1.99, Total = 149.25},
            new {Id=8, OrderDate = DateTime.Parse("2015-05-05"), Region = "Central", Rep = "Jardine", Item = "Pencil", Units = 90, UnitCost = 4.99, Total = 449.1},
            new {Id=9, OrderDate = DateTime.Parse("2015-05-22"), Region = "West", Rep = "Thompson", Item = "Pencil", Units = 32, UnitCost = 1.99, Total = 63.68},
            new {Id=10, OrderDate = DateTime.Parse("2015-06-08"), Region = "East", Rep = "Jones", Item = "Binder", Units = 60, UnitCost = 8.99, Total = 539.4},
            new {Id=11, OrderDate = DateTime.Parse("2015-06-25"), Region = "Central", Rep = "Morgan", Item = "Pencil", Units = 90, UnitCost = 4.99, Total = 449.1},
            new {Id=12, OrderDate = DateTime.Parse("2015-07-12"), Region = "East", Rep = "Howard", Item = "Binder", Units = 29, UnitCost = 1.99, Total = 57.71},
            new {Id=13, OrderDate = DateTime.Parse("2015-07-29"), Region = "East", Rep = "Parent", Item = "Binder", Units = 81, UnitCost = 19.99, Total = 1619.19},
            new {Id=14, OrderDate = DateTime.Parse("2015-08-15"), Region = "East", Rep = "Jones", Item = "Pencil", Units = 35, UnitCost = 4.99, Total = 174.65},
            new {Id=15, OrderDate = DateTime.Parse("2015-09-01"), Region = "Central", Rep = "Smith", Item = "Desk", Units = 2, UnitCost = 125, Total = 250},
            new {Id=16, OrderDate = DateTime.Parse("2015-09-18"), Region = "East", Rep = "Jones", Item = "Pen Set", Units = 16, UnitCost = 15.99, Total = 255.84},
            new {Id=17, OrderDate = DateTime.Parse("2015-10-05"), Region = "Central", Rep = "Morgan", Item = "Binder", Units = 28, UnitCost = 8.99, Total = 251.72},
            new {Id=18, OrderDate = DateTime.Parse("2015-10-22"), Region = "East", Rep = "Jones", Item = "Pen", Units = 64, UnitCost = 8.99, Total = 575.36},
            new {Id=19, OrderDate = DateTime.Parse("2015-11-08"), Region = "East", Rep = "Parent", Item = "Pen", Units = 15, UnitCost = 19.99, Total = 299.85},
            new {Id=20, OrderDate = DateTime.Parse("2015-11-25"), Region = "Central", Rep = "Kivell", Item = "Pen Set", Units = 96, UnitCost = 4.99, Total = 479.04},
            new {Id=21, OrderDate = DateTime.Parse("2015-12-12"), Region = "Central", Rep = "Smith", Item = "Pencil", Units = 67, UnitCost = 1.29, Total = 86.43},
            new {Id=22, OrderDate = DateTime.Parse("2015-12-29"), Region = "East", Rep = "Parent", Item = "Pen Set", Units = 74, UnitCost = 15.99, Total = 1183.26},
            new {Id=23, OrderDate = DateTime.Parse("2016-01-15"), Region = "Central", Rep = "Gill", Item = "Binder", Units = 46, UnitCost = 8.99, Total = 413.54},
            new {Id=24, OrderDate = DateTime.Parse("2016-02-01"), Region = "Central", Rep = "Smith", Item = "Binder", Units = 87, UnitCost = 15, Total = 1305},
            new {Id=25, OrderDate = DateTime.Parse("2016-02-18"), Region = "East", Rep = "Jones", Item = "Binder", Units = 4, UnitCost = 4.99, Total = 19.96},
            new {Id=26, OrderDate = DateTime.Parse("2016-03-07"), Region = "West", Rep = "Sorvino", Item = "Binder", Units = 7, UnitCost = 19.99, Total = 139.93},
            new {Id=27, OrderDate = DateTime.Parse("2016-03-24"), Region = "Central", Rep = "Jardine", Item = "Pen Set", Units = 50, UnitCost = 4.99, Total = 249.5},
            new {Id=28, OrderDate = DateTime.Parse("2016-04-10"), Region = "Central", Rep = "Andrews", Item = "Pencil", Units = 66, UnitCost = 1.99, Total = 131.34},
            new {Id=29, OrderDate = DateTime.Parse("2016-04-27"), Region = "East", Rep = "Howard", Item = "Pen", Units = 96, UnitCost = 4.99, Total = 479.04},
            new {Id=30, OrderDate = DateTime.Parse("2016-05-14"), Region = "Central", Rep = "Gill", Item = "Pencil", Units = 53, UnitCost = 1.29, Total = 68.37},
            new {Id=31, OrderDate = DateTime.Parse("2016-05-31"), Region = "Central", Rep = "Gill", Item = "Binder", Units = 80, UnitCost = 8.99, Total = 719.2},
            new {Id=32, OrderDate = DateTime.Parse("2016-06-17"), Region = "Central", Rep = "Kivell", Item = "Desk", Units = 5, UnitCost = 125, Total = 625},
            new {Id=33, OrderDate = DateTime.Parse("2016-07-04"), Region = "East", Rep = "Jones", Item = "Pen Set", Units = 62, UnitCost = 4.99, Total = 309.38},
            new {Id=34, OrderDate = DateTime.Parse("2016-07-21"), Region = "Central", Rep = "Morgan", Item = "Pen Set", Units = 55, UnitCost = 12.49, Total = 686.95},
            new {Id=35, OrderDate = DateTime.Parse("2016-08-07"), Region = "Central", Rep = "Kivell", Item = "Pen Set", Units = 42, UnitCost = 23.95, Total = 1005.9},
            new {Id=36, OrderDate = DateTime.Parse("2016-08-24"), Region = "West", Rep = "Sorvino", Item = "Desk", Units = 3, UnitCost = 275, Total = 825},
            new {Id=37, OrderDate = DateTime.Parse("2016-09-10"), Region = "Central", Rep = "Gill", Item = "Pencil", Units = 7, UnitCost = 1.29, Total = 9.03},
            new {Id=38, OrderDate = DateTime.Parse("2016-09-27"), Region = "West", Rep = "Sorvino", Item = "Pen", Units = 76, UnitCost = 1.99, Total = 151.24},
            new {Id=39, OrderDate = DateTime.Parse("2016-10-14"), Region = "West", Rep = "Thompson", Item = "Binder", Units = 57, UnitCost = 19.99, Total = 1139.43},
            new {Id=40, OrderDate = DateTime.Parse("2016-10-31"), Region = "Central", Rep = "Andrews", Item = "Pencil", Units = 14, UnitCost = 1.29, Total = 18.06},
            new {Id=41, OrderDate = DateTime.Parse("2016-11-17"), Region = "Central", Rep = "Jardine", Item = "Binder", Units = 11, UnitCost = 4.99, Total = 54.89},
            new {Id=42, OrderDate = DateTime.Parse("2016-12-04"), Region = "Central", Rep = "Jardine", Item = "Binder", Units = 94, UnitCost = 19.99, Total = 1879.06},
            new {Id=43, OrderDate = DateTime.Parse("2016-12-21"), Region = "Central", Rep = "Andrews", Item = "Binder", Units = 28, UnitCost = 4.99, Total = 139.72}
        };

        ws1.FirstCell()
            .CellBelow()
            .CellRight()
            .InsertTable(data, "Table1");

        var ws2 = wb.AddWorksheet("MixedData");
        ws2.FirstCell().InsertData(new object[] { 1, 2.0, "3", 3, new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified), true, new TimeSpan(10, 5, 30, 10) });

        return wb;
    }

    private static async Task AssertScalarToNumberConversion(string functionName, double result)
    {
        // Scalar blank is converted to 0
        await Assert.That((double)XLWorkbook.EvaluateExpr($"{functionName}(IF(TRUE,), 1)")).IsEqualTo(result).Within(tolerance);

        // Scalar logical is converted to a number
        await Assert.That((double)XLWorkbook.EvaluateExpr($"{functionName}(FALSE, TRUE)")).IsEqualTo(result).Within(tolerance);
        await Assert.That((double)XLWorkbook.EvaluateExpr($"{functionName}(0, TRUE)")).IsEqualTo(result).Within(tolerance);
        await Assert.That((double)XLWorkbook.EvaluateExpr($"{functionName}(FALSE, 1)")).IsEqualTo(result).Within(tolerance);

        // Scalar text is converted to a number
        await Assert.That((double)XLWorkbook.EvaluateExpr($"{functionName}(\"0\", \"1\")")).IsEqualTo(result).Within(tolerance);
        await Assert.That((double)XLWorkbook.EvaluateExpr($"{functionName}(\"1\", \"0 0/2\")")).IsEqualTo(result).Within(tolerance);

        // Scalar text that is not convertible returns error
        await Assert.That(XLWorkbook.EvaluateExpr($"{functionName}(5, \"Hello\")")).IsEqualTo(XLError.IncompatibleValue);
    }

    /// <summary>
    /// Assert that a function propagates any error, whether from scalar, array or reference argument.
    /// </summary>
    /// <param name="functionName">Name of a function that accepts any value as argument.</param>
    private static async Task AssertAnyErrorIsPropagated(string functionName)
    {
        // Scalar error is propagated
        await Assert.That(XLWorkbook.EvaluateExpr($"{functionName}(1, #NULL!)")).IsEqualTo(XLError.NullValue);

        // Array error is propagated
        await Assert.That(XLWorkbook.EvaluateExpr($"{functionName}({{1, #NULL!}})")).IsEqualTo(XLError.NullValue);

        // Reference error is propagated
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B1").Value = XLError.NoValueAvailable;
        ws.Cell("B2").Value = 1;
        await Assert.That(ws.Evaluate($"{functionName}(B1)")).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(ws.Evaluate($"{functionName}(B1:B2)")).IsEqualTo(XLError.NoValueAvailable);
    }
}
