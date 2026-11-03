using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Behaviour the Excel conformance corpus (spec 59) showed XLibur getting wrong. Each expected value
/// is what Excel returned when the corpus was recorded.
/// </summary>
[SetCulture("en-US")]
public class ExcelConformanceFixesTests
{
    [Test]
    [Arguments("SUM(9E307,9E307)")]
    [Arguments("PRODUCT(1E200,1E200)")]
    [Arguments("9E307*10")]
    [Arguments("DEGREES(9E307)")]
    [Arguments("AVERAGE(9E307,9E307)")]
    public async Task Overflow_to_infinity_is_a_number_error(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task An_overflowed_value_is_not_a_number_and_reaches_the_iferror_fallback()
    {
        await Assert.That((bool)XLWorkbook.EvaluateExpr("ISNUMBER(9E307*10)")).IsFalse();
        await Assert.That(XLWorkbook.EvaluateExpr("IFERROR(9E307*10,\"x\")")).IsEqualTo("x");
    }

    [Test]
    public async Task Radians_and_median_do_not_overflow_in_an_intermediate_step()
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr("RADIANS(9E307)")).IsEqualTo(1.5707963267948965E+306);
        await Assert.That((double)XLWorkbook.EvaluateExpr("MEDIAN(9E307,9E307)")).IsEqualTo(9E307);
    }

    [Test]
    [Arguments("ROUNDUP(2.345,5)", 2.345)]
    [Arguments("ROUNDUP(2.345,\"5\")", 2.345)]
    [Arguments("ROUNDUP(2.345,2)", 2.35)]
    [Arguments("ROUNDUP(-2.341,2)", -2.35)]
    [Arguments("ROUNDUP(1234.5678,-2)", 1300)]
    [Arguments("ROUNDUP(9E307,2)", 9E307)]
    [Arguments("ROUNDUP(2.345,1000000)", 2.345)]
    [Arguments("ROUNDUP(1E-300,2)", 0.01)]
    [Arguments("ROUNDDOWN(0.29,2)", 0.29)]
    [Arguments("ROUNDDOWN(1234.5678,-2)", 1200)]
    [Arguments("ROUNDDOWN(9E307,2)", 9E307)]
    [Arguments("TRUNC(8.9,2.5)", 8.9)]
    [Arguments("TRUNC(8.9)", 8)]
    [Arguments("TRUNC(-8.97,1)", -8.9)]
    [Arguments("TRUNC(8.9,1000000)", 8.9)]
    public async Task Rounding_towards_a_direction_works_on_the_decimal_form_of_the_value(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("BIN2DEC(\"\")", 0)]
    [Arguments("HEX2DEC(\"\")", 0)]
    [Arguments("OCT2DEC(\"\")", 0)]
    public async Task An_empty_string_converts_to_zero(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("BIN2DEC(\" \")")]
    [Arguments("BIN2DEC(\"-1.5\")")]
    [Arguments("HEX2DEC(\"-1.5\")")]
    [Arguments("OCT2DEC(\"-1.5\")")]
    [Arguments("HEX2BIN(\"-1.5\")")]
    [Arguments("BIN2HEX(\"-1.5\")")]
    public async Task A_minus_sign_or_only_spaces_in_a_base_conversion_is_a_number_error(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }
}
