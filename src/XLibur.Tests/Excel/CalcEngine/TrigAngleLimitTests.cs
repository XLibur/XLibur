using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// Excel's SIN, COS and TAN give <c>#NUM!</c> for an angle of 2^27 radians (134,217,728) or more, in
/// either direction, and a value for anything below it. Measured in Excel over COM.
/// </summary>
[SetCulture("en-US")]
public class TrigAngleLimitTests
{
    [Test]
    [Arguments("SIN", "134217728")]
    [Arguments("SIN", "-134217728")]
    [Arguments("SIN", "1E10")]
    [Arguments("SIN", "1E15")]
    [Arguments("SIN", "9E307")]
    [Arguments("COS", "134217728")]
    [Arguments("COS", "-134217728")]
    [Arguments("COS", "1E10")]
    [Arguments("COS", "1E15")]
    [Arguments("COS", "9E307")]
    [Arguments("TAN", "134217728")]
    [Arguments("TAN", "1E15")]
    public async Task An_angle_of_two_to_the_27th_or_more_is_a_number_error(string function, string angle)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"{function}({angle})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments("SIN", -0.9559933287726706)]
    [Arguments("COS", -0.2933884035577216)]
    [Arguments("TAN", 3.258456425611884)]
    public async Task The_largest_angle_below_the_limit_still_has_a_result(string function, double expected)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"{function}(134217727)");

        await Assert.That(actual).IsEqualTo(expected).Within(1e-9);
    }
}
