using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

public class FinancialTests
{
    [Test]
    [Arguments("FV(0.06/12,10,-200,-500,1)", 2581.4033740601362)]
    [Arguments("FV(0.12/12,12,-1000)", 12682.503013196976)]
    [Arguments("FV(0.11/12,35,-2000,,1)", 82846.24637190059)]
    [Arguments("FV(0.06/12,12,-100,-1000,1)", 2301.4018303409139)]
    public async Task Fv_ReferenceExamplesFromExcelDocumentations(string formula, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr(formula);
        await Assert.That(actual).IsEqualTo(expectedResult).Within(XLHelper.Epsilon);
    }

    [Test]
    [Arguments("FV(0,1,1000)", -1000)] // Zero interest rate
    [Arguments("FV(0,5,10000,5000)", -55000.00)] // Zero interest rate with present value
    [Arguments("FV(-0.4,2,1000)", -1600.00)] // Negative interest rate
    [Arguments("FV(0.01,0.5,1000)", -498.75621120889502)] // Non-integer period
    [Arguments("FV(0.1,-2,1000)", 1735.5371900826453)] // Negative periods
    [Arguments("FV(0.1,2,0,4)", -4.84)] // No PMT, but present value
    [Arguments("FV(0,2,-1000)", 2000.00)] // Negative PMT - money is paid to us
    [Arguments("FV(0.000001,1000,1000)", -1000499.6661261424)] // Small number and high number of periods, check for stability
    public async Task Fv_EdgeCases(string formula, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr(formula);
        await Assert.That(actual).IsEqualTo(expectedResult).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task Fv_DefaultFutureValueIsZero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("FV(0.1,2,1000,0)")).IsEqualTo(XLWorkbook.EvaluateExpr("FV(0.1,2,1000)"));
    }

    [Test]
    public async Task Fv_DefaultTypeIsZero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("FV(0.1,5,1000,0,0)")).IsEqualTo(XLWorkbook.EvaluateExpr("FV(0.1,5,1000)"));
    }

    [Test]
    public async Task Fv_ZeroPeriodsReturnsPresentValue()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("FV(0.1,0,1000, 100)")).IsEqualTo(-100);
    }

    [Test]
    [Arguments("IPMT(0.1/12,1,3*12,8000)", -66.666666666666686)]
    [Arguments("IPMT(0.1,3,3,8000)", -292.4471299093658)]
    public async Task Ipmt_ReferenceExamplesFromExcelDocumentations(string formula, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr(formula);
        await Assert.That(actual).IsEqualTo(expectedResult).Within(XLHelper.Epsilon);
    }

    [Test]
    [Arguments("IPMT(0,1,1,1000)", 0)] // Zero interest rate
    [Arguments("IPMT(0,1,5,10000,5000)", 0)] // Zero interest rate with future value
    [Arguments("IPMT(-0.4,1,2,1000)", 400.00)] // Negative interest rate
    [Arguments("IPMT(0.01,1,0.5,1000)", -10.00)] // Non-integer period
    [Arguments("IPMT(0.01,1,1.4,1000)", -10.00)] // Different non-integer period
    [Arguments("IPMT(0.1,1,2,0,4)", 0)] // No principal, but future value
    [Arguments("IPMT(0.1,1,2,-1000)", 100)] // Negative principal - money is paid to us
    [Arguments("IPMT(0.000001,1,1000,1000)", -0.001)] // Small number and high number of periods, check for stability
    public async Task Ipmt_EdgeCases(string formula, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr(formula);
        await Assert.That(actual).IsEqualTo(expectedResult).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task Ipmt_DefaultFutureValueIsZero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("IPMT(0.1,1,2,1000,0)")).IsEqualTo(XLWorkbook.EvaluateExpr("IPMT(0.1,1,2,1000)"));
    }

    [Test]
    public async Task Ipmt_DefaultTypeIsZero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("IPMT(0.1,1,5,1000,0,0)")).IsEqualTo(XLWorkbook.EvaluateExpr("IPMT(0.1,1,5,1000)"));
    }

    [Test]
    public async Task Ipmt_ZeroOrNegativePeriodsReturnsNumError()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("IPMT(0.1,1,0,1000)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("IPMT(0.1,1,-1,1000)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(-1)]
    [Arguments(-1.5)]
    [Arguments(-100)]
    public async Task Ipmt_RateLessOrEqualMinusOneReturnsNumError(double rate)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"IPMT({rate},2,3,1000,10000,1)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Ipmt_PeriodOutOfRangeReturnsNumError()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("IPMT(0.1,0,1,1000)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("IPMT(0.1,2,1,1000)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments("PMT(0.08/12,10,10000)", -1037.03208935915)]
    [Arguments("PMT(0.08/12,10,10000,0,1)", -1030.16432717797)]
    public async Task Pmt_ReferenceExamplesFromExcelDocumentations(string formula, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr(formula);
        await Assert.That(actual).IsEqualTo(expectedResult).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task Pmt_PaymentsMustPayForPrincipalAndFutureValue()
    {
        var actual = (double)XLWorkbook.EvaluateExpr("PMT(0,2,5000,10000)");
        await Assert.That(actual).IsEqualTo(-7500);
    }

    [Test]
    [Arguments("PMT(0,1,1000)", -1000)] // Zero interest rate
    [Arguments("PMT(0,5,10000,5000)", -3000)] // Zero interest rate for 5 years, (10k principal, pay all and have 5k in bank at the end = payment is 3k/year)
    [Arguments("PMT(-0.4,2,1000)", -225)] // Negative interest rate
    [Arguments("PMT(0.01,0.5,1000)", -2014.98756211209)] // Non-integer period
    [Arguments("PMT(0.1,-2,1000)", 476.19047619048)] // Negative periods
    [Arguments("PMT(0.1,2,0,4)", -1.90476190476)] // No principal, but future value
    [Arguments("PMT(0,2,-1000)", 500)] // Negative principal - money is paid to us
    [Arguments("PMT(0.000001,1000,1000)", -1.00050058333321)] // Small number and high number of periods, check for stability
    public async Task Pmt_EdgeCases(string formula, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr(formula);
        await Assert.That(actual).IsEqualTo(expectedResult).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task Pmt_TypeConvertsNumberToOneOrZero()
    {
        // Spec says "if type is any number other than 0 or 1, #NUM! is returned.", but Excel accepts any number as type
        var formulaFormat = "PMT(0.1,2,1000,500,{0})";
        var zeroType = (double)XLWorkbook.EvaluateExpr(string.Format(formulaFormat, "0"));
        var oneType = (double)XLWorkbook.EvaluateExpr(string.Format(formulaFormat, "1"));
        var nonZeroType = (double)XLWorkbook.EvaluateExpr(string.Format(formulaFormat, "0.000001"));

        await Assert.That(oneType).IsNotEqualTo(zeroType);
        await Assert.That(nonZeroType).IsEqualTo(oneType);
    }

    [Test]
    public async Task Pmt_DefaultFutureValueIsZero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("PMT(0.1,2,1000,0)")).IsEqualTo(XLWorkbook.EvaluateExpr("PMT(0.1,2,1000)"));
    }

    [Test]
    public async Task Pmt_DefaultTypeIsZero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("PMT(0.1,5,1000,0,0)")).IsEqualTo(XLWorkbook.EvaluateExpr("PMT(0.1,5,1000)"));
    }

    [Test]
    public async Task Pmt_ZeroPeriodsReturnsNumError()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("PMT(0.1,0,1000)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(-1)]
    [Arguments(-1.5)]
    [Arguments(-100)]
    public async Task Pmt_RateLessOrEqualMinusOneReturnsNumError(double rate)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"PMT({rate},1,1000,5000,1)")).IsEqualTo(XLError.NumberInvalid);
    }
}
