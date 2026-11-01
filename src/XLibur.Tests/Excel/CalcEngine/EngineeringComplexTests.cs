using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// COMPLEX and the IM* family. Excel has no complex value type — a complex number is the text
/// "3+4i" — so these tests compare the rendered text. Expected values are the worked examples from
/// Microsoft's per-function documentation unless the comment says otherwise.
/// </summary>
[SetCulture("en-US")]
public class EngineeringComplexTests
{
    [Test]
    [Arguments("COMPLEX(3, 4)", "3+4i")]
    [Arguments("COMPLEX(3, 4, \"j\")", "3+4j")]
    [Arguments("COMPLEX(0, 1)", "i")] // A unit coefficient is written bare.
    [Arguments("COMPLEX(1, 0)", "1")] // No imaginary part at all.
    [Arguments("COMPLEX(0, 0)", "0")]
    [Arguments("COMPLEX(0, -1)", "-i")]
    [Arguments("COMPLEX(3, -4)", "3-4i")]
    [Arguments("COMPLEX(-3, 4)", "-3+4i")]
    public async Task Complex_BuildsTheTextForm(string formula, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    [Arguments("COMPLEX(3, 4, \"k\")")] // Only i and j name the imaginary unit.
    [Arguments("COMPLEX(3, 4, \"I\")")] // The suffix is case sensitive.
    [Arguments("COMPLEX(3, 4, \"\")")]
    public async Task Complex_UnknownSuffixReturnsIncompatibleValue(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("IMABS(\"5+12i\")", 13d)]
    [Arguments("IMABS(\"3+4i\")", 5d)]
    [Arguments("IMABS(\"-3\")", 3d)] // A real number is a complex number with no imaginary part.
    [Arguments("IMREAL(\"6-9i\")", 6d)]
    [Arguments("IMREAL(\"i\")", 0d)]
    [Arguments("IMAGINARY(\"3+4i\")", 4d)]
    [Arguments("IMAGINARY(\"0-j\")", -1d)]
    [Arguments("IMAGINARY(4)", 0d)]
    [Arguments("IMAGINARY(\"i\")", 1d)]
    [Arguments("IMARGUMENT(\"3+4i\")", 0.927295218001612d)] // ATAN2(3, 4).
    public async Task ComplexComponents_ReferenceExamplesFromExcelDocumentation(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-12);
    }

    [Test]
    public async Task ImArgument_OfZeroIsDivisionByZero()
    {
        // The angle of the origin is not defined.
        await Assert.That(XLWorkbook.EvaluateExpr("IMARGUMENT(\"0\")")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    [Arguments("IMCONJUGATE(\"3+4i\")", "3-4i")]
    [Arguments("IMCONJUGATE(\"3-4j\")", "3+4j")] // The j spelling is echoed back.
    [Arguments("IMSUM(\"3+4i\", \"5-3i\")", "8+i")]
    [Arguments("IMSUB(\"13+4i\", \"5+3i\")", "8+i")]
    [Arguments("IMPRODUCT(\"3+4i\", \"5-3i\")", "27+11i")]
    [Arguments("IMDIV(\"-238+240i\", \"10+24i\")", "5+12i")]
    [Arguments("IMPOWER(\"2+3i\", 3)", "-46+9.00000000000001i")]
    public async Task ComplexArithmetic_ReferenceExamplesFromExcelDocumentation(string formula, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    public async Task ImSum_AndImProduct_AcceptASingleArgumentAndManyArguments()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("IMSUM(\"3+4i\")")).IsEqualTo("3+4i");
        await Assert.That(XLWorkbook.EvaluateExpr("IMSUM(\"1+i\", \"1+i\", \"1+i\")")).IsEqualTo("3+3i");
        await Assert.That(XLWorkbook.EvaluateExpr("IMPRODUCT(\"1+i\", \"1+i\", \"1+i\", \"1+i\")")).IsEqualTo("-4");
    }

    [Test]
    public async Task ComplexFunctions_RefuseToMixTheTwoImaginaryUnits()
    {
        // Excel will not add a number written with i to one written with j.
        await Assert.That(XLWorkbook.EvaluateExpr("IMSUM(\"1+i\", \"1+j\")")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr("IMSUB(\"1+i\", \"1+j\")")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr("IMDIV(\"1+i\", \"1+j\")")).IsEqualTo(XLError.IncompatibleValue);

        // A number with no imaginary part carries no spelling, so it mixes with either.
        await Assert.That(XLWorkbook.EvaluateExpr("IMSUM(\"1\", \"1+j\")")).IsEqualTo("2+j");
    }

    [Test]
    [Arguments("IMEXP(\"1+i\")", "1.46869393991589+2.28735528717884i")]
    [Arguments("IMLN(\"3+4i\")", "1.6094379124341+0.927295218001612i")]
    [Arguments("IMLOG10(\"3+4i\")", "0.698970004336019+0.402719196273373i")]
    [Arguments("IMLOG2(\"3+4i\")", "2.32192809488736+1.33780421245098i")]
    [Arguments("IMSQRT(\"1+i\")", "1.09868411346781+0.455089860562227i")]
    [Arguments("IMSQRT(\"3+4i\")", "2+i")] // The exact answer, once rounded to 15 digits.
    public async Task ComplexTranscendentals_ReferenceExamplesFromExcelDocumentation(string formula, string expected)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected);
    }

    [Test]
    // Compared component by component rather than as text: the last of the 15 digits Excel prints
    // is at the mercy of how the identity is grouped, and a one-ulp difference there says nothing
    // about correctness. Each expectation is sin/cos/sinh/cosh of the argument applied through the
    // standard addition formulae, e.g. cosh(4+3i) = cosh4·cos3 + i·sinh4·sin3.
    [Arguments("IMSIN(\"3+4i\")", 3.85373803791938d, -27.0168132580039d)]
    [Arguments("IMCOS(\"1+i\")", 0.833730025131149d, -0.988897705762865d)]
    [Arguments("IMSINH(\"4+3i\")", -27.0168132580039d, 3.85373803791938d)]
    [Arguments("IMCOSH(\"4+3i\")", -27.0349456030742d, 3.85115333481178d)]
    [Arguments("IMTAN(\"4+3i\")", 0.00490825806749606d, 1.00070953606723d)]
    [Arguments("IMCOT(\"4+3i\")", 0.00490118239430447d, -0.999266927805902d)]
    [Arguments("IMSEC(\"4+3i\")", -0.0652940278579471d, -0.0752249603027732d)]
    [Arguments("IMCSC(\"4+3i\")", -0.0754898329158637d, 0.0648774713706355d)]
    [Arguments("IMSECH(\"4+3i\")", -0.0362534969158689d, -0.00516434460775318d)]
    [Arguments("IMCSCH(\"4+3i\")", -0.0362758896286264d, -0.00517447318401943d)]
    public async Task ComplexTrigonometry_MatchesTheAdditionFormulae(string formula, double expectedReal, double expectedImaginary)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"IMREAL({formula})")).IsEqualTo(expectedReal).Within(1e-10);
        await Assert.That((double)XLWorkbook.EvaluateExpr($"IMAGINARY({formula})")).IsEqualTo(expectedImaginary).Within(1e-10);
    }

    [Test]
    public async Task ComplexTrigonometry_ReciprocalsInvertTheirPrimitives()
    {
        // sec = 1/cos, csc = 1/sin, sech = 1/cosh, csch = 1/sinh — so multiplying the two back
        // together has to give 1.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").FormulaA1 = "IMPRODUCT(IMSEC(\"4+3i\"), IMCOS(\"4+3i\"))";
        ws.Cell("A2").FormulaA1 = "IMPRODUCT(IMCSC(\"4+3i\"), IMSIN(\"4+3i\"))";
        ws.Cell("A3").FormulaA1 = "IMPRODUCT(IMSECH(\"4+3i\"), IMCOSH(\"4+3i\"))";
        ws.Cell("A4").FormulaA1 = "IMPRODUCT(IMCSCH(\"4+3i\"), IMSINH(\"4+3i\"))";

        foreach (var address in new[] { "A1", "A2", "A3", "A4" })
        {
            await Assert.That((double)XLWorkbook.EvaluateExpr($"IMREAL(\"{ws.Cell(address).Value}\")")).IsEqualTo(1d).Within(1e-12);
            await Assert.That((double)XLWorkbook.EvaluateExpr($"IMAGINARY(\"{ws.Cell(address).Value}\")")).IsEqualTo(0d).Within(1e-12);
        }
    }

    [Test]
    [Arguments("IMLN(\"0\")")] // The logarithm is singular at the origin.
    [Arguments("IMLOG10(\"0\")")]
    [Arguments("IMLOG2(\"0\")")]
    [Arguments("IMPOWER(\"0\", 0)")] // 0^0 is undefined.
    [Arguments("IMPOWER(\"0\", -1)")]
    [Arguments("IMDIV(\"1+i\", \"0\")")]
    [Arguments("IMCSC(\"0\")")] // 1/sin(0).
    [Arguments("IMCOT(\"0\")")]
    [Arguments("IMCSCH(\"0\")")]
    public async Task ComplexFunctions_UndefinedResultsReturnNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments("IMABS(\"3+4k\")")] // Not a recognised imaginary unit.
    [Arguments("IMABS(\"3+4I\")")] // The suffix is case sensitive.
    [Arguments("IMABS(\"abc\")")]
    [Arguments("IMABS(\"3++4i\")")]
    [Arguments("IMREAL(\"1,5+2i\")")] // Complex text always uses a period.
    public async Task ComplexFunctions_MalformedTextReturnsNumberInvalid(string formula)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(formula)).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments("IMREAL(\"1.5E+3+2i\")", 1500d)] // An exponent's sign is not the imaginary part's sign.
    [Arguments("IMAGINARY(\"1.5E+3+2i\")", 2d)]
    [Arguments("IMREAL(\"-2.5e-2-3i\")", -0.025d)]
    [Arguments("IMAGINARY(\"-2.5e-2-3i\")", -3d)]
    [Arguments("IMREAL(\"5\")", 5d)]
    [Arguments("IMAGINARY(\"-i\")", -1d)]
    [Arguments("IMAGINARY(\"+i\")", 1d)]
    public async Task ComplexParsing_HandlesExponentsAndImplicitCoefficients(string formula, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(expected).Within(1e-12);
    }

    [Test]
    public async Task ComplexFunctions_RoundTripThroughAWorksheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 3;
        ws.Cell("A2").Value = 4;
        ws.Cell("B1").FormulaA1 = "COMPLEX(A1, A2)";
        ws.Cell("B2").FormulaA1 = "IMABS(B1)";
        ws.Cell("B3").FormulaA1 = "IMPRODUCT(B1, IMCONJUGATE(B1))"; // z·z̄ = |z|².

        await Assert.That(ws.Cell("B1").Value).IsEqualTo("3+4i");
        await Assert.That((double)ws.Cell("B2").Value).IsEqualTo(5d).Within(1e-12);
        await Assert.That(ws.Cell("B3").Value).IsEqualTo("25");
    }
}
