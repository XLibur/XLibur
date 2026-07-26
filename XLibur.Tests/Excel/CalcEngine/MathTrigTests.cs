
using XLibur.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

[SetCulture("en-US")]
public class MathTrigTests
{
    private const double tolerance = 1e-10;

    private static readonly int[] MultinomialDataB2 = [2, 0, 5];
    private static readonly int[] MultinomialDataA5 = [3, 6];

    // NUnit's [Range(from, to)] is inclusive at both ends.
    public static IEnumerable<int> InvalidBaseRadixes() =>
        Enumerable.Range(-2, 1 - -2 + 1).Concat(Enumerable.Range(37, 40 - 37 + 1));

    public static IEnumerable<int> InvalidDecimalRadixes() =>
        Enumerable.Range(37, 255 - 37 + 1).Concat(Enumerable.Range(-5, 1 - -5 + 1));

    [Test]
    [MatrixDataSource]
    public async Task Abs_ReturnsItselfOnPositiveNumbers([MatrixRange<double>(0, 10, 0.1)] double input)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ABS({input.ToString(CultureInfo.InvariantCulture)})");
        await Assert.That(actual).IsEqualTo(input).Within(tolerance * 10);
    }

    [Test]
    [MatrixDataSource]
    public async Task Abs_ReturnsTheCorrectValueOnNegativeInput([MatrixRange<double>(-10, -0.1, 0.1)] double input)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ABS({input.ToString(CultureInfo.InvariantCulture)})");
        await Assert.That(actual).IsEqualTo(-input).Within(tolerance * 10);
    }

    [Test]
    [Arguments(-1, 3.141592654)]
    [Arguments(-0.9, 2.690565842)]
    [Arguments(-0.8, 2.498091545)]
    [Arguments(-0.7, 2.346193823)]
    [Arguments(-0.6, 2.214297436)]
    [Arguments(-0.5, 2.094395102)]
    [Arguments(-0.4, 1.982313173)]
    [Arguments(-0.3, 1.875488981)]
    [Arguments(-0.2, 1.772154248)]
    [Arguments(-0.1, 1.670963748)]
    [Arguments(0, 1.570796327)]
    [Arguments(0.1, 1.470628906)]
    [Arguments(0.2, 1.369438406)]
    [Arguments(0.3, 1.266103673)]
    [Arguments(0.4, 1.159279481)]
    [Arguments(0.5, 1.047197551)]
    [Arguments(0.6, 0.927295218)]
    [Arguments(0.7, 0.79539883)]
    [Arguments(0.8, 0.643501109)]
    [Arguments(0.9, 0.451026812)]
    [Arguments(1, 0)]
    public async Task Acos_ReturnsCorrectValue(double input, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ACOS({input})");
        await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance * 10);
    }

    [Test]
    [MatrixDataSource]
    public async Task Acos_returns_error_when_number_outside_range([MatrixRange<double>(1.1, 3, 0.1)] double input)
    {
        // checking input and it's additive inverse as both are outside range.
        await Assert.That(XLWorkbook.EvaluateExpr($"ACOS({input})")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr($"ACOS({-input})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [MatrixDataSource]
    public async Task Acosh_NumbersBelow1ThrowNumberException([MatrixRange<double>(-1, 0.9, 0.1)] double input)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"ACOSH({input})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(1.2, 0.622362504)]
    [Arguments(1.5, 0.96242365)]
    [Arguments(1.8, 1.192910731)]
    [Arguments(2.1, 1.372859144)]
    [Arguments(2.4, 1.522079367)]
    [Arguments(2.7, 1.650193455)]
    [Arguments(3, 1.762747174)]
    [Arguments(3.3, 1.863279351)]
    [Arguments(3.6, 1.954207529)]
    [Arguments(3.9, 2.037266466)]
    [Arguments(4.2, 2.113748231)]
    [Arguments(4.5, 2.184643792)]
    [Arguments(4.8, 2.250731414)]
    [Arguments(5.1, 2.312634419)]
    [Arguments(5.4, 2.370860342)]
    [Arguments(5.7, 2.425828318)]
    [Arguments(6, 2.47788873)]
    [Arguments(1, 0)]
    public async Task Acosh_returns_correct_number(double angle, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ACOSH({angle})");
        await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance * 10);
    }

    [Test]
    [Arguments(-10, 3.041924001)]
    [Arguments(-9, 3.030935432)]
    [Arguments(-8, 3.017237659)]
    [Arguments(-7, 2.999695599)]
    [Arguments(-6, 2.976443976)]
    [Arguments(-5, 2.944197094)]
    [Arguments(-4, 2.89661399)]
    [Arguments(-3, 2.819842099)]
    [Arguments(-2, 2.677945045)]
    [Arguments(-1, 2.35619449)]
    [Arguments(0, 1.570796327)]
    [Arguments(1, 0.785398163)]
    [Arguments(2, 0.463647609)]
    [Arguments(3, 0.321750554)]
    [Arguments(4, 0.244978663)]
    [Arguments(5, 0.19739556)]
    [Arguments(6, 0.165148677)]
    [Arguments(7, 0.141897055)]
    [Arguments(8, 0.124354995)]
    [Arguments(9, 0.110657221)]
    [Arguments(10, 0.099668652)]
    public async Task Acot_returns_correct_number(double angle, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ACOT({angle})");
        await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance * 10);
    }

    [Test]
    [MatrixDataSource]
    public async Task Acoth_returns_error_for_absolute_angle_smaller_than_one([MatrixRange<double>(-0.9, 0.9, 0.1)] double angle)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"ACOTH({angle})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(-10, -0.100335348)]
    [Arguments(-9, -0.111571776)]
    [Arguments(-8, -0.125657214)]
    [Arguments(-7, -0.143841036)]
    [Arguments(-6, -0.168236118)]
    [Arguments(-5, -0.202732554)]
    [Arguments(-4, -0.255412812)]
    [Arguments(-3, -0.34657359)]
    [Arguments(-2, -0.549306144)]
    [Arguments(2, 0.549306144)]
    [Arguments(3, 0.34657359)]
    [Arguments(4, 0.255412812)]
    [Arguments(5, 0.202732554)]
    [Arguments(6, 0.168236118)]
    [Arguments(7, 0.143841036)]
    [Arguments(8, 0.125657214)]
    [Arguments(9, 0.111571776)]
    [Arguments(10, 0.100335348)]
    [Arguments(1E+100, 1E-100)]
    public async Task Acoth_returns_correct_number(double angle, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ACOTH({angle})");
        await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance * 10);
    }

    [Test]
    [Arguments("LVII", 57)]
    [Arguments("mcmxii", 1912)]
    [Arguments("", 0)]
    [Arguments("-IV", -4)]
    [Arguments("   XIV   ", 14)]
    [Arguments("MCMLXXXIII ", 1983)]
    [Arguments("IIIIIIIIM", 992)]
    [Arguments("CIVIIX", 102)]
    [Arguments("IIX", 8)]
    [Arguments("VIII", 8)]
    public async Task Arabic_returns_correct_number(string roman, int arabic)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ARABIC(\"{roman}\")");
        await Assert.That(actual).IsEqualTo(arabic);
    }

    [Test]
    public async Task Arabic_solitary_minus_is_not_valid_roman_number()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("ARABIC(\"-\")")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Arabic_can_have_at_most_255_chars()
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"ARABIC(\"{new string('M', 255)}\")")).IsEqualTo(255000);
        await Assert.That(XLWorkbook.EvaluateExpr($"ARABIC(\"{new string('M', 256)}\")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("- I")]
    [Arguments("roman")]
    public async Task Arabic_returns_conversion_error_on_invalid_numbers(string invalidRoman)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"ARABIC(\"{invalidRoman}\")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments(-1, -1.570796327)]
    [Arguments(-0.9, -1.119769515)]
    [Arguments(-0.8, -0.927295218)]
    [Arguments(-0.7, -0.775397497)]
    [Arguments(-0.6, -0.643501109)]
    [Arguments(-0.5, -0.523598776)]
    [Arguments(-0.4, -0.411516846)]
    [Arguments(-0.3, -0.304692654)]
    [Arguments(-0.2, -0.201357921)]
    [Arguments(-0.1, -0.100167421)]
    [Arguments(0, 0)]
    [Arguments(0.1, 0.100167421)]
    [Arguments(0.2, 0.201357921)]
    [Arguments(0.3, 0.304692654)]
    [Arguments(0.4, 0.411516846)]
    [Arguments(0.5, 0.523598776)]
    [Arguments(0.6, 0.643501109)]
    [Arguments(0.7, 0.775397497)]
    [Arguments(0.8, 0.927295218)]
    [Arguments(0.9, 1.119769515)]
    [Arguments(1, 1.570796327)]
    public async Task Asin_ReturnsCorrectResult(double input, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ASIN({input})");
        await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance * 10);
    }

    [Test]
    [MatrixDataSource]
    public async Task Asin_ThrowsNumberExceptionWhenAbsOfInputGreaterThan1([MatrixRange<double>(-3, -1.1, 0.1)] double input)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"ASIN({input})")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr($"ASIN({-input})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(0, 0)]
    [Arguments(0.1, 0.0998340788992076)]
    [Arguments(0.2, 0.198690110349241)]
    [Arguments(0.3, 0.295673047563422)]
    [Arguments(0.4, 0.390035319770715)]
    [Arguments(0.5, 0.481211825059603)]
    [Arguments(0.6, 0.568824898732248)]
    [Arguments(0.7, 0.652666566082356)]
    [Arguments(0.8, 0.732668256045411)]
    [Arguments(0.9, 0.808866935652783)]
    [Arguments(1, 0.881373587019543)]
    [Arguments(2, 1.44363547517881)]
    [Arguments(3, 1.81844645923207)]
    [Arguments(4, 2.0947125472611)]
    [Arguments(5, 2.31243834127275)]
    public async Task Asinh_ReturnsCorrectResult(double input, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ASINH({input})");
        await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance);
        var minusActual = (double)XLWorkbook.EvaluateExpr($"ASINH({-input})");
        await Assert.That(minusActual).IsEqualTo(-expectedResult).Within(tolerance);
    }

    [Test]
    [Arguments(0, 0)]
    [Arguments(0.1, 0.099668652491162)]
    [Arguments(0.2, 0.197395559849881)]
    [Arguments(0.3, 0.291456794477867)]
    [Arguments(0.4, 0.380506377112365)]
    [Arguments(0.5, 0.463647609000806)]
    [Arguments(0.6, 0.540419500270584)]
    [Arguments(0.7, 0.610725964389209)]
    [Arguments(0.8, 0.674740942223553)]
    [Arguments(0.9, 0.732815101786507)]
    [Arguments(1, 0.785398163397448)]
    [Arguments(2, 1.10714871779409)]
    [Arguments(3, 1.24904577239825)]
    [Arguments(4, 1.32581766366803)]
    [Arguments(5, 1.37340076694502)]
    public async Task Atan_ReturnsCorrectResult(double input, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ATAN({input})");
        await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance);
        var minusActual = (double)XLWorkbook.EvaluateExpr($"ATAN({-input})");
        await Assert.That(minusActual).IsEqualTo(-expectedResult).Within(tolerance);
    }

    [Test]
    [MatrixDataSource]
    public async Task Atan2_Returns0OnSecond0AndFirstGreater0([MatrixRange<double>(0.1, 5, 0.4)] double input)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({input}, 0)");
        await Assert.That(actual).IsEqualTo(0).Within(tolerance);
    }

    [Test]
    [Arguments(1, 2, 1.10714871779409)]
    [Arguments(1, 3, 1.24904577239825)]
    [Arguments(2, 3, 0.98279372324733)]
    [Arguments(1, 4, 1.32581766366803)]
    [Arguments(3, 4, 0.92729521800161)]
    [Arguments(1, 5, 1.37340076694502)]
    [Arguments(2, 5, 1.19028994968253)]
    [Arguments(3, 5, 1.03037682652431)]
    [Arguments(4, 5, 0.89605538457134)]
    [Arguments(1, 6, 1.40564764938027)]
    [Arguments(5, 6, 0.87605805059819)]
    [Arguments(1, 7, 1.42889927219073)]
    [Arguments(2, 7, 1.29249666778979)]
    [Arguments(3, 7, 1.16590454050981)]
    [Arguments(4, 7, 1.05165021254837)]
    [Arguments(5, 7, 0.95054684081208)]
    [Arguments(6, 7, 0.86217005466723)]
    public async Task Atan2_ReturnsCorrectResults_EqualOnAllMultiplesOfFraction(double x, double y, double expectedResult)
    {
        for (var i = 1; i < 5; i++)
        {
            var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({x * i}, {y * i})");
            await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance);
        }
    }

    [Test]
    [MatrixDataSource]
    public async Task Atan2_ReturnsHalfPiOn0AsFirstInputWhenSecondGreater0([MatrixRange<double>(0.1, 5, 0.4)] double input)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2(0, {input})");
        await Assert.That(actual).IsEqualTo(0.5 * Math.PI).Within(tolerance);
    }

    [Test]
    [MatrixDataSource]
    public async Task Atan2_ReturnsMinus3QuartersOfPiWhenFirstSmaller0AndSecondItsNegative([MatrixRange<double>(-5, -0.1, 0.3)] double input)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({input}, {input})");
        await Assert.That(actual).IsEqualTo(-0.75 * Math.PI).Within(tolerance);
    }

    [Test]
    [MatrixDataSource]
    public async Task Atan2_ReturnsMinusHalfPiOn0AsFirstInputWhenSecondSmaller0([MatrixRange<double>(-5, -0.1, 0.4)] double input)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2(0, {input})");
        await Assert.That(actual).IsEqualTo(-0.5 * Math.PI).Within(tolerance);
    }

    [Test]
    [MatrixDataSource]
    public async Task Atan2_ReturnsPiOn0AsSecondInputWhenFirstSmaller0([MatrixRange<double>(-5, -0.1, 0.4)] double input)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({input}, 0)");
        await Assert.That(actual).IsEqualTo(Math.PI).Within(tolerance);
    }

    [Test]
    [MatrixDataSource]
    public async Task Atan2_ReturnsQuarterOfPiWhenInputsAreEqualAndGreater0([MatrixRange<double>(0.1, 5, 0.3)] double input)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ATAN2({input}, {input})");
        await Assert.That(actual).IsEqualTo(0.25 * Math.PI).Within(tolerance);
    }

    [Test]
    public async Task Atan2_ThrowsDiv0ExceptionOn0And0()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("ATAN2(0, 0)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    [Arguments(-0.99, -2.64665241236225)]
    [Arguments(-0.9, -1.47221948958322)]
    [Arguments(-0.8, -1.09861228866811)]
    [Arguments(-0.6, -0.693147180559945)]
    [Arguments(-0.4, -0.423648930193602)]
    [Arguments(-0.2, -0.202732554054082)]
    [Arguments(0, 0)]
    [Arguments(0.2, 0.202732554054082)]
    [Arguments(0.4, 0.423648930193602)]
    [Arguments(0.6, 0.693147180559945)]
    [Arguments(0.8, 1.09861228866811)]
    [Arguments(-0.9, -1.47221948958322)]
    [Arguments(-0.990, -2.64665241236225)]
    [Arguments(-0.999, -3.8002011672502)]
    public async Task Atanh_ReturnsCorrectResults(double input, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"ATANH({input})");
        await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance * 10);
    }

    [Test]
    [MatrixDataSource]
    public async Task Atanh_ThrowsNumberExceptionWhenAbsOfInput1OrGreater([MatrixRange<double>(1, 5, 0.2)] double input)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"ATANH({input})")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr($"ATANH({-input})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(0, 36, "0")]
    [Arguments(1, 36, "1")]
    [Arguments(2, 36, "2")]
    [Arguments(3, 36, "3")]
    [Arguments(4, 36, "4")]
    [Arguments(5, 36, "5")]
    [Arguments(6, 36, "6")]
    [Arguments(7, 36, "7")]
    [Arguments(8, 36, "8")]
    [Arguments(9, 36, "9")]
    [Arguments(10, 36, "A")]
    [Arguments(11, 36, "B")]
    [Arguments(12, 36, "C")]
    [Arguments(13, 36, "D")]
    [Arguments(14, 36, "E")]
    [Arguments(15, 36, "F")]
    [Arguments(16, 36, "G")]
    [Arguments(17, 36, "H")]
    [Arguments(18, 36, "I")]
    [Arguments(19, 36, "J")]
    [Arguments(20, 36, "K")]
    [Arguments(21, 36, "L")]
    [Arguments(22, 36, "M")]
    [Arguments(23, 36, "N")]
    [Arguments(24, 36, "O")]
    [Arguments(25, 36, "P")]
    [Arguments(26, 36, "Q")]
    [Arguments(27, 36, "R")]
    [Arguments(28, 36, "S")]
    [Arguments(29, 36, "T")]
    [Arguments(30, 36, "U")]
    [Arguments(31, 36, "V")]
    [Arguments(32, 36, "W")]
    [Arguments(33, 36, "X")]
    [Arguments(34, 36, "Y")]
    [Arguments(35, 36, "Z")]
    [Arguments(36, 36, "10")]
    [Arguments(255, 29, "8N")]
    [Arguments(255, 2, "11111111")]
    public async Task Base_returns_number_in_specified_base(int input, int radix, string expectedResult)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"BASE({input},{radix})");
        await Assert.That(actual).IsEqualTo(expectedResult);
    }

    [Test]
    [Arguments(255, 2, 3, "11111111")]
    [Arguments(255, 2, 8, "11111111")]
    [Arguments(255, 2, 10, "0011111111")]
    [Arguments(10, 3, 4, "0101")]
    [Arguments(0, 10, 0, "")]
    public async Task Base_returns_text_of_at_least_minimal_length(int input, int radix, int minLength, string expectedResult)
    {
        var actual = (string)XLWorkbook.EvaluateExpr($"BASE({input},{radix},{minLength})");
        await Assert.That(actual).IsEqualTo(expectedResult);
    }

    [Test]
    public async Task Base_min_length_must_be_at_most_255()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("BASE(0,2,256)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(@"""x""", "2", "2")]
    [Arguments("0", @"""x""", "2")]
    [Arguments("0", "2", @"""x""")]
    public async Task Base_coercion(string input, string radix, string minLength)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"BASE({input},{radix},{minLength})")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    // Was [Range(-2, 1), Range(37, 40)] in NUnit, which unions both ranges into one set of
    // values for the parameter. TUnit's [MatrixRange] cannot be repeated, so the union is
    // supplied explicitly instead.
    [MethodDataSource(nameof(InvalidBaseRadixes))]
    public async Task Base_radix_must_be_between_2_and_36(int radix)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"BASE(0,{radix})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [MatrixDataSource]
    public async Task Base_number_must_be_zero_or_positive([MatrixRange<int>(-5, -1)] int input)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"BASE({input},2)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Base_number_must_fit_in_double_without_precision_loss()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("BASE(9.007E+15,36)")).IsEqualTo("2GOPQOE5GCG");
        await Assert.That(XLWorkbook.EvaluateExpr("BASE(9.008E+15,36)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(24.3, 5, 25)]
    [Arguments(6.7, 1, 7)]
    [Arguments(-8.1, 2, -8)]
    [Arguments(5.5, 2.1, 6.3)]
    [Arguments(5.5, 0, 0)]
    [Arguments(-5.5, 2.1, -4.2)]
    [Arguments(-5.5, -2.1, -6.3)]
    [Arguments(-5.5, 0, 0)]
    [Arguments(0, 0, 0)]
    [Arguments(0, 0.1, 0)]
    [Arguments(0, -0.1, 0)]
    [Arguments(0.1, 0, 0)]
    [Arguments(-0.1, 0, 0)]
    public async Task Ceiling(double input, double significance, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"CEILING({input}, {significance})");
        await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance);
    }

    [Test]
    [Arguments(6.7, -1)]
    [Arguments(0.1, -0.2)]
    public async Task Ceiling_returns_error_on_different_number_and_significance(double input, double significance)
    {
        // Spec says "if x and significance have different signs, #NUM! is returned.",
        // but in reality it only happens when number is positive and step negative.
        await Assert.That(XLWorkbook.EvaluateExpr($"CEILING({input}, {significance})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(24.3, 5, null, 25)]
    [Arguments(6.7, null, null, 7)]
    [Arguments(-8.1, 2, null, -8)]
    [Arguments(-5.5, 2, -1, -6)]
    [Arguments(-5.5, 2, -0.1, -6)]
    [Arguments(5.5, 2.1, 0, 6.3)]
    [Arguments(5.5, -2.1, 0, 6.3)]
    [Arguments(5.5, 0, 0, 0)]
    [Arguments(5.5, 2.1, -1, 6.3)]
    [Arguments(5.5, -2.1, -1, 6.3)]
    [Arguments(5.5, 0, -1, 0)]
    [Arguments(5.5, 2.1, 10, 6.3)]
    [Arguments(5.5, -2.1, 10, 6.3)]
    [Arguments(5.5, 0, 10, 0)]
    [Arguments(-5.5, 2.1, 0, -4.2)]
    [Arguments(-5.5, -2.1, 0, -4.2)]
    [Arguments(-5.5, 0, 0, 0)]
    [Arguments(-5.5, 2.1, -1, -6.3)]
    [Arguments(-5.5, -2.1, -1, -6.3)]
    [Arguments(-5.5, 0, -1, 0)]
    [Arguments(-5.5, 2.1, 10, -6.3)]
    [Arguments(-5.5, -2.1, 10, -6.3)]
    [Arguments(-5.5, 0, 10, 0)]
    public async Task CeilingMath(double input, double? significance, double? mode, double expectedResult)
    {
        var parameters = new StringBuilder();
        parameters.Append(input);
        if (significance != null)
        {
            parameters.Append(", ").Append(significance);
            if (mode != null)
                parameters.Append(", ").Append(mode);
        }

        var actual = (double)XLWorkbook.EvaluateExpr($"CEILING.MATH({parameters})");
        await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance);
    }

    [Test]
    public async Task Combin()
    {
        var actual1 = XLWorkbook.EvaluateExpr("COMBIN(200, 2)");
        await Assert.That(actual1).IsEqualTo(19900.0);

        var actual2 = XLWorkbook.EvaluateExpr("COMBIN(20.1, 2.9)");
        await Assert.That(actual2).IsEqualTo(190.0);
    }

    [Test]
    [MatrixDataSource]
    public async Task Combin_returns_1_for_k_is_0_or_k_equals_n([MatrixRange<int>(0, 10)] int n)
    {
        var actual = XLWorkbook.EvaluateExpr($"COMBIN({n}, 0)");
        await Assert.That(actual).IsEqualTo(1);

        var actual2 = XLWorkbook.EvaluateExpr($"COMBIN({n}, {n})");
        await Assert.That(actual2).IsEqualTo(1);
    }

    [Test]
    [Arguments(0, 0, 1)]
    [Arguments(1, 0, 1)]
    [Arguments(1, 1, 1)]
    [Arguments(4, 2, 6)]
    [Arguments(5, 2, 10)]
    [Arguments(6, 2, 15)]
    [Arguments(6, 3, 20)]
    [Arguments(7, 2, 21)]
    [Arguments(7, 3, 35)]
    public async Task Combin_calculates_combinations(int n, int k, int expectedResult)
    {
        var actual = XLWorkbook.EvaluateExpr($"COMBIN({n}, {k})");
        await Assert.That(actual).IsEqualTo(expectedResult);

        var actual2 = XLWorkbook.EvaluateExpr($"COMBIN({n}, {n - k})");
        await Assert.That(actual2).IsEqualTo(expectedResult);
    }

    [Test]
    [MatrixDataSource]
    public async Task Combin_returns_n_for_k_is_1_or_k_is_n_minus_1([MatrixRange<int>(1, 10)] int n)
    {
        var actual = XLWorkbook.EvaluateExpr($"COMBIN({n}, 1)");
        await Assert.That(actual).IsEqualTo(n);

        var actual2 = XLWorkbook.EvaluateExpr($"COMBIN({n}, {n - 1})");
        await Assert.That(actual2).IsEqualTo(n);
    }

    [Test]
    public async Task Combin_returns_num_error_when_k_is_larger_than_n()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("COMBIN(5, 6)")).IsEqualTo(XLError.NumberInvalid);

        // Values are floored, so this is COMBIN(5, 5).
        await Assert.That(XLWorkbook.EvaluateExpr("COMBIN(5, 5.5)")).IsEqualTo(1);
    }

    [Test]
    public async Task Combin_returns_num_error_when_value_is_too_large()
    {
        // Maximum int - 1 is maximum computable value in Excel.
        await Assert.That(XLWorkbook.EvaluateExpr("COMBIN(2147483647, 2147483647)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("COMBIN(5E+301, 6)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("COMBIN(6, 5E+301)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(-4)]
    [Arguments(-3)]
    [Arguments(-1)]
    [Arguments(-0.1)]
    public async Task Combin_returns_num_error_for_any_argument_smaller_than_0(double smaller0)
    {
        await Assert.That(XLWorkbook.EvaluateExpr(
            $"COMBIN({smaller0.ToString(CultureInfo.InvariantCulture)}, {(-smaller0).ToString(CultureInfo.InvariantCulture)})")).IsEqualTo(XLError.NumberInvalid);

        await Assert.That(XLWorkbook.EvaluateExpr(
            $"COMBIN({(-smaller0).ToString(CultureInfo.InvariantCulture)}, {smaller0.ToString(CultureInfo.InvariantCulture)})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments("\"no number\"")]
    [Arguments("\"\"")]
    public async Task Combin_returns_value_error_for_any_non_numeric_argument(string input)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"COMBIN({input}, 1)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr($"COMBIN(1, {input})")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments(4, 3, 20)]
    [Arguments(10, 3, 220)]
    [Arguments(0, 0, 1)]
    [Arguments(1, 0, 1)]
    [Arguments(10, 15, 1307504)]
    public async Task Combina_calculates_correct_values(int number, int chosen, int expectedResult)
    {
        var actualResult = XLWorkbook.EvaluateExpr($"COMBINA({number}, {chosen})");
        await Assert.That(actualResult).IsEqualTo(expectedResult);
    }

    [Test]
    [MatrixDataSource]
    public async Task Combina_returns_one_when_chosen_is_zero([MatrixRange<int>(0, 10)] int number)
    {
        var actualResult = XLWorkbook.EvaluateExpr($"COMBINA({number}, 0)");
        await Assert.That(actualResult).IsEqualTo(1);
    }

    [Test]
    [Arguments(-1, 2)]
    [Arguments(-3, -2)]
    [Arguments(2, -2)]
    [Arguments(int.MaxValue + 1d, 1)]
    public async Task Combina_returns_error_on_invalid_values(double number, int chosen)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"COMBINA({number}, {chosen})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(4.23, 3, 20)]
    [Arguments(10.4, 3.14, 220)]
    [Arguments(0, 0.4, 1)]
    public async Task Combina_truncates_numbers_to_zero(double number, double chosen, int expectedResult)
    {
        var actualResult = XLWorkbook.EvaluateExpr($"COMBINA({number}, {chosen})");
        await Assert.That(actualResult).IsEqualTo(expectedResult);
    }

    [Test]
    [Arguments(0, 1)]
    [Arguments(0.4, 0.921060994002885)]
    [Arguments(0.8, 0.696706709347165)]
    [Arguments(1.2, 0.362357754476674)]
    [Arguments(1.6, -0.0291995223012888)]
    [Arguments(2, -0.416146836547142)]
    [Arguments(2.4, -0.737393715541245)]
    [Arguments(2.8, -0.942222340668658)]
    [Arguments(3.2, -0.998294775794753)]
    [Arguments(3.6, -0.896758416334147)]
    [Arguments(4, -0.653643620863612)]
    [Arguments(4.4, -0.307332869978419)]
    [Arguments(4.8, 0.0874989834394464)]
    [Arguments(5.2, 0.468516671300377)]
    [Arguments(5.6, 0.77556587851025)]
    [Arguments(6, 0.960170286650366)]
    [Arguments(6.4, 0.993184918758193)]
    [Arguments(6.8, 0.869397490349825)]
    [Arguments(7.2, 0.608351314532255)]
    [Arguments(7.6, 0.251259842582256)]
    [Arguments(8, -0.145500033808614)]
    [Arguments(8.4, -0.519288654116686)]
    public async Task Cos_ReturnsCorrectResult(double input, double expectedResult)
    {
        var actualResult = (double)XLWorkbook.EvaluateExpr($"COS({input})");
        await Assert.That(actualResult).IsEqualTo(expectedResult).Within(tolerance);
    }

    [Test]
    [Arguments(0, 1)]
    [Arguments(0.4, 1.08107237183845)]
    [Arguments(0.8, 1.33743494630484)]
    [Arguments(1.2, 1.81065556732437)]
    [Arguments(1.6, 2.57746447119489)]
    [Arguments(2, 3.76219569108363)]
    [Arguments(2.4, 5.55694716696551)]
    [Arguments(2.8, 8.25272841686113)]
    [Arguments(3.2, 12.2866462005439)]
    [Arguments(3.6, 18.3127790830626)]
    [Arguments(4, 27.3082328360165)]
    [Arguments(4.4, 40.7315730024356)]
    [Arguments(4.8, 60.7593236328919)]
    [Arguments(5.2, 90.638879219786)]
    [Arguments(5.6, 135.215052644935)]
    [Arguments(6, 201.715636122456)]
    [Arguments(6.4, 300.923349714678)]
    [Arguments(6.8, 448.924202712783)]
    [Arguments(7.2, 669.715755490113)]
    [Arguments(7.6, 999.098197777775)]
    [Arguments(8, 1490.47916125218)]
    [Arguments(8.4, 2223.53348628359)]
    public async Task Cosh_ReturnsCorrectResult(double input, double expectedResult)
    {
        var actualResult = (double)XLWorkbook.EvaluateExpr($"COSH({input})");
        await Assert.That(actualResult).IsEqualTo(expectedResult).Within(tolerance);
        var actualResult2 = (double)XLWorkbook.EvaluateExpr($"COSH({-input})");
        await Assert.That(actualResult2).IsEqualTo(expectedResult).Within(tolerance);
    }

    [Test]
    [Arguments(711)]
    [Arguments(-711)]
    [Arguments(100000)]
    public async Task Cosh_too_large_returns_error(double input)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"COSH({input})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(1, 0.642092616)]
    [Arguments(2, -0.457657554)]
    [Arguments(3, -7.015252551)]
    [Arguments(4, 0.863691154)]
    [Arguments(5, -0.295812916)]
    [Arguments(6, -3.436353004)]
    [Arguments(7, 1.147515422)]
    [Arguments(8, -0.147065064)]
    [Arguments(9, -2.210845411)]
    [Arguments(10, 1.542351045)]
    [Arguments(11, -0.004425741)]
    [Arguments(Math.PI * 0.5, 0)]
    [Arguments(45, 0.617369624)]
    [Arguments(-2, 0.457657554)]
    [Arguments(-3, 7.015252551)]
    public async Task Cot(double angle, double expected)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"COT({angle})");
        await Assert.That(actual).IsEqualTo(expected).Within(tolerance * 10.0);
    }

    [Test]
    public async Task Cot_returns_division_by_zero_error_on_angle_zero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("COT(0)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task Coth_returns_division_by_zero_error_on_angle_zero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("COTH(0)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    [Arguments(-10, -1.000000004)]
    [Arguments(-9, -1.00000003)]
    [Arguments(-8, -1.000000225)]
    [Arguments(-7, -1.000001663)]
    [Arguments(-6, -1.000012289)]
    [Arguments(-5, -1.000090804)]
    [Arguments(-4, -1.00067115)]
    [Arguments(-3, -1.004969823)]
    [Arguments(-2, -1.037314721)]
    [Arguments(-1, -1.313035285)]
    [Arguments(1, 1.313035285)]
    [Arguments(2, 1.037314721)]
    [Arguments(3, 1.004969823)]
    [Arguments(4, 1.00067115)]
    [Arguments(5, 1.000090804)]
    [Arguments(6, 1.000012289)]
    [Arguments(7, 1.000001663)]
    [Arguments(8, 1.000000225)]
    [Arguments(9, 1.00000003)]
    [Arguments(10, 1.000000004)]
    public async Task Coth_returns_correct_number(double angle, double expected)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"COTH({angle})");
        await Assert.That(actual).IsEqualTo(expected).Within(tolerance * 10.0);
    }

    [Test]
    public async Task Csc_returns_division_by_zero_on_angle_zero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("CSC(0)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    [Arguments(-10, 1.838163961)]
    [Arguments(-9, -2.426486644)]
    [Arguments(-8, -1.010756218)]
    [Arguments(-7, -1.522101063)]
    [Arguments(-6, 3.578899547)]
    [Arguments(-5, 1.042835213)]
    [Arguments(-4, 1.321348709)]
    [Arguments(-3, -7.086167396)]
    [Arguments(-2, -1.09975017)]
    [Arguments(-1, -1.188395106)]
    [Arguments(1, 1.188395106)]
    [Arguments(2, 1.09975017)]
    [Arguments(3, 7.086167396)]
    [Arguments(4, -1.321348709)]
    [Arguments(5, -1.042835213)]
    [Arguments(6, -3.578899547)]
    [Arguments(7, 1.522101063)]
    [Arguments(8, 1.010756218)]
    [Arguments(9, 2.426486644)]
    [Arguments(10, -1.838163961)]
    public async Task Csc_returns_correct_number(double angle, double expected)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"CSC({angle})");
        await Assert.That(actual).IsEqualTo(expected).Within(tolerance * 10);
    }

    [Test]
    [Arguments(1, 0.850918128)]
    [Arguments(2, 0.275720565)]
    [Arguments(3, 0.09982157)]
    [Arguments(4, 0.03664357)]
    [Arguments(5, 0.013476506)]
    [Arguments(6, 0.004957535)]
    [Arguments(7, 0.001823765)]
    [Arguments(8, 0.000670925)]
    [Arguments(9, 0.00024682)]
    [Arguments(10, 0.000090799859712122200000)]
    [Arguments(11, 0.0000334034)]
    public async Task Csch_calculates_correct_values(double input, double expectedOutput)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"CSCH({input})")).IsEqualTo(expectedOutput).Within(0.000000001);
    }

    [Test]
    public async Task Csch_returns_division_error_on_angle_zero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("CSCH(0)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    [Arguments("FF", 16, 255)]
    [Arguments("111", 2, 7)]
    [Arguments("zap", 36, 45745)] // Case insensitive
    [Arguments("  1234", 10, 1234)] // Trims start
    [Arguments("123", 10.9, 123)] // Radix truncated
    [Arguments("1F", 10, XLError.NumberInvalid)]
    [Arguments("", 10, 0)]
    public async Task Decimal(string inputString, double radix, object expectedResult)
    {
        var actualResult = XLWorkbook.EvaluateExpr($"DECIMAL(\"{inputString}\", {radix})");
        await Assert.That(actualResult).IsEqualTo(ExpectedCellValue.From(expectedResult));
    }

    [Test]
    // Was [Range(37, 255), Range(-5, 1)] in NUnit -- see InvalidBaseRadixes for why this
    // is a data source rather than a repeated [MatrixRange].
    [MethodDataSource(nameof(InvalidDecimalRadixes))]
    public async Task Decimal_radix_must_be_between_2_and_36(int radix)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"DECIMAL(\"0\", {radix})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [MatrixDataSource]
    public async Task Decimal_zero_is_zero_in_any_radix([MatrixRange<int>(2, 36)] int radix)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"DECIMAL(\"0\", {radix})")).IsEqualTo(0);
    }

    [Test]
    public async Task Decimal_text_must_be_less_than_256_chars_long()
    {
        var text = new string('0', 256);
        await Assert.That(XLWorkbook.EvaluateExpr($"DECIMAL(\"{text}\", 10)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Decimal_returns_number_invalid_when_result_out_of_bounds()
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"DECIMAL(\"{new string('Z', 198)}\", 36)")).IsEqualTo(1.4057081148316923E+308d);
        await Assert.That(XLWorkbook.EvaluateExpr($"DECIMAL(\"{new string('Z', 199)}\", 36)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments("101", "\"1 2/2\"", 5)] // 101 in binary is 5
    public async Task Decimal_coercion(string input, string radix, object expectedResult)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"DECIMAL({input}, {radix})")).IsEqualTo(ExpectedCellValue.From(expectedResult));
    }

    [Test]
    public async Task Degrees()
    {
        var actual = (double)XLWorkbook.EvaluateExpr("DEGREES(PI())");
        await Assert.That(actual).IsEqualTo(180).Within(XLHelper.Epsilon);
    }

    [Test]
    [Arguments(0, 0)]
    [Arguments(Math.PI, 180)]
    [Arguments(Math.PI * 2, 360)]
    [Arguments(1, 57.2957795130823)]
    [Arguments(2, 114.591559026165)]
    [Arguments(3, 171.887338539247)]
    [Arguments(4, 229.183118052329)]
    [Arguments(5, 286.478897565412)]
    [Arguments(6, 343.774677078494)]
    [Arguments(7, 401.070456591576)]
    [Arguments(8, 458.366236104659)]
    [Arguments(9, 515.662015617741)]
    [Arguments(10, 572.957795130823)]
    [Arguments(Math.PI * 0.5, 90)]
    [Arguments(Math.PI * 1.5, 270)]
    [Arguments(Math.PI * 0.25, 45)]
    [Arguments(-1, -57.2957795130823)]
    public async Task Degrees_ReturnsCorrectResult(double input, double expected)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"DEGREES({input})");
        await Assert.That(actual).IsEqualTo(expected).Within(tolerance);
    }

    [Test]
    [Arguments(3, 4)]
    [Arguments(2, 2)]
    [Arguments(-1, -2)]
    [Arguments(-2, -2)]
    [Arguments(0, 0)]
    [Arguments(1.5, 2)]
    [Arguments(2.01, 4)]
    [Arguments(1e+100, 1e+100)]
    [Arguments(Math.PI, 4)]
    public async Task Even(double number, double expectedResult)
    {
        var actual = XLWorkbook.EvaluateExpr($"EVEN({number})");
        await Assert.That(actual).IsEqualTo(expectedResult);
    }

    [Test]
    [Arguments(0, 1)]
    [Arguments(1, Math.E)]
    [Arguments(2, 7.38905609893065)]
    [Arguments(3, 20.0855369231877)]
    [Arguments(4, 54.5981500331442)]
    [Arguments(5, 148.413159102577)]
    [Arguments(6, 403.428793492735)]
    [Arguments(7, 1096.63315842846)]
    [Arguments(8, 2980.95798704173)]
    [Arguments(9, 8103.08392757538)]
    [Arguments(10, 22026.4657948067)]
    [Arguments(11, 59874.1417151978)]
    [Arguments(12, 162754.791419004)]
    [Arguments(-1E+100, 0)]
    public async Task Exp_returns_correct_results(double input, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"EXP({input})");
        await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance);
    }

    [Test]
    [Arguments(710)]
    public async Task Exp_with_too_large_result_return_error(double input)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"EXP({input})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Fact()
    {
        object actual = XLWorkbook.EvaluateExpr("Fact(5.9)");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(120.0));
    }

    [Test]
    [Arguments(0, 1d)]
    [Arguments(1, 1d)]
    [Arguments(2, 2d)]
    [Arguments(3, 6d)]
    [Arguments(4, 24d)]
    [Arguments(5, 120d)]
    [Arguments(6, 720d)]
    [Arguments(7, 5040d)]
    [Arguments(8, 40320d)]
    [Arguments(9, 362880d)]
    [Arguments(10, 3628800d)]
    [Arguments(11, 39916800d)]
    [Arguments(12, 479001600d)]
    [Arguments(13, 6227020800d)]
    [Arguments(14, 87178291200d)]
    [Arguments(15, 1307674368000d)]
    [Arguments(16, 20922789888000d)]
    [Arguments(170.9, 7.257415615308004E+306)]
    [Arguments(0.1, 1L)]
    [Arguments(2.3, 2L)]
    [Arguments(2.8, 2L)]
    public async Task Fact_calculates_factorial(double input, double expectedResult)
    {
        var actual = XLWorkbook.EvaluateExpr($"FACT({input.ToString(CultureInfo.InvariantCulture)})");
        await Assert.That(actual).IsEqualTo(expectedResult);
    }

    [Test]
    [Arguments(-10)]
    [Arguments(-5)]
    [Arguments(-1)]
    [Arguments(-0.1)]
    public async Task Fact_returns_error_for_negative_input(double input)
    {
        var actual = XLWorkbook.EvaluateExpr($"FACT({input.ToString(CultureInfo.InvariantCulture)})");
        await Assert.That(actual).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(171)]
    [Arguments(5000)]
    public async Task Fact_returns_error_for_too_large_result(int input)
    {
        var actual = XLWorkbook.EvaluateExpr($"FACT({input})");
        await Assert.That(actual).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Fact_coercion_fails_for_non_numeric_input()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"FACT(""x"")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments(0, 1L)]
    [Arguments(1, 1L)]
    [Arguments(2, 2L)]
    [Arguments(3, 3L)]
    [Arguments(4, 8L)]
    [Arguments(5, 15L)]
    [Arguments(6, 48L)]
    [Arguments(7, 105L)]
    [Arguments(8, 384L)]
    [Arguments(9, 945L)]
    [Arguments(10, 3840L)]
    [Arguments(11, 10395L)]
    [Arguments(12, 46080L)]
    [Arguments(13, 135135L)]
    [Arguments(14, 645120)]
    [Arguments(15, 2027025)]
    [Arguments(16, 10321920)]
    [Arguments(-1, 1L)]
    [Arguments(0, 1)]
    [Arguments(0.1, 1L)]
    [Arguments(1.4, 1L)]
    [Arguments(2.3, 2L)]
    [Arguments(2.8, 2L)]
    public async Task FactDouble_ReturnsCorrectResult(double input, long expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"FACTDOUBLE({input})");
        await Assert.That(actual).IsEqualTo(expectedResult);
    }

    [Test]
    [Arguments(301)]
    [Arguments(1e+100)]
    public async Task FactDouble_returns_error_on_too_large_value(double n)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"FACTDOUBLE({n})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [MatrixDataSource]
    public async Task FactDouble_ThrowsNumberExceptionForInputSmallerThanMinus1([MatrixRange<int>(-10, -2)] int input)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"FACTDOUBLE({input})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task FactDouble_ThrowsValueExceptionForNonNumericInput()
    {
        await Assert.That(XLWorkbook.EvaluateExpr(@"FACTDOUBLE(""x"")")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments(0, 0, 0)]
    [Arguments(0, 1, 0)]
    [Arguments(24.3, 5, 20)]
    [Arguments(6.7, 1, 6)]
    [Arguments(-8.1, 2, -10)]
    [Arguments(5.5, 2.1, 4.2)]
    [Arguments(-5.5, 2.1, -6.3)]
    [Arguments(-5.5, -2.1, -4.2)]
    public async Task Floor(double input, double significance, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"FLOOR({input}, {significance})");
        await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance);
    }

    [Test]
    [Arguments(6.7, 0)]
    [Arguments(-6.7, 0)]
    public async Task Floor_ThrowsDivisionByZeroOnZeroSignificance(double input, double significance)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"FLOOR({input}, {significance})")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    [Arguments(6.7, -1)]
    public async Task Floor_ThrowsNumberExceptionOnInvalidInput(double input, double significance)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"FLOOR({input}, {significance})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    // Functions have to support a period first before we can implement this
    [Arguments(24.3, 5, null, 20)]
    [Arguments(6.7, null, null, 6)]
    [Arguments(-8.1, 2, null, -10)]
    [Arguments(5.5, 2.1, 0, 4.2)]
    [Arguments(5.5, -2.1, 0, 4.2)]
    [Arguments(5.5, 0, 0, 0)]
    [Arguments(5.5, 2.1, -1, 4.2)]
    [Arguments(5.5, -2.1, -1, 4.2)]
    [Arguments(5.5, 0, -2, 0)]
    [Arguments(5.5, 2.1, 10, 4.2)]
    [Arguments(5.5, -2.1, 10, 4.2)]
    [Arguments(5.5, 0, 10, 0)]
    [Arguments(-5.5, 2.1, 0, -6.3)]
    [Arguments(-5.5, -2.1, 0, -6.3)]
    [Arguments(-5.5, 0, 0, 0)]
    [Arguments(-5.5, 2.1, -1, -4.2)]
    [Arguments(-5.5, -2.1, -1, -4.2)]
    [Arguments(-5.5, 0, -1, 0)]
    [Arguments(-5.5, 2.1, 10, -4.2)]
    [Arguments(-5.5, -2.1, 10, -4.2)]
    [Arguments(-5.5, 0, 0, 0)]
    public async Task FloorMath(double input, double? significance, int? mode, double expectedResult)
    {
        var parameters = new StringBuilder();
        parameters.Append(input);
        if (significance != null)
        {
            parameters.Append(", ").Append(significance);
            if (mode != null)
                parameters.Append(", ").Append(mode);
        }

        var actual = (double)XLWorkbook.EvaluateExpr($"FLOOR.MATH({parameters})");
        await Assert.That(actual).IsEqualTo(expectedResult).Within(tolerance);
    }

    [Test]
    [Arguments("24,36", 12)]
    [Arguments("240,360,30", 30)]
    [Arguments("24.9,36.9", 12)]
    [Arguments("{24,36}", 12)]
    [Arguments("{\"24\",\"36\"}", 12)]
    [Arguments("5,0", 5)]
    [Arguments("0,5", 5)]
    public async Task Gcd(string args, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"GCD({args})")).IsEqualTo(expected);
    }

    [Test]
    public async Task Gcd_accepts_references()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new object[]
        {
            (120, 240),
            ("60", "150"),
        });
        await Assert.That(ws.Evaluate("GCD(A1:A2,B1:B2)")).IsEqualTo(30);

        // Blank is considered 0
        await Assert.That(ws.Evaluate("GCD(A1:A3)")).IsEqualTo(60);

        // Logical are not converted
        ws.Cell("A3").Value = true;
        await Assert.That(ws.Evaluate("GCD(A1:A3)")).IsEqualTo(XLError.IncompatibleValue);

        // Unconvertable text causes error
        ws.Cell("A3").Value = "one";
        await Assert.That(ws.Evaluate("GCD(A1:A3)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments]
    public async Task Gcd_numbers_must_fit_in_double_without_precision_loss()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("GCD(9.007E+15)")).IsEqualTo(9.007E+15);
        await Assert.That(XLWorkbook.EvaluateExpr("GCD(9.008E+15)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments]
    public async Task Gcd_numbers_must_be_zero_or_positive()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("GCD(-1)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(8.9, 8)]
    [Arguments(-8.9, -9)]
    public async Task Int(double input, double expected)
    {
        var actual = XLWorkbook.EvaluateExpr($"INT({input})");
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Arguments("24, 36", 72)]
    [Arguments("24.9, 36.9", 72)]
    [Arguments("{24, 36}", 72)]
    [Arguments("{1,2,3;4,5,6}", 60)]
    [Arguments("{\"1\",\"2\",\"3\"}", 6)]
    [Arguments("240, 360, 30", 720)]
    [Arguments("5, 0", 0)]
    [Arguments("0, 5", 0)]
    public async Task Lcm(string args, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"LCM({args})")).IsEqualTo(expected);
    }

    [Test]
    public async Task Lcm_accepts_references()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new object[]
        {
            (1, 2, 3),
            ("4", "5", "6"),
        });
        await Assert.That(ws.Evaluate("LCM(A1:B2,C1:C2)")).IsEqualTo(60);

        // Blank is considered 0
        await Assert.That(ws.Evaluate("LCM(A1:A3)")).IsEqualTo(0);

        // Logical are not converted
        ws.Cell("A3").Value = true;
        await Assert.That(ws.Evaluate("LCM(A1:A3)")).IsEqualTo(XLError.IncompatibleValue);

        // Unconvertable text causes error
        ws.Cell("A3").Value = "one";
        await Assert.That(ws.Evaluate("LCM(A1:A3)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments]
    public async Task Lcm_numbers_must_fit_in_double_without_precision_loss()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("LCM(9.007E+15)")).IsEqualTo(9.007E+15);
        await Assert.That(XLWorkbook.EvaluateExpr("LCM(9.008E+15)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments]
    public async Task Lcm_numbers_must_be_zero_or_positive()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("LCM(-1)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(86, 4.4543472962)]
    [Arguments(2.7182818, 0.9999999895)]
    [Arguments(20.085536923, 3)]
    public async Task Ln_calculates_logarithm(double x, double ln)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"LN({x})")).IsEqualTo(ln).Within(tolerance);
    }

    [Test]
    [Arguments(0)]
    [Arguments(-0.7)]
    [Arguments(-10)]
    public async Task Ln_non_positive_returns_error(double x)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"LN({x})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(10, 10, 1)]
    [Arguments(8, 2, 3)]
    [Arguments(86, 2.7182818, 4.4543473428883)]
    public async Task Log_calculates_logarithm(double x, double @base, double result)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"LOG({x}, {@base})")).IsEqualTo(result).Within(tolerance);
    }

    [Test]
    public async Task Log_default_base_is_10()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("LOG(100)")).IsEqualTo(2);
    }

    [Test]
    public async Task Log_error_conditions()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("LOG(0)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("LOG(1,0)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("LOG(0,0)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("LOG(10,1)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    [Arguments(86, 1.93449845124)]
    [Arguments(10, 1)]
    [Arguments(1E5, 5)]
    public async Task Log10_calculates_logarithm(double x, double expectedResult)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"LOG10({x})")).IsEqualTo(expectedResult).Within(tolerance);
    }

    [Test]
    [Arguments(0)]
    [Arguments(-5)]
    [Arguments(-0.5)]
    public async Task Log10_error_conditions(double x)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"LOG10({x})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Log10_is_detected_inside_expression()
    {
        // Because LOG10 is extracted from CellFunction, make sure it is properly read even in the middle of expression.
        await Assert.That(XLWorkbook.EvaluateExpr("0 + LOG10(10)")).IsEqualTo(1);
    }

    [Test]
    public async Task MDeterm()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new object[]
        {
            (2, 4),
            (3, 5),
        });

        ws.Cell("A5").FormulaA1 = "MDETERM(A1:B2)";
        var actual = ws.Cell("A5").Value;
        await Assert.That((double)actual).IsEqualTo(-2).Within(tolerance);

        ws.Cell("A6").FormulaA1 = "SUM(A5)";
        actual = ws.Cell("A6").Value;
        await Assert.That((double)actual).IsEqualTo(-2).Within(tolerance);

        ws.Cell("A7").FormulaA1 = "SUM(MDETERM(A1:B2))";
        actual = ws.Cell("A7").Value;
        await Assert.That((double)actual).IsEqualTo(-2).Within(tolerance);
    }

    [Test]
    public async Task MDeterm_examples()
    {
        // Examples from spec
        await Assert.That((double)XLWorkbook.EvaluateExpr("MDETERM({3,6,1;1,1,0;3,10,2})")).IsEqualTo(1).Within(tolerance);
        await Assert.That(XLWorkbook.EvaluateExpr("MDETERM({3,6;1,1})")).IsEqualTo(-3);

        // Example from office website
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new object[]
        {
            ("Data", "Data", "Data", "Data"),
            (1, 3, 8, 5),
            (1, 3, 6, 1),
            (1, 1, 1, 0),
            (7, 3, 10, 2),
        });
        await Assert.That((double)ws.Evaluate("MDETERM(A2:D5)")).IsEqualTo(88).Within(tolerance);
    }

    [Test]
    public async Task MDeterm_requires_equal_number_of_rows_and_columns()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("MDETERM({1,2})")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task MDeterm_singular_matrix_returns_zero()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("MDETERM({1,2;1,2})")).IsEqualTo(0);
    }

    [Test]
    public async Task MDeterm_requires_all_array_elements_are_numbers()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new object[]
        {
            (2, 4),
            (3, 5),
        });

        ws.Cell("B2").Value = Blank.Value;
        await Assert.That(ws.Evaluate("MDETERM(A1:B2)")).IsEqualTo(XLError.IncompatibleValue);

        ws.Cell("B2").Value = "1";
        await Assert.That(ws.Evaluate("MDETERM(A1:B2)")).IsEqualTo(XLError.IncompatibleValue);

        ws.Cell("B2").Value = true;
        await Assert.That(ws.Evaluate("MDETERM(A1:B2)")).IsEqualTo(XLError.IncompatibleValue);

        ws.Cell("B2").Value = XLError.NameNotRecognized;
        await Assert.That(ws.Evaluate("MDETERM(A1:B2)")).IsEqualTo(XLError.NameNotRecognized);
    }

    [Test]
    public async Task MInverse()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new[]
        {
            (1, 2, 1),
            (3, 4, -1),
            (0, 2, 0),
        });

        ws.Cell("A5").FormulaA1 = "MINVERSE(A1:C3)";
        var actual = ws.Cell("A5").Value;
        await Assert.That((double)actual).IsEqualTo(0.25).Within(tolerance);

        ws.Cell("A6").FormulaA1 = "SUM(A5)";
        actual = ws.Cell("A6").Value;
        await Assert.That((double)actual).IsEqualTo(0.25).Within(tolerance);

        ws.Cell("A7").FormulaA1 = "SUM(MINVERSE(A1:C3))";
        actual = ws.Cell("A7").Value;
        await Assert.That((double)actual).IsEqualTo(0.5).Within(tolerance);
    }

    [Test]
    public async Task MInverse_returns_error_on_singular_matrix()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new[]
        {
            (1, 2),
            (1, 2),
        });
        await Assert.That(ws.Evaluate("MINVERSE(A1:B2)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task MInverse_requires_square_matrix()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("MINVERSE({1,2,3;7,5,5})")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task MInverse_all_array_elements_must_be_numbers()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new[]
        {
            (1, 2),
            (8, 4),
        });

        ws.Cell("B2").Value = Blank.Value;
        await Assert.That(ws.Evaluate("MINVERSE(A1:B2)")).IsEqualTo(XLError.IncompatibleValue);

        ws.Cell("B2").Value = true;
        await Assert.That(ws.Evaluate("MINVERSE(A1:B2)")).IsEqualTo(XLError.IncompatibleValue);

        ws.Cell("B2").Value = "1";
        await Assert.That(ws.Evaluate("MINVERSE(A1:B2)")).IsEqualTo(XLError.IncompatibleValue);

        ws.Cell("B2").Value = XLError.DivisionByZero;
        await Assert.That(ws.Evaluate("MINVERSE(A1:B2)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    public async Task MMult()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new[]
        {
            (2, 4),
            (3, 5),
            (2, 4),
            (3, 5),
        });

        ws.Cell("A5").FormulaA1 = "MMULT(A1:B2, A3:B4)";
        var actual = ws.Cell("A5").Value;
        await Assert.That(actual).IsEqualTo(16.0);

        ws.Cell("A6").FormulaA1 = "SUM(A5)";
        actual = ws.Cell("A6").Value;
        await Assert.That(actual).IsEqualTo(16.0);

        ws.Cell("A7").FormulaA1 = "SUM(MMULT(A1:B2, A3:B4))";
        actual = ws.Cell("A7").Value;
        await Assert.That(actual).IsEqualTo(102.0);
    }

    [Test]
    public async Task MMult_handles_non_square_matrices()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new object[]
        {
            // 2x3
            (1, 3, 5),
            (2, 4, 6),
            // 3x4
            (10, 13, 16, 19),
            (11, 14, 17, 20),
            (12, 15, 18, 21),
        });

        // 2x4 output expected:
        // 103, 130, 157, 184
        // 136, 172, 208, 244
        ws.Cell("A6").FormulaA1 = "MMult(A1:C2, A3:D5)";
        var actual = ws.Cell("A6").Value;
        await Assert.That(actual).IsEqualTo(103.0);

        ws.Cell("A7").FormulaA1 = "Sum(MMult(A1:C2, A3:D5))";
        actual = ws.Cell("A7").Value;
        await Assert.That(actual).IsEqualTo(1334);
    }

    [Test]
    [Arguments("A2:C2", "A3:C3")] // 1x3 and 1x3
    [Arguments("A2:C4", "A5:C5")] // 3x3 and 1x3
    [Arguments("A2:C5", "A6:D6")] // 3x4 and 1x4
    public async Task MMult_array1_rows_must_match_array2_column(string array1Range, string array2Range)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cells($"{array1Range}").Value = 1.0;
        ws.Cells($"{array2Range}").Value = 1.0;

        ws.Cell("A1").FormulaA1 = $"MMULT({array1Range},{array2Range})";

        await Assert.That(ws.Cell("A1").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("")]
    [Arguments("Text")]
    public async Task MMult_ThrowsWhenCellsContainInvalidInput(string invalidInput)
    {
        var ws = new XLWorkbook().AddWorksheet("Sheet1");

        // 2x3
        ws.Cell("A1").SetValue(1).CellRight().SetValue(3).CellRight().SetValue(invalidInput);
        ws.Cell("A2").SetValue(2).CellRight().SetValue(4).CellRight().SetValue(6);

        // 3x4
        ws.Cell("A3").SetValue(10).CellRight().SetValue(13).CellRight().SetValue(16).CellRight().SetValue(19);
        ws.Cell("A4").SetValue(11).CellRight().SetValue(14).CellRight().SetValue(17).CellRight().SetValue(20);
        ws.Cell("A5").SetValue(12).CellRight().SetValue(15).CellRight().SetValue(18).CellRight().SetValue(21);

        ws.Cell("A6").FormulaA1 = "MMULT(A1:C2,A3:D4)";

        await Assert.That(ws.Cell("A6").Value).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments(1.5, 1, 0.5)]
    [Arguments(3, 2, 1)]
    [Arguments(-3, 2, 1)]
    [Arguments(-3, -2, -1)]
    [Arguments(-4.3, -0.5, -0.3)]
    [Arguments(6.9, -0.2, -0.1)]
    [Arguments(0.7, 0.6, 0.1)]
    [Arguments(6.2, 1.1, 0.7)]
    public async Task Mod(double x, double y, double result)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"MOD({x}, {y})");
        await Assert.That(actual).IsEqualTo(result).Within(tolerance);
    }

    [Test]
    public async Task Mod_divisor_zero_returns_error()
    {
        // Spec says that "If y is 0, the return value is unspecified", but Excel says #DIV/0!, so let's go with that.
        await Assert.That(XLWorkbook.EvaluateExpr("MOD(1, 0)")).IsEqualTo(XLError.DivisionByZero);
        await Assert.That(XLWorkbook.EvaluateExpr("MOD(0, 0)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    [Arguments(10, 3, 9.0)]
    [Arguments(10.5, 3, 12.0)]
    [Arguments(10.4, 3, 9.0)]
    [Arguments(-10, -3, -9.0)]
    [Arguments(1.3, 0.2, 1.4)]
    [Arguments(5677.912288, 10, 5680.0)]
    [Arguments(5674.912288, 10, 5670.0)]
    [Arguments(0.5, 1, 1.0)]
    [Arguments(0.49999, 1, 0.0)]
    [Arguments(0.5, 1, 1.0)]
    [Arguments(0.49999, 1, 0.0)]
    [Arguments(0.5, 1, 1.0)]
    [Arguments(0.49999, 1, 0.0)]
    [Arguments(-13.4, -3, -12.0)]
    [Arguments(-13.5, -3, -15.0)]
    [Arguments(0.9, 0.2, 1.0)]
    [Arguments(0.89999, 0.2, 0.8)]
    [Arguments(15.5, 3, 15.0)]
    [Arguments(1.4, 0.5, 1.5)]
    [Arguments(3, 7, 0)]
    [Arguments(3, 0, 0)]
    public async Task MRound(double number, double multiple, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"MROUND({number}, {multiple})")).IsEqualTo(expected).Within(1e-12);
    }

    [Test]
    [Arguments(123456.123, -10)]
    [Arguments(-123456.123, 5)]
    public async Task MRoundExceptions(double number, double multiple)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"MROUND({number}, {multiple})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Multinomial()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("MULTINOMIAL(2)")).IsEqualTo(1);
        await Assert.That(XLWorkbook.EvaluateExpr("MULTINOMIAL(2,3)")).IsEqualTo(10);
        await Assert.That(XLWorkbook.EvaluateExpr("MULTINOMIAL(2,3,4)")).IsEqualTo(1260);
        await Assert.That(XLWorkbook.EvaluateExpr("MULTINOMIAL(1E+100)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Multinomial_accepts_ranges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("B2").InsertData(MultinomialDataB2);
        ws.Cell("A5").InsertData(MultinomialDataA5);

        await Assert.That(ws.Evaluate("MULTINOMIAL(B:XFD, 2, A5:A6)")).IsEqualTo(3087564480d);
    }

    [Test]
    public async Task Multinomial_doesnt_accept_negative_values()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("MULTINOMIAL(5, -1)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Multinomial_coercion()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = true;
        ws.Cell("A2").Value = 5;
        ws.Cell("A3").Value = "1 2/2";
        ws.Cell("A4").Value = "one";

        // True is not converted
        await Assert.That(ws.Evaluate("MULTINOMIAL(A1:A2)")).IsEqualTo(XLError.IncompatibleValue);

        // Text is coerced
        await Assert.That(ws.Evaluate("MULTINOMIAL(A2:A3)")).IsEqualTo(21);

        // Text is coerced, errors are propagates
        await Assert.That(ws.Evaluate("MULTINOMIAL(A2:A4)")).IsEqualTo(XLError.IncompatibleValue);

        // Errors are propagates
        await Assert.That(ws.Evaluate("MULTINOMIAL(5, #DIV/0!)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    [Arguments(1.5, 3)]
    [Arguments(3, 3)]
    [Arguments(2, 3)]
    [Arguments(-1, -1)]
    [Arguments(-2, -3)]
    [Arguments(0, 1)]
    [Arguments(1E+100, 1E+100)]
    public async Task Odd(double number, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"ODD({number})")).IsEqualTo(expected).Within(1e-12);
    }

    [Test]
    public async Task Pi()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("PI()")).IsEqualTo(Math.PI);
    }

    [Test]
    [Arguments(2, 3, 8)]
    [Arguments(2, 0.5, 1.414213562373)]
    [Arguments(-1.234, 5.0, -2.861381721051)]
    [Arguments(1.234, 5.1, 2.9221823578798)]
    public async Task Power(double x, double y, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"POWER({x}, {y})")).IsEqualTo(expected).Within(1e-12);
    }

    [Test]
    public async Task Power_errors()
    {
        // Negative base and fractional exponent
        await Assert.That(XLWorkbook.EvaluateExpr("POWER(-5, 0.5)")).IsEqualTo(XLError.NumberInvalid);

        // Spec says this should be #DIV/0!, but Excel says #NUM!
        await Assert.That(XLWorkbook.EvaluateExpr("POWER(0, 0)")).IsEqualTo(XLError.NumberInvalid);

        // base is zero and exponent is negative -> #NUM!
        await Assert.That(XLWorkbook.EvaluateExpr("POWER(0, -5)")).IsEqualTo(XLError.DivisionByZero);

        // Result is not representable (e.g. out fo range)
        await Assert.That(XLWorkbook.EvaluateExpr("POWER(1e+100, 1e+100)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Product()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT(2,3,4)")).IsEqualTo(24d);

        // Examples from specification
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT(1)")).IsEqualTo(1d);
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT(1,2,3,4,5)")).IsEqualTo(120d);
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT({1,2;3,4})")).IsEqualTo(24d);
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT({2,3},4,\"5\")")).IsEqualTo(120d);

        // If no arguments are passed, return 0
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT({\"hello\"})")).IsEqualTo(0);

        // Scalar blank is skipped
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT(IF(TRUE,), 1)")).IsEqualTo(1);

        // Scalar logical is converted to number
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT(FALSE, 1)")).IsEqualTo(0);
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT(2, TRUE)")).IsEqualTo(2);

        // Scalar text is converted to number
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT(\"5\")")).IsEqualTo(5);

        // Scalar text that is not convertible return error
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT(1, \"Hello\")")).IsEqualTo(XLError.IncompatibleValue);

        // Array non-number arguments are ignored
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT({5, \"Hello\", FALSE, TRUE})")).IsEqualTo(5);

        // Reference argument only uses number, ignores blanks, logical and text
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").Value = true;
        ws.Cell("A3").Value = "100";
        ws.Cell("A4").Value = "hello";
        ws.Cell("A5").Value = 2;
        ws.Cell("A6").Value = 3;
        await Assert.That(ws.Evaluate("PRODUCT(A1:A6)")).IsEqualTo(6);

        // Scalar error is propagated
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT(1, #NULL!)")).IsEqualTo(XLError.NullValue);

        // Array error is propagated
        await Assert.That(XLWorkbook.EvaluateExpr("PRODUCT({1, #NULL!})")).IsEqualTo(XLError.NullValue);

        // Reference error is propagated
        ws.Cell("A1").Value = XLError.NoValueAvailable;
        await Assert.That(ws.Evaluate("PRODUCT(A1)")).IsEqualTo(XLError.NoValueAvailable);
    }

    [Test]
    [Arguments(5, 2, 2)]
    [Arguments(4.5, 3.1, 1)]
    [Arguments(-10, 3, -3)]
    [Arguments(-10, -4, 2)]
    [Arguments(1E+100, 1E+40, 1E+60)]
    public async Task Quotient(double x, double y, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"QUOTIENT({x}, {y})")).IsEqualTo(expected);
    }

    [Test]
    public async Task Quotient_errors()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("QUOTIENT(1, 0)")).IsEqualTo(XLError.DivisionByZero);
    }

    [Test]
    [Arguments(270, 4.71238898038469)]
    [Arguments(-180, -Math.PI)]
    public async Task Radians(double angle, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"RADIANS({angle})")).IsEqualTo(expected).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task Rand()
    {
        for (var i = 0; i < 100; ++i)
        {
            var randomNumber = (double)XLWorkbook.EvaluateExpr("RAND()");
            await Assert.That(randomNumber).IsGreaterThanOrEqualTo(0);
            await Assert.That(randomNumber).IsLessThan(1);
        }
    }

    [Test]
    public async Task RandBetween()
    {
        for (var i = 0; i < 100; ++i)
        {
            var randomNumber = (double)XLWorkbook.EvaluateExpr("RANDBETWEEN(10, 20)");
            await Assert.That(randomNumber).IsGreaterThanOrEqualTo(10);
            await Assert.That(randomNumber).IsLessThanOrEqualTo(20);
        }

        await Assert.That((double)XLWorkbook.EvaluateExpr("RANDBETWEEN(100.5, 100.9)")).IsEqualTo(101);
        await Assert.That(XLWorkbook.EvaluateExpr("RANDBETWEEN(100.9, 100.5)")).IsEqualTo(XLError.NumberInvalid);
        await Assert.That(XLWorkbook.EvaluateExpr("RANDBETWEEN(20, 5)")).IsEqualTo(XLError.NumberInvalid);
        var randBetweenLarge = (double)XLWorkbook.EvaluateExpr("RANDBETWEEN(1E+100, 1E+110)");
        await Assert.That(randBetweenLarge).IsGreaterThanOrEqualTo(1E+100);
        await Assert.That(randBetweenLarge).IsLessThanOrEqualTo(1E+110);
    }

    [Test]
    [Arguments(1, 0, "I")]
    [Arguments(3046, 1, "MMMVLI")]
    [Arguments(3999, 1, "MMMLMVLIV")]
    [Arguments(999, 0, "CMXCIX")]
    [Arguments(999.99, 0.9, "CMXCIX")]
    [Arguments(999, 1, "LMVLIV")]
    [Arguments(999, 2, "XMIX")]
    [Arguments(999, 3, "VMIV")]
    [Arguments(999, 4, "IM")]
    public async Task Roman(double value, double form, string expected)
    {
        await Assert.That((string)XLWorkbook.EvaluateExpr($"ROMAN({value}, {form})")).IsEqualTo(expected);
    }

    [Test]
    public async Task Roman_value_0_is_empty_string()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("ROMAN(0, 0)")).IsEqualTo(string.Empty);
    }

    [Test]
    public async Task Roman_has_optional_second_argument_with_default_value_0()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("ROMAN(999)")).IsEqualTo("CMXCIX");
    }

    [Test]
    public async Task Roman_form_must_be_between_0_and_4()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("ROMAN(1, -1)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr("ROMAN(1, 5)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task Roman_value_must_be_between_0_and_3999()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("ROMAN(-1, 0)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(XLWorkbook.EvaluateExpr("ROMAN(4000, 0)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments(2.15, 1, 2.2)]
    [Arguments(2.149, 1, 2.1)]
    [Arguments(-1.475, 2, -1.48)]
    [Arguments(21.5, -1, 20.0)]
    [Arguments(626.3, -3, 1000.0)]
    [Arguments(1.98, -1, 0.0)]
    [Arguments(-50.55, -2, -100.0)]
    [Arguments(31.565, 2, 31.57)]
    [Arguments(-31.565, 2, -31.57)]
    [Arguments(1E+100, 2, 1E+100)]
    [Arguments(1.25, 0, 1)]
    [Arguments(1, -1E+100, 0)]
    [Arguments(1.123456, 1E+100, 1.123456)] // Excel says 0 for anything over 2147483646
    public async Task Round(double number, double digits, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"ROUND({number}, {digits})")).IsEqualTo(expected);
    }

    [Test]
    [Arguments(3.2, 0, 3.0)]
    [Arguments(76.9, 0, 76.0)]
    [Arguments(3.14159, 3, 3.141)]
    [Arguments(-3.14159, 1, -3.1)]
    [Arguments(31415.92654, -2, 31400.0)]
    [Arguments(0, 3, 0)]
    public async Task RoundDown(double number, double digits, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"ROUNDDOWN({number}, {digits})")).IsEqualTo(expected);
    }

    [Test]
    [Arguments(3.2, 0, 4)]
    [Arguments(76.9, 0, 77.0)]
    [Arguments(3.14159, 3, 3.142)]
    [Arguments(-3.14159, 1, -3.2)]
    [Arguments(31415.92654, -2, 31500.0)]
    [Arguments(0, 3, 0)]
    [Arguments(11, 0, 11)]
    public async Task RoundUp(double number, double digits, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"ROUNDUP({number}, {digits})")).IsEqualTo(expected);
    }

    [Test]
    [Arguments("0", 0)]
    [Arguments("10.5", 1)]
    [Arguments("-5.4", -1)]
    [Arguments("-0.00001", -1)]
    [Arguments("-1E+300", -1)]
    [Arguments("\"0 1/2\"", 1)]
    [Arguments("FALSE", 0)]
    [Arguments("TRUE", 1)]
    public async Task Sign(string arg, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"SIGN({arg})")).IsEqualTo(expected);
    }

    [Test]
    [Arguments("0", 0)]
    [Arguments("1", 0.8414709848078965)]
    [Arguments("-1", -0.8414709848078965)]
    [Arguments("PI()", 0)]
    [Arguments("PI()/2", 1)]
    [Arguments("30*PI()/180", 0.5)]
    [Arguments("RADIANS(30)", 0.5)]
    public async Task Sin(string arg, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"SIN({arg})")).IsEqualTo(expected).Within(tolerance);
    }

    [Test]
    [Arguments("0", 0)]
    [Arguments("1", 1.1752011936438014)]
    [Arguments("10", 11013.232874703393)]
    [Arguments("100", 1.3440585709080678E+43)]
    [Arguments("100", 1.3440585709080678E+43)]
    [Arguments("711", XLError.NumberInvalid)]
    [Arguments("-711", XLError.NumberInvalid)]
    public async Task Sinh(string arg, object result)
    {
        var actual = XLWorkbook.EvaluateExpr($"SINH({arg})");
        await Assert.That(actual).IsEqualTo(ExpectedCellValue.From(result));
    }

    [Test]
    [Arguments(0, 1)]
    [Arguments(0.3, 1.0467516)]
    [Arguments(0.6, 1.21162831)]
    [Arguments(0.9, 1.60872581)]
    [Arguments(1.2, 2.759703601)]
    [Arguments(1.5, 14.1368329)]
    [Arguments(1.8, -4.401367872)]
    [Arguments(2.1, -1.980801656)]
    [Arguments(2.4, -1.356127641)]
    [Arguments(2.7, -1.10610642)]
    [Arguments(3.0, -1.010108666)]
    [Arguments(3.3, -1.012678974)]
    [Arguments(3.6, -1.115127532)]
    [Arguments(3.9, -1.377538917)]
    [Arguments(4.2, -2.039730601)]
    [Arguments(4.5, -4.743927548)]
    [Arguments(4.8, 11.42870421)]
    [Arguments(5.1, 2.645658426)]
    [Arguments(5.4, 1.575565187)]
    [Arguments(5.7, 1.198016873)]
    [Arguments(6.0, 1.041481927)]
    [Arguments(6.3, 1.000141384)]
    [Arguments(6.6, 1.052373922)]
    [Arguments(6.9, 1.225903187)]
    [Arguments(7.2, 1.643787029)]
    [Arguments(7.5, 2.884876262)]
    [Arguments(7.8, 18.53381902)]
    [Arguments(8.1, -4.106031636)]
    [Arguments(8.4, -1.925711244)]
    [Arguments(8.7, -1.335743646)]
    [Arguments(9.0, -1.097537906)]
    [Arguments(9.3, -1.007835594)]
    [Arguments(9.6, -1.015550252)]
    [Arguments(9.9, -1.124617578)]
    [Arguments(10.2, -1.400039323)]
    [Arguments(10.5, -2.102886109)]
    [Arguments(10.8, -5.145888341)]
    [Arguments(11.1, 9.593612018)]
    [Arguments(11.4, 2.541355049)]
    [Arguments(45, 1.90359)]
    [Arguments(30, 6.48292)]
    public async Task Sec_returns_correct_number(double angle, double expectedOutput)
    {
        var result = (double)XLWorkbook.EvaluateExpr($"SEC({angle})");
        await Assert.That(result).IsEqualTo(expectedOutput).Within(0.00001);

        // as the secant is symmetric for positive and negative numbers, let's assert twice:
        var resultForNegative = (double)XLWorkbook.EvaluateExpr($"SEC({-angle})");
        await Assert.That(resultForNegative).IsEqualTo(expectedOutput).Within(0.00001);
    }

    [Test]
    [Arguments(-9, 0.00024682)]
    [Arguments(-8, 0.000670925)]
    [Arguments(-7, 0.001823762)]
    [Arguments(-6, 0.004957474)]
    [Arguments(-5, 0.013475282)]
    [Arguments(-4, 0.036618993)]
    [Arguments(-3, 0.099327927)]
    [Arguments(-2, 0.265802229)]
    [Arguments(-1, 0.648054274)]
    [Arguments(0, 1)]
    [Arguments(1E+100, 0)]
    [Arguments(1E-100, 1)]
    public async Task Sech_returns_correct_number(double angle, double expectedOutput)
    {
        var result = (double)XLWorkbook.EvaluateExpr($"SECH({angle})");
        await Assert.That(result).IsEqualTo(expectedOutput).Within(0.00001);

        // as the secant is symmetric for positive and negative numbers, let's assert twice:
        var resultForNegative = (double)XLWorkbook.EvaluateExpr($"SECH({-angle})");
        await Assert.That(resultForNegative).IsEqualTo(expectedOutput).Within(0.00001);
    }

    [Test]
    public async Task SeriesSum()
    {
        await Assert.That(XLWorkbook.EvaluateExpr("SERIESSUM(2,3,4,5)")).IsEqualTo(40.0);

        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A2").FormulaA1 = "PI()/4";
        ws.Cell("A3").Value = 1;
        ws.Cell("A4").FormulaA1 = "-1/FACT(2)";
        ws.Cell("A5").FormulaA1 = "1/FACT(4)";
        ws.Cell("A6").FormulaA1 = "-1/FACT(6)";

        var actual = ws.Evaluate("SERIESSUM(A2,0,2,A3:A6)");
        await Assert.That(actual).IsEqualTo(0.70710321482284566);
    }

    [Test]
    [Arguments("{1,2,3;4,5,6}")]
    [Arguments("{1,2,3,4,5,6}")]
    [Arguments("{1,2;3,4;5,6}")]
    public async Task SeriesSum_takes_coefficients_row_by_row_left_to_right(string array)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"SERIESSUM(2,2,1,{array})")).IsEqualTo(1284);
    }

    [Test]
    public async Task SeriesSum_returns_invalid_number_error_when_result_is_too_large()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").InsertData(new object[] { 1, 2, 3, 4, 5 });
        await Assert.That(ws.Evaluate("SERIESSUM(10,100,100,A1:A3)")).IsEqualTo(3E+300);
        await Assert.That(ws.Evaluate("SERIESSUM(10,100,100,A1:A4)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task SeriesSum_coercion()
    {
        // For some weird reason, SERIESSUM doesn't convert logical
        foreach (var invalidValue in new[] { "\"\"", "TRUE" })
        {
            await Assert.That(XLWorkbook.EvaluateExpr($"SERIESSUM({invalidValue},1,1,1)")).IsEqualTo(XLError.IncompatibleValue);
            await Assert.That(XLWorkbook.EvaluateExpr($"SERIESSUM(1,{invalidValue},1,1)")).IsEqualTo(XLError.IncompatibleValue);
            await Assert.That(XLWorkbook.EvaluateExpr($"SERIESSUM(1,1,{invalidValue},1)")).IsEqualTo(XLError.IncompatibleValue);
            await Assert.That(XLWorkbook.EvaluateExpr($"SERIESSUM(1,1,1,{invalidValue})")).IsEqualTo(XLError.IncompatibleValue);
        }

        // Blank and text values are coerced to a number
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        foreach (var validArg in new[] { "A1", "\"0 0/2\"" })
        {
            await Assert.That(ws.Evaluate($"SERIESSUM({validArg},1,1,1)")).IsEqualTo(0);
            await Assert.That(ws.Evaluate($"SERIESSUM(1,{validArg},1,1)")).IsEqualTo(1);
            await Assert.That(ws.Evaluate($"SERIESSUM(1,1,{validArg},1)")).IsEqualTo(1);
        }

        // Text is not converted in an area and causes conversion error
        ws.Cell("B2").Value = "0";
        ws.Cell("B3").Value = 5;
        await Assert.That(ws.Evaluate("SERIESSUM(1,1,1,B2:B3)")).IsEqualTo(XLError.IncompatibleValue);

        // Blank is interpreted as 0
        ws.Cell("C1").Value = Blank.Value;
        ws.Cell("C2").Value = 2;
        await Assert.That(ws.Evaluate("SERIESSUM(1,1,1,C1:C2)")).IsEqualTo(2);
    }

    [Test]
    [Arguments(0, 0)]
    [Arguments(1, 1)]
    [Arguments(2, 1.4142135624)]
    [Arguments(1E+300, 1E+150)]
    public async Task Sqrt(double x, double result)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"SQRT({x})")).IsEqualTo(result).Within(tolerance);
    }

    [Test]
    [Arguments(-1)]
    [Arguments(-0.0001)]
    public async Task Sqrt_returns_invalid_number_for_negative_numbers(double x)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"SQRT({x})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task SqrtPi()
    {
        var actual = (double)XLWorkbook.EvaluateExpr("SQRTPI(1)");
        await Assert.That(actual).IsEqualTo(1.7724538509055159).Within(tolerance);

        actual = (double)XLWorkbook.EvaluateExpr("SQRTPI(2)");
        await Assert.That(actual).IsEqualTo(2.5066282746310002).Within(tolerance);

        await Assert.That(XLWorkbook.EvaluateExpr("SQRTPI(-1)")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    public async Task Subtotal()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Non-existent functions return error
        await Assert.That(ws.Evaluate("SUBTOTAL(0, A1)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(ws.Evaluate("SUBTOTAL(0.9, A1)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(ws.Evaluate("SUBTOTAL(12, A1)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(ws.Evaluate("SUBTOTAL(100.9, A1)")).IsEqualTo(XLError.IncompatibleValue);
        await Assert.That(ws.Evaluate("SUBTOTAL(112, A1)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task SubtotalAverage()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").FormulaA1 = "SUBTOTAL(1,A1,A2)";
        ws.Cell("A4").Value = "A";

        await Assert.That(ws.Cell("A3").Value).IsEqualTo(2.5);
        await Assert.That(ws.Evaluate("SUBTOTAL(1, A1:A4)")).IsEqualTo(2.5);

        ws.Row(2).Hide();
        await Assert.That(ws.Evaluate("SUBTOTAL(101, A1:A4)")).IsEqualTo(2);
    }

    [Test]
    public async Task Subtotal10Calc()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.DefinedNames.Add("subtotalrange", "$A$37:$A$38");

        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 4;
        ws.Cell("A3").FormulaA1 = "SUBTOTAL(9, A1:A2)"; // simple add subtotal
        ws.Cell("A4").Value = 8;
        ws.Cell("A5").Value = 16;
        ws.Cell("A6").FormulaA1 = "SUBTOTAL(9, A4:A5)"; // simple add subtotal
        ws.Cell("A7").Value = 32;
        ws.Cell("A8").Value = 64;
        ws.Cell("A9").FormulaA1 = "SUM(A7:A8)"; // func but not subtotal
        ws.Cell("A10").Value = 128;
        ws.Cell("A11").Value = 256;
        ws.Cell("A12").FormulaA1 = "SUBTOTAL(1, A10:A11)"; // simple avg subtotal
        ws.Cell("A13").Value = 512;
        ws.Cell("A14").FormulaA1 = "SUBTOTAL(9, A1:A13)"; // subtotals in range
        ws.Cell("A15").Value = 1024;
        ws.Cell("A16").Value = 2048;
        ws.Cell("A17").FormulaA1 = "42 + SUBTOTAL(9, A15:A16)"; // simple add subtotal in formula
        ws.Cell("A18").Value = 4096;
        ws.Cell("A19").FormulaA1 = "SUBTOTAL(9, A15:A18)"; // subtotals in range
        ws.Cell("A20").Value = 8192;
        ws.Cell("A21").Value = 16384;
        ws.Cell("A22").FormulaA1 = @"32768 * SEARCH(""SUBTOTAL(9, A1:A2)"", A28)"; // subtotal literal in formula
        ws.Cell("A23").FormulaA1 = "SUBTOTAL(9, A20:A22)"; // subtotal literal in formula in range
        ws.Cell("A24").Value = 65536;
        ws.Cell("A25").FormulaA1 = "A23"; // link to subtotal
        ws.Cell("A26").FormulaA1 = "PRODUCT(SUBTOTAL(9, A24:A25), 2)"; // subtotal as parameter in func
        ws.Cell("A27").Value = 131072;
        ws.Cell("A28").Value = "SUBTOTAL(9, A1:A2)"; // subtotal literal
        ws.Cell("A29").FormulaA1 = "SUBTOTAL(9, A27:A28)"; // subtotal literal in range
        ws.Cell("A30").FormulaA1 = "SUBTOTAL(9, A31:A32)"; // simple add subtotal backward
        ws.Cell("A31").Value = 262144;
        ws.Cell("A32").Value = 524288;
        ws.Cell("A33").FormulaA1 = "SUBTOTAL(9, A20:A32)"; // subtotals in range
        ws.Cell("A34").FormulaA1 = @"SUBTOTAL(VALUE(""9""), A1:A33, A35:A41)"; // func as parameter in subtotal and many ranges
        ws.Cell("A35").Value = 1048576;
        ws.Cell("A36").FormulaA1 = "SUBTOTAL(9, A31:A32, A35)"; // many ranges
        ws.Cell("A37").Value = 2097152;
        ws.Cell("A38").Value = 4194304;
        ws.Cell("A39").FormulaA1 = "SUBTOTAL(3*3, subtotalrange)"; // formula as parameter in subtotal and named range
        ws.Cell("A40").Value = 8388608;
        ws.Cell("A41").FormulaA1 = "PRODUCT(SUBTOTAL(A4+1, A35:A40), 2)"; // formula with link as parameter in subtotal
        ws.Cell("A42").FormulaA1 = "PRODUCT(SUBTOTAL(A4+1, A35:A40), 2) + SUBTOTAL(A4+1, A35:A40)"; // two subtotals in one formula

        await Assert.That(ws.Cell("A3").Value).IsEqualTo(6);
        await Assert.That(ws.Cell("A6").Value).IsEqualTo(24);
        await Assert.That(ws.Cell("A12").Value).IsEqualTo(192);
        await Assert.That(ws.Cell("A14").Value).IsEqualTo(1118);
        await Assert.That(ws.Cell("A17").Value).IsEqualTo(3114);
        await Assert.That(ws.Cell("A19").Value).IsEqualTo(7168);
        await Assert.That(ws.Cell("A23").Value).IsEqualTo(57344);
        await Assert.That(ws.Cell("A26").Value).IsEqualTo(245760);
        await Assert.That(ws.Cell("A29").Value).IsEqualTo(131072);
        await Assert.That(ws.Cell("A30").Value).IsEqualTo(786432);
        await Assert.That(ws.Cell("A33").Value).IsEqualTo(1097728);
        await Assert.That(ws.Cell("A34").Value).IsEqualTo(16834654);
        await Assert.That(ws.Cell("A36").Value).IsEqualTo(1835008);
        await Assert.That(ws.Cell("A39").Value).IsEqualTo(6291456);
        await Assert.That(ws.Cell("A41").Value).IsEqualTo(31457280);
        await Assert.That(ws.Cell("A42").Value).IsEqualTo(47185920);
    }

    [Test]
    public async Task Subtotal100Calc()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Cell("A1").Value = 1;
        ws.Cell("B1").Value = 2;
        ws.Cell("C1").Value = Blank.Value;
        ws.Cell("A2").Value = "A";
        ws.Cell("B2").Value = 4;
        ws.Cell("C2").Value = 8;
        ws.Cell("A3").FormulaA1 = "SUBTOTAL(109, A1:A2)";
        ws.Cell("B3").FormulaA1 = "SUBTOTAL(109, B1:B2)";
        ws.Cell("C3").FormulaA1 = "SUBTOTAL(109, C1:C2)";
        ws.Cell("A4").Value = 16;
        ws.Cell("B4").Value = 32;
        ws.Cell("C4").Value = 64;
        ws.Cell("A5").Value = 128;
        ws.Cell("B5").Value = 256;
        ws.Cell("C5").Value = 512;
        ws.Cell("A6").FormulaA1 = "SUBTOTAL(109, A1:A5)";
        ws.Cell("B6").FormulaA1 = "SUBTOTAL(109, B1:B5)";
        ws.Cell("C6").FormulaA1 = "SUBTOTAL(109, C1:C5)";

        ws.Row(2).Hide();
        ws.Row(5).Hide();

        await Assert.That(ws.Cell("A3").Value).IsEqualTo(1);
        await Assert.That(ws.Cell("B3").Value).IsEqualTo(2);
        await Assert.That(ws.Cell("C3").Value).IsEqualTo(0);
        await Assert.That(ws.Cell("A6").Value).IsEqualTo(17);
        await Assert.That(ws.Cell("B6").Value).IsEqualTo(34);
        await Assert.That(ws.Cell("C6").Value).IsEqualTo(64);
    }

    [Test]
    public async Task SubtotalCount()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(2,A1:A3)";

        await Assert.That(ws.Cell("A4").Value).IsEqualTo(2);
        await Assert.That(ws.Evaluate("SUBTOTAL(2,A2:A4)")).IsEqualTo(1);

        ws.Row(2).Hide();
        await Assert.That(ws.Evaluate("SUBTOTAL(102,A1:A4)")).IsEqualTo(1);
    }

    [Test]
    public async Task SubtotalCountA()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = string.Empty;
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(3,A1,A2,A3)";

        await Assert.That(ws.Cell("A4").Value).IsEqualTo(3);
        await Assert.That(ws.Evaluate("SUBTOTAL(3,A1:A4)")).IsEqualTo(3);

        ws.Row(1).Hide();
        await Assert.That(ws.Evaluate("SUBTOTAL(103,A1:A4)")).IsEqualTo(2);
    }

    [Test]
    public async Task SubtotalMax()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(4,A1,A2,A3) + 10";

        await Assert.That(ws.Cell("A4").Value).IsEqualTo(13);
        await Assert.That(ws.Evaluate("SUBTOTAL(4,A1:A4)")).IsEqualTo(3);

        ws.Cell("A5").Value = 2.5;
        ws.Row(2).Hide();
        await Assert.That(ws.Evaluate("SUBTOTAL(104,A1:A5)")).IsEqualTo(2.5);
    }

    [Test]
    public async Task SubtotalMin()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(5,A1,A2,A3) - 10";

        await Assert.That(ws.Cell("A4").Value).IsEqualTo(-8);
        await Assert.That(ws.Evaluate("SUBTOTAL(5,A1:A4)")).IsEqualTo(2);

        ws.Cell("A5").Value = 2.5;
        ws.Row(1).Hide();
        await Assert.That(ws.Evaluate("SUBTOTAL(105,A1:A5)")).IsEqualTo(2.5);
    }

    [Test]
    public async Task SubtotalProduct()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(6,A1,A2,A3)";

        await Assert.That(ws.Cell("A4").Value).IsEqualTo(6);
        await Assert.That(ws.Evaluate("SUBTOTAL(6,A1:A4)")).IsEqualTo(6);

        ws.Row(2).Hide();
        ws.Cell("A5").Value = 4;
        await Assert.That(ws.Evaluate("SUBTOTAL(106,A1:A5)")).IsEqualTo(8);
    }

    [Test]
    public async Task SubtotalStDev()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(7,A1:A3,A5)";
        ws.Cell("A5").Value = 5;

        await Assert.That((double)ws.Cell("A4").Value).IsEqualTo(1.5275252316).Within(XLHelper.Epsilon);
        await Assert.That((double)ws.Evaluate("SUBTOTAL(7,A1:A5)")).IsEqualTo(1.5275252316).Within(XLHelper.Epsilon);

        ws.Row(2).Hide();
        await Assert.That((double)ws.Evaluate("SUBTOTAL(107,A1:A5)")).IsEqualTo(2.1213203435).Within(XLHelper.Epsilon);
    }

    [Test]
    public async Task SubtotalStDevP()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(8,A1,A2,A3)";

        await Assert.That(ws.Cell("A4").Value).IsEqualTo(0.5);
        await Assert.That(ws.Evaluate("SUBTOTAL(8,A1:A4)")).IsEqualTo(0.5);

        ws.Row(2).Hide();
        ws.Cell("A5").Value = 3;
        await Assert.That(ws.Evaluate("SUBTOTAL(108,A1:A5)")).IsEqualTo(0.5);
    }

    [Test]
    public async Task SubtotalSum()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(9,A1,A2,A3)";

        await Assert.That(ws.Cell("A4").Value).IsEqualTo(5);
        await Assert.That(ws.Evaluate("SUBTOTAL(9,A1:A4)")).IsEqualTo(5);

        ws.Row(2).Hide();

        await Assert.That(ws.Evaluate("SUBTOTAL(109, A1:A4)")).IsEqualTo(2);
    }

    [Test]
    public async Task SubtotalVar()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 5;
        ws.Cell("A2").Value = 4;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").Value = 8;
        ws.Cell("A5").Value = 5;
        ws.Cell("A6").FormulaA1 = "SUBTOTAL(10,A1:A5)";

        await Assert.That(ws.Cell("A6").Value).IsEqualTo(3);
        await Assert.That(ws.Evaluate("SUBTOTAL(10,A1:A6)")).IsEqualTo(3);

        ws.Row(1).Hide();
        ws.Row(5).Hide();
        await Assert.That(ws.Evaluate("SUBTOTAL(110,A1:A6)")).IsEqualTo(8);
    }

    [Test]
    public async Task SubtotalVarP()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 2;
        ws.Cell("A2").Value = 3;
        ws.Cell("A3").Value = "A";
        ws.Cell("A4").FormulaA1 = "SUBTOTAL(11,A1,A2,A3)";

        await Assert.That(ws.Cell("A4").Value).IsEqualTo(0.25);
        await Assert.That(ws.Evaluate("SUBTOTAL(11,A1:A4)")).IsEqualTo(0.25);

        ws.Row(2).Hide();
        ws.Cell("A5").Value = 4;
        await Assert.That(ws.Evaluate("SUBTOTAL(111,A1:A5)")).IsEqualTo(1);
    }

    [Test]
    public async Task Sum()
    {
        var cell = new XLWorkbook().AddWorksheet("Sheet1").FirstCell();
        var fCell = cell.SetValue(1).CellBelow().SetValue(2).CellBelow();
        fCell.FormulaA1 = "sum(A1:A2)";

        await Assert.That(fCell.Value).IsEqualTo(3.0);
    }

    [Test]
    public async Task SumDateTimeAndNumber()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = new DateTime(2018, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        await Assert.That(ws.Evaluate("SUM(A1:A2)")).IsEqualTo(43102);

        ws.Cell("A1").Value = 2;
        ws.Cell("A2").FormulaA1 = "DATE(2018,1,1)";
        await Assert.That(ws.Evaluate("SUM(A1:A2)")).IsEqualTo(43103);
    }

    [Test]
    [Arguments(9, "SUMIF(A:B, \"A*\", C:C)")]
    [Arguments(9, "SUMIF(A1:B6, \"A*\", C1:C6)")]
    public async Task SumIf_InputRangeHasMultipleColumns(int expectedOutcome, string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        var data = new object[]
        {
            new { Id = "AA", Id2 = "BA", Value = 2},
            new { Id = "AB", Id2 = "BB", Value = 3},
            new { Id = "BA", Id2 = "AA", Value = 2},
            new { Id = "BB", Id2 = "AB", Value = 1},
            new { Id = "AC", Id2 = "AC", Value = 4},
        };
        ws.Cell("A1").InsertTable(data);

        await Assert.That(ws.Evaluate(formula)).IsEqualTo(expectedOutcome);
    }

    /// <summary>
    /// refers to Example 1 from the Excel documentation,
    /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
    /// </summary>
    /// <param name="expectedOutcome"></param>
    /// <param name="formula"></param>
    [Test]
    [Arguments(63000, "SUMIF(A1:A4,\">160000\", B1:B4)")]
    [Arguments(900000, "SUMIF(A1:A4,\">160000\")")]
    [Arguments(21000, "SUMIF(A1:A4, 300000, B1:B4)")]
    [Arguments(28000, "SUMIF(A1:A4, \">\" &C1, B1:B4)")]
    public async Task SumIf_ReturnsCorrectValues_ReferenceExample1FromMicrosoft(int expectedOutcome, string formula)
    {
        using var wb = new XLWorkbook();
        wb.ReferenceStyle = XLReferenceStyle.A1;

        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell(1, 1).Value = 100000;
        ws.Cell(1, 2).Value = 7000;
        ws.Cell(2, 1).Value = 200000;
        ws.Cell(2, 2).Value = 14000;
        ws.Cell(3, 1).Value = 300000;
        ws.Cell(3, 2).Value = 21000;
        ws.Cell(4, 1).Value = 400000;
        ws.Cell(4, 2).Value = 28000;

        ws.Cell(1, 3).Value = 300000;

        await Assert.That((double)ws.Evaluate(formula)).IsEqualTo(expectedOutcome);
    }

    /// <summary>
    /// refers to Example 2 from the Excel documentation,
    /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
    /// </summary>
    /// <param name="expectedOutcome"></param>
    /// <param name="formula"></param>
    [Test]
    [Arguments(2000, "SUMIF(A2:A7,\"Fruits\", C2:C7)")]
    [Arguments(12000, "SUMIF(A2:A7,\"Vegetables\", C2:C7)")]
    [Arguments(4300, "SUMIF(B2:B7, \"*es\", C2:C7)")]
    [Arguments(400, "SUMIF(A2:A7, \"\", C2:C7)")]
    public async Task SumIf_ReturnsCorrectValues_ReferenceExample2FromMicrosoft(int expectedOutcome, string formula)
    {
        using var wb = new XLWorkbook();
        wb.ReferenceStyle = XLReferenceStyle.A1;

        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell(2, 1).Value = "Vegetables";
        ws.Cell(3, 1).Value = "Vegetables";
        ws.Cell(4, 1).Value = "Fruits";
        ws.Cell(5, 1).Value = "";
        ws.Cell(6, 1).Value = "Vegetables";
        ws.Cell(7, 1).Value = "Fruits";

        ws.Cell(2, 2).Value = "Tomatoes";
        ws.Cell(3, 2).Value = "Celery";
        ws.Cell(4, 2).Value = "Oranges";
        ws.Cell(5, 2).Value = "Butter";
        ws.Cell(6, 2).Value = "Carrots";
        ws.Cell(7, 2).Value = "Apples";

        ws.Cell(2, 3).Value = 2300;
        ws.Cell(3, 3).Value = 5500;
        ws.Cell(4, 3).Value = 800;
        ws.Cell(5, 3).Value = 400;
        ws.Cell(6, 3).Value = 4200;
        ws.Cell(7, 3).Value = 1200;

        ws.Cell(1, 3).Value = 300000;

        await Assert.That((double)ws.Evaluate(formula)).IsEqualTo(expectedOutcome);
    }

    [Test]
    public async Task SumIf_ReturnsCorrectValues_WhenCalledOnFullColumn()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        var data = new object[]
        {
            new { Id = "A", Value = 2},
            new { Id = "B", Value = 3},
            new { Id = "C", Value = 2},
            new { Id = "A", Value = 1},
            new { Id = "B", Value = 4}
        };
        ws.Cell("A1").InsertTable(data);
        var formula = "=SUMIF(A:A,\"=A\",B:B)";
        var value = ws.Evaluate(formula);
        await Assert.That(value).IsEqualTo(3);
    }

    [Test]
    public async Task SumIf_ReturnsCorrectValues_WhenFormulaBelongToSameRange()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        var data = new object[]
        {
            new { Id = "A", Value = 2},
            new { Id = "B", Value = 3},
            new { Id = "C", Value = 2},
            new { Id = "A", Value = 1},
            new { Id = "B", Value = 4},
        };
        ws.Cell("A1").InsertTable(data);
        ws.Cell("A7").SetValue("Sum A");
        // SUMIF formula
        var formula = "=SUMIF(A:A,\"=A\",B:B)";
        ws.Cell("B7").SetFormulaA1(formula);
        var value = ws.Cell("B7").Value;
        await Assert.That(value).IsEqualTo(3);
    }

    [Test]
    public async Task SumIfs_MultidimensionalRanges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.FirstCell().InsertData(new object[]
        {
            (10, 10, 1, 2),
            (20, 15, 2, 4),
            (30, 20, 3, 6),
            (40, 25, 4, 8),
            (50, 30, 5, 10),
        });
        await Assert.That(ws.Evaluate("SUMIFS(C1:D5,A1:B5,\">20\")")).IsEqualTo(30);
    }

    /// <summary>
    /// refers to Example 2 to SumIf from the Excel documentation.
    /// As SumIfs should behave the same if called with three parameters, we can take that example here again.
    /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
    /// </summary>
    /// <param name="expectedResult"></param>
    /// <param name="formula"></param>
    [Test]
    [Arguments(2000, "SUMIFS(C2:C7, A2:A7, \"Fruits\")")]
    [Arguments(12000, "SUMIFS(C2:C7, A2:A7, \"Vegetables\")")]
    [Arguments(4300, "SUMIFS(C2:C7, B2:B7, \"*es\")")]
    [Arguments(400, "SUMIFS(C2:C7, A2:A7, \"\")")]
    public async Task SumIfs_ReturnsCorrectValues_ReferenceExample2FromMicrosoft(int expectedResult, string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell(2, 1).Value = "Vegetables";
        ws.Cell(3, 1).Value = "Vegetables";
        ws.Cell(4, 1).Value = "Fruits";
        ws.Cell(5, 1).Value = "";
        ws.Cell(6, 1).Value = "Vegetables";
        ws.Cell(7, 1).Value = "Fruits";

        ws.Cell(2, 2).Value = "Tomatoes";
        ws.Cell(3, 2).Value = "Celery";
        ws.Cell(4, 2).Value = "Oranges";
        ws.Cell(5, 2).Value = "Butter";
        ws.Cell(6, 2).Value = "Carrots";
        ws.Cell(7, 2).Value = "Apples";

        ws.Cell(2, 3).Value = 2300;
        ws.Cell(3, 3).Value = 5500;
        ws.Cell(4, 3).Value = 800;
        ws.Cell(5, 3).Value = 400;
        ws.Cell(6, 3).Value = 4200;
        ws.Cell(7, 3).Value = 1200;

        ws.Cell(1, 3).Value = 300000;

        var actualResult = (double)ws.Evaluate(formula);
        await Assert.That(actualResult).IsEqualTo(expectedResult);
    }

    /// <summary>
    /// refers to Example 1 to SumIf from the Excel documentation.
    /// As SumIfs should behave the same if called with three parameters, but in a different order
    /// <see cref="https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
    /// </summary>
    /// <param name="expectedOutcome"></param>
    /// <param name="formula"></param>
    [Test]
    [Arguments(63000, "SUMIFS(B1:B4, A1:A4, \">160000\")")]
    [Arguments(21000, "SUMIFS(B1:B4, A1:A4, 300000)")]
    [Arguments(28000, "SUMIFS(B1:B4, A1:A4, \">\" &C1)")]
    public async Task SumIfs_ReturnsCorrectValues_ReferenceExampleForSumIf1FromMicrosoft(int expectedOutcome, string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell(1, 1).Value = 100000;
        ws.Cell(1, 2).Value = 7000;
        ws.Cell(2, 1).Value = 200000;
        ws.Cell(2, 2).Value = 14000;
        ws.Cell(3, 1).Value = 300000;
        ws.Cell(3, 2).Value = 21000;
        ws.Cell(4, 1).Value = 400000;
        ws.Cell(4, 2).Value = 28000;

        ws.Cell(1, 3).Value = 300000;

        await Assert.That((double)ws.Evaluate(formula)).IsEqualTo(expectedOutcome);
    }

    /// <summary>
    /// refers to example data and formula to SumIfs in the Excel documentation,
    /// <see cref="https://support.office.com/en-us/article/SUMIFS-function-c9e748f5-7ea7-455d-9406-611cebce642b?ui=en-US&amp;rs=en-US&amp;ad=US"/>
    /// </summary>
    [Test]
    [Arguments(20, "=SUMIFS(A2:A9, B2:B9, \"=A*\", C2:C9, \"Tom\")")]
    [Arguments(30, "=SUMIFS(A2:A9, B2:B9, \"<>Bananas\", C2:C9, \"Tom\")")]
    public async Task SumIfs_ReturnsCorrectValues_ReferenceExampleFromMicrosoft(
        int expectedResult,
        string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        var row = 2;

        ws.Cell(row, 1).Value = 5;
        ws.Cell(row, 2).Value = "Apples";
        ws.Cell(row, 3).Value = "Tom";
        row++;

        ws.Cell(row, 1).Value = 4;
        ws.Cell(row, 2).Value = "Apples";
        ws.Cell(row, 3).Value = "Sarah";
        row++;

        ws.Cell(row, 1).Value = 15;
        ws.Cell(row, 2).Value = "Artichokes";
        ws.Cell(row, 3).Value = "Tom";
        row++;

        ws.Cell(row, 1).Value = 3;
        ws.Cell(row, 2).Value = "Artichokes";
        ws.Cell(row, 3).Value = "Sarah";
        row++;

        ws.Cell(row, 1).Value = 22;
        ws.Cell(row, 2).Value = "Bananas";
        ws.Cell(row, 3).Value = "Tom";
        row++;

        ws.Cell(row, 1).Value = 12;
        ws.Cell(row, 2).Value = "Bananas";
        ws.Cell(row, 3).Value = "Sarah";
        row++;

        ws.Cell(row, 1).Value = 10;
        ws.Cell(row, 2).Value = "Carrots";
        ws.Cell(row, 3).Value = "Tom";
        row++;

        ws.Cell(row, 1).Value = 33;
        ws.Cell(row, 2).Value = "Carrots";
        ws.Cell(row, 3).Value = "Sarah";

        var actualResult = ws.Evaluate(formula);

        await Assert.That((double)actualResult).IsEqualTo(expectedResult).Within(tolerance);
    }

    [Test]
    [Arguments("SUMIFS(D1:E5,A1:B5,\"A*\",C1:C5,\">2\")")]
    [Arguments("SUMIFS(H1:I3,A1:B3,1,D1:F2,2)")]
    [Arguments("SUMIFS(D:E,A:B,\"A*\",C:C,\">2\")")]
    public async Task SumIfs_ReturnsErrorWhenRangeDimensionsAreNotSame(string formula)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        await Assert.That(ws.Evaluate(formula)).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    [Arguments("SUMIFS(A1:A3, B1:B3,\"<>B\")", 11)]
    [Arguments("SUMIFS(A1:A3, B1:B3,\"<>\")", 110)]
    public async Task SumIfs_matches_blank_cells_when_criteria_is_not_equal(string formula, double expectedSum)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = 1;
        ws.Cell("A2").Value = 10;
        ws.Cell("A3").Value = 100;
        ws.Cell("B1").Value = Blank.Value;
        ws.Cell("B2").Value = string.Empty;
        ws.Cell("B3").Value = "B";

        await Assert.That(ws.Evaluate(formula)).IsEqualTo(expectedSum);
    }

    [Test]
    public async Task SumProduct()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.FirstCell().InsertData(Enumerable.Range(1, 10));
        ws.FirstCell().CellRight().InsertData(Enumerable.Range(1, 10).Reverse());

        await Assert.That(ws.Evaluate("SUMPRODUCT(A2)")).IsEqualTo(2);
        await Assert.That(ws.Evaluate("SUMPRODUCT(A1:A10)")).IsEqualTo(55);
        await Assert.That(ws.Evaluate("SUMPRODUCT(A1:A10, B1:B10)")).IsEqualTo(220);

        await Assert.That(ws.Evaluate("SUMPRODUCT(A1:A10, B1:B5)")).IsEqualTo(XLError.IncompatibleValue);

        // Scalar, one element array and single cell area are compatible
        await Assert.That(ws.Evaluate("SUMPRODUCT(A5, 4, {3})")).IsEqualTo(60);

        // An array can be an argument
        await Assert.That(ws.Evaluate("SUMPRODUCT(A1:A3, {3;2;1})")).IsEqualTo(10);

        // An array must have correct orientation, otherwise dimensions don't match
        await Assert.That(ws.Evaluate("SUMPRODUCT(A1:A3, {3,2,1})")).IsEqualTo(XLError.IncompatibleValue);

        // Anything but number is counted as zero. The second array is zero for all values = result is 0.
        await Assert.That(ws.Evaluate("SUMPRODUCT({1,2,3,4}, {TRUE,FALSE,\"1\",\"\"})")).IsEqualTo(0);

        // Any error returns error
        await Assert.That(ws.Evaluate("SUMPRODUCT({1,2}, {1,#N/A})")).IsEqualTo(XLError.NoValueAvailable);
        await Assert.That(ws.Evaluate("SUMPRODUCT(A1, #N/A)")).IsEqualTo(XLError.NoValueAvailable);
        ws.Cell("A2").Value = XLError.NoValueAvailable;
        await Assert.That(ws.Evaluate("SUMPRODUCT(A2, 5)")).IsEqualTo(XLError.NoValueAvailable);

        // Blank cells and cells with text should be treated as zeros
        ws.Range("A1:A5").Clear();
        await Assert.That(ws.Evaluate("SUMPRODUCT(A1:A10, B1:B10)")).IsEqualTo(110);

        // Non-number values are treated as zero
        ws.Range("A1:A5").SetValue("asdf");
        await Assert.That(ws.Evaluate("SUMPRODUCT(A1:A10, B1:B10)")).IsEqualTo(110);

        // Blank cell is considered as a blank and cause #VALUE! error
        await Assert.That(ws.Evaluate("SUMPRODUCT(Z1, 5)")).IsEqualTo(XLError.IncompatibleValue);

        // Blank value will cause #VALUE! error
        await Assert.That(ws.Evaluate("SUMPRODUCT(IF(TRUE,,), 5)")).IsEqualTo(XLError.IncompatibleValue);
    }

    [Test]
    public async Task SumSq()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // Examples from specification
        await Assert.That(XLWorkbook.EvaluateExpr("SUMSQ(2)")).IsEqualTo(4.0);
        await Assert.That(XLWorkbook.EvaluateExpr("SUMSQ(2.5, -3.6)")).IsEqualTo(19.21);
        await Assert.That(XLWorkbook.EvaluateExpr("SUMSQ({ 2.5, -3.6}, 2.4)")).IsEqualTo(24.97);

        // Scalar blank is converted to 0
        await Assert.That(XLWorkbook.EvaluateExpr("SUMSQ(IF(TRUE,), 4)")).IsEqualTo(16);

        // Scalar logical is converted to number
        await Assert.That(XLWorkbook.EvaluateExpr("SUMSQ(3, TRUE)")).IsEqualTo(10);

        // Scalar text is converted to number
        await Assert.That(XLWorkbook.EvaluateExpr("SUMSQ(\"4\", \"3\")")).IsEqualTo(25);

        // Scalar text that is not convertible return error
        await Assert.That(XLWorkbook.EvaluateExpr("SUMSQ(1, \"Hello\")")).IsEqualTo(XLError.IncompatibleValue);

        // Array logical arguments are ignored
        await Assert.That(XLWorkbook.EvaluateExpr("SUMSQ({2,TRUE,TRUE,FALSE,FALSE})")).IsEqualTo(4);

        // Array text arguments are ignored
        await Assert.That(XLWorkbook.EvaluateExpr("SUMSQ({4, 2, \"hello\", \"10\" })")).IsEqualTo(20);

        // Blank, logical and text from reference are ignored
        ws.Cell("A1").Value = Blank.Value;
        ws.Cell("A2").Value = true;
        ws.Cell("A3").Value = "100";
        ws.Cell("A4").Value = "hello";
        ws.Cell("A5").Value = 1;
        ws.Cell("A6").Value = 4;
        await Assert.That(ws.Evaluate("SUMSQ(A1:A6)")).IsEqualTo(17);

        // Scalar error is propagated
        await Assert.That(XLWorkbook.EvaluateExpr("SUMSQ(1, #NULL!)")).IsEqualTo(XLError.NullValue);

        // Array error is propagated
        await Assert.That(XLWorkbook.EvaluateExpr("SUMSQ({1, #NULL!})")).IsEqualTo(XLError.NullValue);

        // Reference error is propagated
        ws.Cell("A1").Value = XLError.NoValueAvailable;
        await Assert.That(ws.Evaluate("SUMSQ(A1)")).IsEqualTo(XLError.NoValueAvailable);
    }

    [Test]
    [Arguments(-1, -1.5574077247)]
    [Arguments(0, 0)]
    [Arguments(1, 1.5574077247)]
    [Arguments(134217727, 3.2584564256)]
    [Arguments(-134217727, -3.2584564256)]
    public async Task Tan(double radians, double expected)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"TAN({radians})")).IsEqualTo(expected).Within(tolerance);
    }

    [Test]
    [Arguments(134217728)]
    [Arguments(-134217728)]
    [Arguments(1E+100)]
    public async Task Tan_returns_invalid_number_for_radians_outside_limit(double radians)
    {
        await Assert.That(XLWorkbook.EvaluateExpr($"TAN({radians})")).IsEqualTo(XLError.NumberInvalid);
    }

    [Test]
    [Arguments(-1, -0.761594156)]
    [Arguments(0, 0)]
    [Arguments(1, 0.761594156)]
    [Arguments(1E+300, 1)]
    [Arguments(-1E+300, -1)]
    public async Task Tanh(double number, double result)
    {
        await Assert.That((double)XLWorkbook.EvaluateExpr($"TANH({number})")).IsEqualTo(result).Within(tolerance);
    }

    [Test]
    [Arguments(27.64799257, null, 27)]
    [Arguments(0, null, 0)]
    [Arguments(0, 0, 0)]
    [Arguments(3.1415926, 0, 3)]
    [Arguments(3.1415926, 1, 3.1)]
    [Arguments(3.1415926, 3, 3.141)]
    [Arguments(3.1415926, 5, 3.14159)]
    [Arguments(-4.3, 0, -4)]
    [Arguments(8.9, null, 8)]
    [Arguments(-8.9, null, -8)]
    [Arguments(0.45, null, 0)]
    public async Task Trunc(double number, double? digits, object expectedResult)
    {
        var formula = digits is null ? $"TRUNC({number})" : $"TRUNC({number}, {digits})";
        await Assert.That((double)XLWorkbook.EvaluateExpr(formula)).IsEqualTo(ExpectedCellValue.From(expectedResult));
    }

    [Test]
    [Arguments(27.64799257, -1, 20)]
    [Arguments(27.64799257, 0, 27)]
    [Arguments(27.64799257, 1, 27.6)]
    [Arguments(27.64799257, 4, 27.6479)]
    public async Task Trunc_Specify_Digits(double input, int digits, double expectedResult)
    {
        var actual = (double)XLWorkbook.EvaluateExpr($"TRUNC({input.ToString(CultureInfo.InvariantCulture)}, {digits})");
        await Assert.That(actual).IsEqualTo(expectedResult);
    }
}
