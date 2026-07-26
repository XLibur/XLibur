using System;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

[SetCulture("en-US")]
public class TextToNumberCoercionTests
{
    private const double Tolerance = 0.000001;

    [Test]
    public async Task TimeSpan_MaximumResolutionIsOneMs()
    {
        var firstValue = (double)XLWorkbook.EvaluateExpr("\"0:0:0.0015\" * 1");
        var secondValue = (double)XLWorkbook.EvaluateExpr("\"0:0:0.0024\" * 1");
        await Assert.That(secondValue).IsEqualTo(firstValue);
    }

    [Test]
    [Arguments("100%", 1)]
    [Arguments("-100%", -1)]
    [Arguments("200%", 2)]
    [Arguments("0000%", 0)]
    [Arguments("1%", 0.01)]
    [Arguments("+1%", 0.01)]
    [Arguments(" -75 % ", -0.75)]
    [Arguments(" - 100 % ", -1, Skip = ".NET parser doesn't allow whitespace between sign and number.")]
    public async Task Percent_Format9(string percent, double? expectedValue) // Format 9 '0%'
    {
        await AssertCoercion(percent, expectedValue);
    }

    [Test]
    [Arguments("100.5%", 1.005)]
    [Arguments("100 . 5%", null)]
    [Arguments(" - 100.59 % ", -1.0059, Skip = ".NET parser doesn't allow whitespace between sign and number.")]
    [Arguments("0.123456%", 0.00123456)]
    [Arguments(".5%", 0.005)]
    [Arguments("  -.375 % ", -0.00375)]
    [Arguments("100.%", 1)]
    public async Task Percent_Format10(string percent, double? expectedValue) // Format 10 '0.00%'
    {
        await AssertCoercion(percent, expectedValue);
    }

    [Test]
    [Arguments("(100%)", -1, Skip = ".NET parser doesn't parse percents.")]
    [Arguments("(-100%)", null)] // Can't have minus sign inside the brackets
    [Arguments("-(100%)", null)] // Can't have minus sign outside the brackets
    [Arguments("1,000.00%", 10)]
    [Arguments("(1,000.00%)", -10, Skip = ".NET parser doesn't parse percents.")]
    [Arguments(" % 100", 1)] // Percents can be at start or end, position doesn't matter
    public async Task Percent_UnlistedFormats(string percent, double? expectedValue) //
    {
        await AssertCoercion(percent, expectedValue);
    }

    [Test]
    [Arguments("0 1/2", 0.5)]
    [Arguments("0 /20", null)]
    [Arguments("0 1/32768", null)] // Denominator can be at most 2^15-1
    [Arguments("0 1/32767", 3.0518509475997192E-05d)]
    [Arguments("0 32768/1", null)] // Nominator can be at most 2^15-1
    [Arguments("0 32767/1", 32767)]
    [Arguments("1 32767/032767", null)] // Fraction can be only 5 digits at most
    [Arguments("1 00100/025", 5)]
    [Arguments("1 100/-2", null)] // Fractions can't be negative
    [Arguments("1 -1/2", null)]
    [Arguments("- 1 1/2", -1.5)] // can use minus sign
    [Arguments("+1 1/2", 1.5)] // or plus sign
    [Arguments("1.5 1/2", null)] // Can't use dot in whole part
    [Arguments("   1 10/20  ", 1.5)]
    [Arguments("1  1/2", null)] // Between whole part and nominator must be exactly one space
    [Arguments("1 1 /2", null)] // Can't have spaces between nominator and denominator
    [Arguments("1 1/ 2", null)]
    [Arguments("1	1/2", null)] // Tab and other whitespaces aren't allowed
    [Arguments("0 1/0", null)] // Division by zero
    public async Task Fraction_Format12_13(string fraction, double? expectedValue) // Format 12+13 '# ??/??' and  '# ?/?'
    {
        await AssertCoercion(fraction, expectedValue);
    }

    [Test]
    [Arguments("02/28/20", 43889)]
    [Arguments("002/28/20", null)]
    [Arguments("02/028/20", null)]
    [Arguments("02/28/022", null)]
    public async Task Date_Format14(string date, double? expectedValue) // Format 14 is taken from region setting, but for en (and MS errata) says 'm/d/yyyy'
    {
        await AssertCoercion(date, expectedValue);
    }

    [Test]
    [Arguments("30-apr-2000", 36646)]
    [Arguments("30-apr-20", 43951)] // 2020-04-30
    [Arguments("31-dec-9999", 2958465)]
    [Arguments("1-jan-10000", null)]
    [Arguments("1 - jan - 2022  ", 44562)] // Can have whitespace in the date
    [Arguments(" 1-jan-2022", null, Skip = ".NET parser doesn't respect the whitespace styles of a date during parsing.")] // Can't have whitespaces at the start
    [Arguments("31-dec-1899", null)] // Check 1900 "leap" year issue...
    [Arguments("1-jan-1900", 1)]
    [Arguments("28-feb-1900", 59)]
    [Arguments("1-mar-1900", 61)]
    public async Task Date_Format15(string date, double? expectedValue) // Format 15 d-mmm-yy
    {
        await AssertCoercion(date, expectedValue);
    }

    [Test]
    [Arguments("0-mar", null)] // Zero day not accepted
    [Arguments("1-mar", 44621)]
    [Arguments("1-marc", 44621, Skip = ".NET parser recognizes only abbreviation or full name of a month.")]
    [Arguments("1-march", 44621)]
    [Arguments(" 1 - apr  ", 44652)] // Unlike many others, this format also allows space at the start, not just inside and at the end
    [Arguments("31-apr", null)] // April has only 30 days
    public async Task Date_Format16(string text, double? expectedValue) // Format 16 'd-mmm'
    {
        if (expectedValue is not null)
        {
            var date = DateTime.FromOADate(expectedValue.Value);
            expectedValue = new DateTime(DateTime.Now.Year, date.Month, date.Day, 0, 0, 0, DateTimeKind.Unspecified).ToOADate();
        }

        await AssertCoercion(text, expectedValue);
    }

    [Test]
    [SetCulture("cs-CZ")]
    [Arguments("3-leden", 36528)] // Serial datetime is for 03-01-2000
    [Arguments("3-led", 36528)] // Serial datetime is for 03-01-2000
    public async Task Date_Format16_UsesCulture(string text, double? expectedValue) // Format 16 'd-mmm'
    {
        expectedValue += new DateTime(DateTime.Now.Year, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).ToOADate() - new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).ToOADate();
        await AssertCoercion(text, expectedValue);
    }

    // In en locale, there should be an extra pattern MMM-dd that is before the standard MMM-yy, but .NET Framework doesn't have it.
    // To overcome missing locale, use numbers over 31 for year (otherwise they should be interpreted as days)
    [Test]
    [Arguments("jan-02", 44563, Skip = ".NET misses culture, en interprets it as MMM-dd, but czech as MMM-yy, so the MMM-dd is the extra culture for en.")] // interpreted as 2022-01-02
    [Arguments("jan-31", 44592, Skip = "Missing excel culture mapping")] // 2022-01-02
    [Arguments("jan-32", 11689)] // 1932-01-01
    [Arguments("feb-29", 47150, Skip = "Missing excel culture mapping")] // 2029-02-01
    [Arguments("feb-30", 10990)] // 1930-02-01
    [Arguments("feb-31", 11355)] // 1931-02-01
    [Arguments("feb-003", null)] // three digits not allowed
    [Arguments("aug   -   55", 20302)] // spaces are allowed inside the pattern
    [Arguments(" aug-55", null, Skip = ".NET allow whitespaces even without specified DateTimeStyle.AllowLeadingWhite")] // starting spaces not allowed
    [Arguments("aug-55 ", 20302)] // trailing spaces allowed
    [Arguments("MaR-42", 15401)] // case insensitive
    [Arguments("marc-2", 44622, Skip = ".NET parser recognizes only abbreviation or full name of a month.")] // name can be more than three long abbr
    [Arguments("march-55", 20149)]
    [Arguments("ma-2", null)] // Name of month must be at least three chars long
    public async Task Date_Format17(string text, double? expectedValue) // Format 17 'mmm-yy'
    {
        await AssertCoercion(text, expectedValue);
    }

    [Test]
    [Arguments("1:20 AM", 0.055555555555555552d)]
    [Arguments("1:20 aM", 0.055555555555555552d)]
    [Arguments("1:60 AM", null)] // Minutes must be 0-59 range
    [Arguments("1:59 AM", 0.082638888888888887d)]
    [Arguments("13:00 AM", null)] // AM only allows hours in 0-12 range
    [Arguments("7:30 A", 0.3125)] // only starting letter of AM
    [Arguments("1:9 AM", 0.04791666666666667d)] // Single digit minutes
    public async Task Date_Format18(string text, double? expectedValue) // Format 18 'h:mm AM/PM'
    {
        await AssertCoercion(text, expectedValue);
    }

    [Test]
    [Arguments("12:0:0 PM", 0.5)]
    [Arguments("12:0:18 aM", 0.00020833333333333335d)] // case insensitive AM designator
    [Arguments("13:0:0 PM", null)] // hours can't be outside of 0-12, unlike other format
    [Arguments("13:0:0 AM", null)]
    [Arguments("00:60:00 AM", null)] // minutes can't be outside of 0-59, unlike other format
    [Arguments("00:59:00 AM", 0.040972222222222222d)]
    [Arguments("00:00:60 AM", null)] // seconds can't be outside of 0-59, unlike other format
    [Arguments("00:00:59 AM", 0.00068287037037037036d)]
    [Arguments("00:00: AM", null)] // can't omit second part (differs from time span).
    [Arguments("1:2:3 AM", 0.043090277777777776d)]
    public async Task Date_Format19(string text, double? expectedValue) // Format 19 'h:mm:ss AM/PM'
    {
        await AssertCoercion(text, expectedValue);
    }

    [Test]
    [Arguments("2/5/2022 0:0", 44597)]
    [Arguments("05/5/2022 0:0", 44686)] // Extra zero padding allowed
    [Arguments("005/5/2022 0:0", null)] // 0 prefix requires at most 2 digits
    [Arguments("13/5/2022 0:0", null)] // Month outside of range
    [Arguments("11/030/2022 0:0", null)]
    [Arguments("11/30/02022 0:0", null)] // Extra zero before year not allowed
    [Arguments("11/30/2022 24:59", 44896.04097, Skip = "Excel can have out of range parts, but .NET parsers can't.")]
    [Arguments("11/30/2022 24:60", null)] // Both parts are out of range
    [Arguments("11/30/2022 23:160", 44896.06944, Skip = "Excel can have one of of range part, but .NET parser can't.")]
    [Arguments("11/30/2022 9999:59", 45311.66597, Skip = "Excel parser accepts numbers over limit for hours.")]
    [Arguments("11/30/2022 10000:59", null)] // Hours can't be over 9999
    [Arguments("aug 10, 2022 14:10", 44783.590277777781d, Skip = "Excel specific parsing of months accepts anything from three letters up to full name, but such pattern is not in any en-US DateTimeFormat pattern.")]
    [Arguments("august 10, 2022 14:10", 44783.590277777781d)]
    public async Task DateTime_Format22(string text, double? expectedValue) // Format 22 'm/d/yyyy h:mm'. Specification incorrectly states 'm/d/yy h:mm', but fixed per MS errata.
    {
        await AssertCoercion(text, expectedValue);
    }

    [Test]
    [Arguments("00:00", 0)] // Can parse zero
    [Arguments("90:00", 3.75)] // Minutes can be can be over 60
    [Arguments("59:59", 2.499305556)] // Even if looks like mm:ss, it is actually parsed as h:mm
    [Arguments("10:", 0.416666667)] // Last part can be omitted and zero is used
    [Arguments("9999:", 416.625)] // Upper limit of first part is parseable
    [Arguments("10000:", null)] // Part value over a limit is not parseable
    [Arguments(":5", null)] // Can't omit first part
    [Arguments("24:60", null)] // Only one part can be outside of limit, here are both
    [Arguments("30:59", 1.290972222)] // Hour part can be over 23
    [Arguments("23:300", 1.166666667)] // Minute part over over 59
    public async Task TimeSpan_Format20(string timeSpan, double? expectedValue) // 'h:mm'
    {
        await AssertCoercion(timeSpan, expectedValue, Tolerance);
    }

    [Test]
    [Arguments("0:01:01", 0.000706019)]
    [Arguments("000:01:01", null)] // Extra zeros.
    [Arguments("00:001:01", null)] // Three digits in a part that starts with 0
    [Arguments("0:01:001", null)] // Three digits in a part that starts with 0
    [Arguments("00:60:60", null)] // Only one part can be over the limit, but here are minutes and seconds
    [Arguments("24:60:00", null)] // Only one part can be over the limit, but here are hours and minutes
    [Arguments("24:00:60", null)] // Only one part can be over the limit, but here are hours and seconds
    [Arguments("23:60:06", 1.000069444)]
    [Arguments("  24   :  00  :   59  ", 1.00068287)] // Extra padding
    [Arguments("24:0:", 1)] // Last part can be omitted
    [Arguments("0::0", null)] // Parts in the middle can't be omitted
    [Arguments(":0:0", null)] // First part can't be omitted
    public async Task TimeSpan_Format21(string timeSpan, double? expectedValue) // 'h:mm:ss'
    {
        await AssertCoercion(timeSpan, expectedValue, Tolerance);
    }

    [Test]
    [Arguments("14:30.0", 0.010069444)] // Happy case, can be over 12 (to differ from AM/PM times)
    [Arguments("14:300.0", 0.013194444)] // Seconds part can be outside of normal range
    [Arguments("140:30.0", 0.097569444)] // Minutes part can be outside of normal range
    [Arguments("30:300.0", 0.024305556)] // Both parts can be outside the range
    [Arguments("140:60.0", null)] // Both hours and minutes are out of range
    [Arguments("60:000.0", null)] // The minutes part starts with 0, but has over 2 digits
    [Arguments("59:300.0", 0.044444444)] // Seconds are added to the minutes, the result is 1:04 minutes
    [Arguments("59:300.59", 0.044451273)] // Can specify 2 digit ms
    [Arguments("00:57.180", 0.000661806)] // Can specify 3 digit ms
    public async Task TimeSpan_Format47(string timeSpan, double? expectedValue) // 'mm:ss.0'
    {
        await AssertCoercion(timeSpan, expectedValue, Tolerance);
    }

    [Test]
    [Arguments("1,000", 1000)]
    [Arguments("1,00", null, Skip = ".NET parse methods ignores thousands separator, but excel enforces them.")]
    [Arguments("1,000,000", 1000000)]
    [Arguments("1,00,000", null, Skip = ".NET parse methods ignores thousands separator, but excel enforces them.")]
    [Arguments("(1,000)", -1000)]
    [Arguments("(100)", -100)]
    [Arguments("(-1)", null)]
    public async Task Number_Format37_38(string number, double? expectedValue) // Format 37+38 '#,##0 ;(#,##0)' '#,##0 ;[Red](#,##0)'
    {
        await AssertCoercion(number, expectedValue);
    }

    [Test]
    [Arguments("1,000.15", 1000.15)]
    [Arguments("(1,000.54)", -1000.54)]
    [Arguments("  (   1,000.54  )  ", -1000.54, Skip = "Excel can parse spaces within braces, but .NET parse method can't.")]
    public async Task Number_Format39_40(string number, double? expectedValue) // Format 39+40 '#,##0.00;(#,##0.00)'  '#,##0.00;[Red](#,##0.00)'
    {
        await AssertCoercion(number, expectedValue);
    }

    [Test]
    [Arguments("1e3", 1000)]
    [Arguments("1e+3", 1000)]
    [Arguments("1e-5", 0.00001)]
    [Arguments("1e0", 1)]
    [Arguments("1.5e2", 150)]
    [Arguments("1e2.5", null)] // Exponent can't be a fraction
    [Arguments("1.52e1", 15.2)]
    [Arguments("-1e2", -100)]
    [Arguments("1E2", 100)]
    public async Task Number_Format48_11(string number, double? expectedValue) // Format 48+11 '##0.0E+0' '0.00E+00'
    {
        await AssertCoercion(number, expectedValue);
    }

    [Test]
    [Arguments("$1", 1)]
    [Arguments("1$", null, Skip = ".NET parser allows currency symbol at the start or end, but Excel requires correct placement.")]
    [Arguments("($1)", -1)]
    [Arguments("-($1)", null)]
    [Arguments("$100.5", 100.5)]
    [Arguments("$100%", null)]
    [Arguments("($100%)", null)]
    public async Task Currency(string currency, double? expectedValue) // Currency doesn't have a format in ECMA-376, Part 1, §18.8.30, but VALUE includes currency formats
    {
        await AssertCoercion(currency, expectedValue);
    }

    [Test]
    [SetCulture("cs-CZ")]
    [Arguments("$1", null)] // Fallback currency doesn't work nor it should
    [Arguments("Kč 1", null, Skip = "Excel requires correct placement of currency symbol, while .NET parser accepts any position.")] // incorrect placement
    [Arguments("100.5", null)] // incorrect decimal placement
    [Arguments("10e2 Kč", 1000)]
    [Arguments("30-apr-2000", null)]
    [Arguments("02/28/20", null)]
    [Arguments("10:30 AM", 0.4375)] // AM seems to work for some reason
    [Arguments("10:30 dop.", 0.4375)]
    [Arguments("1-leden-2020", 43831)]
    [Arguments("1-led-2020", 43831)]
    [Arguments("led-5", 38353)]
    [Arguments("12:0:18 odp.", 0.50020833333333337d)]
    [Arguments("12:0:18 PM", 0.50020833333333337d)]
    [Arguments("12:0:18 odp", 0.50020833333333337d, Skip = "Excel can parse even partial PM designator, but .NET requires a full PM designator including the dot at the end.")]
    [Arguments("12:0:18 PM.", 0.50020833333333337d, Skip = "Excel allows PM designator with a dot at the end.")]
    [Arguments("11/30/2022 25:59", null)]
    [Arguments("25:70,05", 0.018171875)] // For min:sec fraction timespan, both can be over limit, also note use of decimal separator
    public async Task ParsingTokensAndFormatsDependOnCulture(string currency, double? expectedValue)
    {
        await AssertCoercion(currency, expectedValue);
    }

    private static async Task AssertCoercion(string text, double? expectedValue, double tolerance = 0)
    {
        using var wb = new XLWorkbook();
        var parsedValue = wb.Evaluate($"\"{text}\"*1");
        if (expectedValue is null)
            await Assert.That(parsedValue).IsEqualTo(XLError.IncompatibleValue);
        else
            await Assert.That((double)parsedValue).IsEqualTo(expectedValue.Value).Within(tolerance);
    }
}
