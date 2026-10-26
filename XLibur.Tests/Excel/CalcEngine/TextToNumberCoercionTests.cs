using System;
using System.Globalization;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
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
    [Arguments(" - 100 % ", -1)]
    public async Task Percent_Format9(string percent, double? expectedValue) // Format 9 '0%'
    {
        await AssertCoercion(percent, expectedValue);
    }

    [Test]
    [Arguments("100.5%", 1.005)]
    [Arguments("100 . 5%", null)]
    [Arguments(" - 100.59 % ", -1.0059)]
    [Arguments("0.123456%", 0.00123456)]
    [Arguments(".5%", 0.005)]
    [Arguments("  -.375 % ", -0.00375)]
    [Arguments("100.%", 1)]
    public async Task Percent_Format10(string percent, double? expectedValue) // Format 10 '0.00%'
    {
        await AssertCoercion(percent, expectedValue);
    }

    [Test]
    [Arguments("(100%)", -1)]
    [Arguments("(-100%)", null)] // Can't have minus sign inside the brackets
    [Arguments("-(100%)", null)] // Can't have minus sign outside the brackets
    [Arguments("1,000.00%", 10)]
    [Arguments("(1,000.00%)", -10)]
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

    // A whole number too big for an int is still a number. Excel 365 (probed over COM) gives
    // =VALUE("99999999999 1/2") = 99999999999.5 and =VALUE("12345678901234567890 1/2") = 1.23456789012345E+19.
    // This threw OverflowException before (#817).
    [Test]
    [Arguments("2147483647 1/2", 2147483647.5)]
    [Arguments("2147483648 1/2", 2147483648.5)]
    [Arguments("99999999999 1/2", 99999999999.5)]
    [Arguments("-99999999999 1/2", -99999999999.5)]
    [Arguments("12345678901234567890 1/2", 12345678901234567890d)]
    [Arguments("00000000000000000000001 1/2", 1.5)]
    public async Task Fraction_WholeNumberBiggerThanInt(string fraction, double expectedValue)
    {
        await AssertCoercion(fraction, expectedValue);
    }

    // The same answers Excel gave when probed over COM, through the calc engine. They go through
    // arithmetic on text, as XLibur's VALUE function has its own parser that doesn't read
    // fractions at all ("fraction isn't parsed, maybe in the future" in Text.cs).
    [Test]
    [Arguments("\"99999999999 1/2\"+0", 99999999999.5)]
    [Arguments("\"2147483648 1/2\"+0", 2147483648.5)]
    [Arguments("\"1 1/2\"+0", 1.5)]
    public async Task Fraction_Arithmetic_MatchesExcel(string formula, double expectedValue)
    {
        using var wb = new XLWorkbook();
        await Assert.That((double)wb.Evaluate(formula)).IsEqualTo(expectedValue);
    }

    [Test]
    [Arguments("(\"1 1/2\"&CHAR(10))+0")]
    [Arguments("(\"1 1/2\"&CHAR(13))+0")]
    public async Task Fraction_Arithmetic_LineBreakIsNotANumber(string formula)
    {
        using var wb = new XLWorkbook();
        await Assert.That(wb.Evaluate(formula)).IsEqualTo(XLError.IncompatibleValue);
    }

    // A whole number beyond double.MaxValue is not a number. Excel's VALUE caps its text at
    // 255 characters, which XLibur doesn't copy, so there is no Excel answer to match.
    [Test]
    public async Task Fraction_WholeNumberBiggerThanDouble_IsNotANumber()
    {
        await Assert.That(FractionParser.TryParse(new string('9', 400) + " 1/2", out _)).IsFalse();
    }

    // The regex ended in '$', which also matches before a final '\n', so "1 1/2\n" was a number.
    // Excel gives #VALUE! for =VALUE("1 1/2"&CHAR(10)) and for CHAR(13), so a trailing line break
    // is not a number now. Only the space U+0020 is allowed around a fraction.
    [Test]
    [Arguments("1 1/2\n")]
    [Arguments("1 1/2\r")]
    [Arguments("1 1/2\r\n")]
    [Arguments("\n1 1/2")]
    [Arguments("1 1/2 \n")]
    [Arguments("1 1/2\t")]
    [Arguments("1 1/2\u00A0")] // no-break space, U+00A0
    public async Task Fraction_LineBreaksAndOtherWhitespace_AreNotANumber(string text)
    {
        await Assert.That(FractionParser.TryParse(text, out _)).IsFalse();
    }

    /// <summary>
    /// FractionParser was a regex with a 150 ms timeout, and the timeout is wall-clock time: a
    /// thread descheduled or paused by GC mid-match threw RegexMatchTimeoutException whatever the
    /// text, and TextToNumber reaches it for every text that is not a plain number (#817). It is a
    /// scan now. This pins the scan to the language the regex accepted, and to the value it parsed,
    /// with two changes: the regex ended in '$', which also matches before a final '\n', so
    /// "1 1/2\n" was a number; and a whole number that overflows an int threw. The oracle ends in
    /// '\z' instead, and reads the whole number as a double from its own capture group. A whole
    /// number beyond the range of a double is left out: the scan says it is not a number, which
    /// Fraction_WholeNumberBiggerThanDouble_IsNotANumber pins.
    /// </summary>
    [Test]
    public async Task FractionParser_accepts_the_language_the_regex_accepted()
    {
        var oracle = new System.Text.RegularExpressions.Regex(
            @"^ *([+-]?) *([0-9]+) ([0-9]{1,5})/([0-9]{1,5}) *\z",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant
            | System.Text.RegularExpressions.RegexOptions.NonBacktracking);

        // Random strings almost never form a fraction, so build one from parts and then mutate it.
        // '\u0663' is an Arabic-Indic digit, which '\d' would match and '[0-9]' does not.
        string[] signs = ["", "+", "-", " +", "- ", " - ", "  "];
        string[] wholes = ["0", "1", "12", "007", "2147483647", "\u0663"];
        string[] fractionParts = ["0", "1", "00", "9", "32767", "32768", "00001", "99999", "100000", "000000", "\u0663"];
        const string noise = " /+-0123456789\n\r\t\u0663.";

        var random = new Random(17);
        var accepted = 0;
        const int cases = 100_000;
        for (var i = 0; i < cases; i++)
        {
            // One in four wholes has 10 to 20 digits, which no int can hold.
            var whole = random.Next(4) == 0
                ? RandomDigits(random, random.Next(10, 21))
                : wholes[random.Next(wholes.Length)];
            var text = signs[random.Next(signs.Length)]
                       + whole
                       + (random.Next(12) == 0 ? "  " : " ")
                       + fractionParts[random.Next(fractionParts.Length)]
                       + "/"
                       + fractionParts[random.Next(fractionParts.Length)]
                       + new string(' ', random.Next(3));

            // Mutate: drop, insert or replace up to two characters on most inputs.
            var chars = new System.Text.StringBuilder(text);
            for (var m = random.Next(3); m > 0 && random.Next(4) != 0; m--)
            {
                var at = random.Next(chars.Length + 1);
                switch (random.Next(3))
                {
                    case 0 when at < chars.Length:
                        chars.Remove(at, 1);
                        break;
                    case 1:
                        chars.Insert(at, noise[random.Next(noise.Length)]);
                        break;
                    case 2 when at < chars.Length:
                        chars[at] = noise[random.Next(noise.Length)];
                        break;
                }
            }

            text = chars.ToString();

            var match = oracle.Match(text);
            if (match.Success && double.IsInfinity(double.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture)))
                continue; // Not a number now, where the old code threw.

            bool expectedOk;
            double expected = 0;
            if (!match.Success)
            {
                expectedOk = false;
            }
            else
            {
                var numerator = int.Parse(match.Groups[3].Value);
                var denominator = int.Parse(match.Groups[4].Value);
                expectedOk = denominator != 0 && denominator <= short.MaxValue && numerator <= short.MaxValue;
                if (expectedOk)
                {
                    var wholeNumber = double.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
                    var fraction = wholeNumber + numerator / (double)denominator;
                    expected = match.Groups[1].Value == "-" ? -fraction : fraction;
                }
            }

            var actualOk = FractionParser.TryParse(text, out var actual);
            if (actualOk != expectedOk || (expectedOk && actual != expected))
            {
                await Assert.That((actualOk, actual)).IsEqualTo((expectedOk, expected)).Because($"'{text}'");
            }

            if (expectedOk)
                accepted++;
        }

        // The comparison is only worth something if it reaches valid fractions.
        await Assert.That(accepted).IsGreaterThan(cases / 10)
            .Because($"the oracle accepted {accepted} of {cases} inputs");
    }

    private static string RandomDigits(Random random, int length)
    {
        var digits = new char[length];
        for (var i = 0; i < digits.Length; i++)
            digits[i] = (char)('0' + random.Next(10));

        return new string(digits);
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
    [Arguments(" 1-jan-2022", null)] // Can't have whitespaces at the start
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
    [Arguments("1-marc", 44621)]
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
    [Arguments("jan-32", 11689)] // 1932-01-01
    [Arguments("feb-29", 47150)] // 2029-02-01
    [Arguments("feb-30", 10990)] // 1930-02-01
    [Arguments("feb-31", 11355)] // 1931-02-01
    [Arguments("feb-003", null)] // three digits not allowed
    [Arguments("aug   -   55", 20302)] // spaces are allowed inside the pattern
    [Arguments(" aug-55", null)] // starting spaces not allowed
    [Arguments("aug-55 ", 20302)] // trailing spaces allowed
    [Arguments("MaR-42", 15401)] // case-insensitive
    [Arguments("march-55", 20149)]
    [Arguments("ma-2", null)] // Name of month must be at least three chars long
    public async Task Date_Format17(string text, double? expectedValue) // Format 17 'mmm-yy'
    {
        await AssertCoercion(text, expectedValue);
    }

    // Cultures that write the month before the day have an extra 'mmm-dd' pattern ahead of 'mmm-yy',
    // so a number that is a valid day of the month is read as a day of the current year rather than
    // as a year. Numbers that are not ('jan-32', 'feb-29') stay in Date_Format17 above.
    [Test]
    [Arguments("jan-02", 44563)] // 2022-01-02
    [Arguments("jan-31", 44592)] // 2022-01-31
    [Arguments("marc-2", 44622)] // 2022-03-02, a month name can be any prefix at least three letters long
    public async Task Date_Format17_DayOfCurrentYear(string text, double expectedValue) // Format 17 'mmm-dd'
    {
        var date = DateTime.FromOADate(expectedValue);
        var inCurrentYear = new DateTime(DateTime.Now.Year, date.Month, date.Day, 0, 0, 0, DateTimeKind.Unspecified);

        await AssertCoercion(text, inCurrentYear.ToOADate());
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
    [Arguments("12:0:18 aM", 0.00020833333333333335d)] // case-insensitive AM designator
    [Arguments("13:0:0 PM", null)] // hours can't be outside 0-12, unlike other format
    [Arguments("13:0:0 AM", null)]
    [Arguments("00:60:00 AM", null)] // minutes can't be outside 0-59, unlike other format
    [Arguments("00:59:00 AM", 0.040972222222222222d)]
    [Arguments("00:00:60 AM", null)] // seconds can't be outside 0-59, unlike other format
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
    [Arguments("11/30/2022 24:59", 44896.040972222225d)] // Hours overflow into the next day
    [Arguments("11/30/2022 24:60", null)] // Both parts are out of range
    [Arguments("11/30/2022 23:160", 44896.069444444445d)] // Minutes overflow into the next day
    [Arguments("11/30/2022 9999:59", 45311.665972222225d)] // Hours can reach 9999
    [Arguments("11/30/2022 10000:59", null)] // Hours can't be over 9999
    [Arguments("aug 10, 2022 14:10", 44783.590277777781d)]
    [Arguments("august 10, 2022 14:10", 44783.590277777781d)]
    // ReSharper disable once GrammarMistakeInComment
    public async Task DateTime_Format22(string text, double? expectedValue) // Format 22 'm/d/yyyy h:mm'. Specification incorrectly states 'm/d/yy h:mm', but fixed per MS errata.
    {
        await AssertCoercion(text, expectedValue);
    }

    [Test]
    [Arguments("00:00", 0)] // Can parse zero
    [Arguments("90:00", 3.75)] // Minutes can be over 60
    [Arguments("59:59", 2.499305556)] // Even if it looks like mm:ss, it is actually parsed as h:mm
    [Arguments("10:", 0.416666667)] // Last part can be omitted and zero is used
    [Arguments("9999:", 416.625)] // Upper limit of first part is parseable
    [Arguments("10000:", null)] // Part value over a limit is not parseable
    [Arguments(":5", null)] // Can't omit first part
    [Arguments("24:60", null)] // Only one part can be outside of limit, here are both
    [Arguments("30:59", 1.290972222)] // Hour part can be over 23
    [Arguments("23:300", 1.166666667)] // Minute part over 59
    [Arguments("4294967306:", null)] // 2^32 + 10, would wrap to 10 in 32-bit arithmetic
    [Arguments("2147483648:", null)] // 2^31, would wrap to int.MinValue
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
    [Arguments("24   :  00  :   59  ", 1.00068287)] // Extra padding inside and at the end
    [Arguments("  24   :  00  :   59  ", null)] // but not at the start, measured over COM (#831)
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
    [Arguments("00:57.18000000000", 0.000661806)] // Fraction digits past an int's range don't wrap
    public async Task TimeSpan_Format47(string timeSpan, double? expectedValue) // 'mm:ss.0'
    {
        await AssertCoercion(timeSpan, expectedValue, Tolerance);
    }

    [Test]
    [Arguments("1,000", 1000)]
    [Arguments("1,00", null)]
    [Arguments("1,000,000", 1000000)]
    [Arguments("1,00,000", null)]
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
    [Arguments("  (   1,000.54  )  ", -1000.54)]
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
    [Arguments("1$", null)]
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
    [Arguments("Kč 1", null)] // incorrect placement
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
    [Arguments("12:0:18 odp", 0.50020833333333337d)]
    [Arguments("12:0:18 PM.", 0.50020833333333337d)]
    [Arguments("11/30/2022 25:59", null)]
    [Arguments("25:70,05", 0.018171875)] // For min:sec fraction timespan, both can be over limit, also note use of decimal separator
    public async Task ParsingTokensAndFormatsDependOnCulture(string currency, double? expectedValue)
    {
        await AssertCoercion(currency, expectedValue);
    }

    // Measured in Excel 365 over COM on 2026-10-26 with Windows set to day-first (en-AU) and then
    // month-first (M/d/yyyy) order (#831). A day and a month with no year are of the current year.
    // Two numbers are read first as a day and a month in the culture's order, and when that is no
    // date, as a month and a year: a two-digit year below 30 is 20xx, 30 and above 19xx.

    [Test]
    [SetCulture("en-AU")]
    [Arguments("1/2", 2, 1)]
    [Arguments("3/4", 4, 3)]
    [Arguments("31/12", 12, 31)]
    [Arguments("13/1", 1, 13)]
    [Arguments("1-1", 1, 1)]
    [Arguments("1-2", 2, 1)]
    [Arguments("31-12", 12, 31)]
    public async Task Date_DayAndMonth_DayFirst(string text, int month, int day)
    {
        await AssertCoercion(text, ThisYear(month, day));
    }

    [Test]
    [Arguments("1/2", 1, 2)]
    [Arguments("3/4", 3, 4)]
    [Arguments("12/31", 12, 31)]
    [Arguments("1/13", 1, 13)]
    [Arguments("1-1", 1, 1)]
    [Arguments("1-2", 1, 2)]
    [Arguments("12-31", 12, 31)]
    [Arguments("1-22", 1, 22)]
    [Arguments("12/22", 12, 22)]
    public async Task Date_MonthAndDay_MonthFirst(string text, int month, int day)
    {
        await AssertCoercion(text, ThisYear(month, day));
    }

    [Test]
    [SetCulture("en-AU")]
    [Arguments("12/31", 11658)] // 1931-12-01: 31 is no month, so it is a year
    [Arguments("1/13", 41275)] // 2013-01-01
    [Arguments("1/2022", 44562)]
    [Arguments("1-2022", 44562)]
    [Arguments("12-2022", 44896)]
    [Arguments("1-22", 44562)]
    [Arguments("1/22", 44562)]
    [Arguments("12/22", 44896)]
    [Arguments("1-99", 36161)]
    [Arguments("1-1999", 36161)]
    [Arguments("30/2", null)] // Neither 30 February nor a month 30
    [Arguments("0/1", null)]
    [Arguments("1/0", null)] // Day-first takes no one-digit year
    [Arguments("13-2022", null)]
    [Arguments("13/22", null)]
    [Arguments("0-2022", null)]
    [Arguments("1 2", null)] // A space doesn't separate two numbers
    [Arguments("1 2022", null)]
    [Arguments("1/2/3/4", null)]
    public async Task Date_MonthAndYear_DayFirst(string text, double? expectedValue)
    {
        await AssertCoercion(text, expectedValue);
    }

    [Test]
    [Arguments("2/30", 10990)] // 1930-02-01
    [Arguments("1/0", 36526)] // 2000-01-01: month-first takes a one-digit year
    [Arguments("1/2022", 44562)]
    [Arguments("1-2022", 44562)]
    [Arguments("31/12", null)]
    [Arguments("13/1", null)]
    [Arguments("31-12", null)]
    [Arguments("13-1", null)]
    [Arguments("13-2022", null)]
    [Arguments("13/22", null)]
    public async Task Date_MonthAndYear_MonthFirst(string text, double? expectedValue)
    {
        await AssertCoercion(text, expectedValue);
    }

    [Test]
    [Arguments("1 mar", 3, 1)]
    [Arguments("1 March", 3, 1)]
    [Arguments("31 mar", 3, 31)]
    [Arguments("1  mar", 3, 1)]
    [Arguments(" 1 mar", 3, 1)] // Format 16 allows a leading space
    [Arguments("1 marc", 3, 1)]
    public async Task Date_Format16_WithSpace(string text, int month, int day)
    {
        await AssertCoercion(text, ThisYear(month, day));
    }

    [Test]
    [Arguments("32 mar", null)]
    [Arguments("1 mar 2022", 44621)]
    [Arguments("1 mar 22", 44621)]
    [Arguments("2000/04/30", 36646)]
    [Arguments("2000/4/30", 36646)]
    [Arguments("2000/04/30 10:30", 36646.4375)]
    [Arguments("2000/13/1", null)]
    [Arguments("2000/1/32", null)]
    public async Task Date_DayMonthNameAndYearFirst(string text, double? expectedValue)
    {
        await AssertCoercion(text, expectedValue);
    }

    // Excel requires each group after the first to have at least the group size of digits, not
    // exactly it, and puts no limit on the first group.
    [Test]
    [Arguments("1,0000", 10000)]
    [Arguments("12,0000", 120000)]
    [Arguments("123,0000", 1230000)]
    [Arguments("1234,0000", 12340000)]
    [Arguments("1,00000", 100000)]
    [Arguments("1,000000", 1000000)]
    [Arguments("1,0000,000", 10000000)]
    [Arguments("1,000,0000", 10000000)]
    [Arguments("1,0000.5", 10000.5)]
    [Arguments("$1,0000", 10000)]
    [Arguments("-1,0000", -10000)]
    [Arguments("(1,0000)", -10000)]
    [Arguments("1,0000%", 100)]
    [Arguments("1,0", null)]
    [Arguments("12,34,567", null)]
    [Arguments(",1000", null)]
    [Arguments("1,", null)]
    public async Task Number_GroupsOfAtLeastTheGroupSize(string number, double? expectedValue)
    {
        await AssertCoercion(number, expectedValue);
    }

    // An AM/PM designator must be apart from the digits before it.
    [Test]
    [Arguments("10:30am", null)]
    [Arguments("10:30pm", null)]
    [Arguments("10:30AM", null)]
    [Arguments("10:30:15am", null)]
    [Arguments("10:30a", null)]
    [Arguments("10am", null)]
    [Arguments("1/2/2022 10:30am", null)]
    [Arguments("10:30 a", 0.4375)]
    [Arguments("10:30 am", 0.4375)]
    [Arguments("10 am", 0.41666666666666669)]
    [Arguments("1/2/2022 10:30 am", 44563.4375)]
    public async Task Time_DesignatorAfterASpace(string text, double? expectedValue)
    {
        await AssertCoercion(text, expectedValue, Tolerance);
    }

    // Text*1 refuses a leading space on every date and time but format 16 ('d-mmm'), and keeps
    // accepting it on numbers. VALUE skips leading spaces, so it reads all of these.
    [Test]
    [Arguments(" 10:30 AM", 0.4375)]
    [Arguments(" 10:30", 0.4375)]
    [Arguments(" 10:30:15", 0.43767361111111114)]
    [Arguments(" 90:00", 3.75)]
    [Arguments(" 1/2/2022", 44563)]
    [Arguments(" 1-mar-2022", 44621)]
    [Arguments(" 30-apr-2000", 36646)]
    [Arguments(" 2000-04-30", 36646)]
    public async Task Date_LeadingSpace_IsRefusedByCoercionButNotByValue(string text, double valueResult)
    {
        await AssertCoercion(text, null);
        await AssertValue(text, valueResult);
    }

    [Test]
    [Arguments(" 1/2", 1, 2)]
    [Arguments(" mar-22", 3, 22)]
    public async Task Date_ThisYear_LeadingSpace_IsRefusedByCoercionButNotByValue(string text, int month, int day)
    {
        await AssertCoercion(text, null);
        await AssertValue(text, ThisYear(month, day));
    }

    private static async Task AssertValue(string text, double expectedValue)
    {
        using var wb = new XLWorkbook();
        await Assert.That((double)wb.Evaluate($"VALUE(\"{text}\")")).IsEqualTo(expectedValue).Within(Tolerance);
    }

    // DATEVALUE and TIMEVALUE skip leading spaces, as VALUE does.
    [Test]
    [Arguments("DATEVALUE(\" 1/2/2022\")", 44563)]
    [Arguments("DATEVALUE(\" 1-jan-2022\")", 44562)]
    [Arguments("TIMEVALUE(\" 10:30 AM\")", 0.4375)]
    [Arguments("TIMEVALUE(\" 10:30\")", 0.4375)]
    public async Task DateValueAndTimeValue_SkipLeadingSpaces(string formula, double expectedValue)
    {
        using var wb = new XLWorkbook();
        await Assert.That((double)wb.Evaluate(formula)).IsEqualTo(expectedValue).Within(Tolerance);
    }

    [Test]
    [Arguments("5 a", 0.20833333333333334)]
    [Arguments("5 A", 0.20833333333333334)]
    [Arguments("5 p", 0.70833333333333337)]
    [Arguments("12 p", 0.5)]
    public async Task Time_HourAndOneLetterDesignator(string text, double expectedValue)
    {
        await AssertCoercion(text, expectedValue, Tolerance);
    }

    [Test]
    [Arguments(" 1-mar")]
    [Arguments(" 1 mar")]
    public async Task Date_Format16_LeadingSpace_IsAccepted(string text)
    {
        await AssertCoercion(text, ThisYear(3, 1));
    }

    [Test]
    [Arguments(" 1 1/2", 1.5)]
    [Arguments(" 1", 1)]
    [Arguments(" 1%", 0.01)]
    [Arguments(" $1", 1)]
    [Arguments(" (1)", -1)]
    [Arguments(" 1e3", 1000)]
    public async Task Number_LeadingSpace_IsAccepted(string text, double expectedValue)
    {
        await AssertCoercion(text, expectedValue);
    }

    // Excel reads a two-digit year as 1930-2029 (Windows' default window) in every format. .NET's
    // window is 1950-2049, so '22/04/30' was 2030.
    [Test]
    [SetCulture("en-AU")]
    [Arguments("22/04/30", 11070)] // 1930-04-22
    [Arguments("22/04/29", 47230)] // 2029-04-22
    [Arguments("1/1/49", 17899)]
    [Arguments("1/1/50", 18264)]
    public async Task Date_TwoDigitYear_DayFirst(string text, double expectedValue)
    {
        await AssertCoercion(text, expectedValue);
    }

    [Test]
    [Arguments("1/2/30 10:30", 10960.4375)] // 1930-01-02 10:30
    [Arguments("1-jan-30", 10959)]
    [Arguments("1-jan-29", 47119)]
    [Arguments("1 jan 30", 10959)]
    [Arguments("1-jan-49", 17899)]
    [Arguments("jan-49", 17899)]
    public async Task Date_TwoDigitYear(string text, double expectedValue)
    {
        await AssertCoercion(text, expectedValue);
    }

    // The two-digit-year culture is cached per culture instance. CultureInfo.Equals compares only
    // the name, so a cache keyed by it handed a second en-US the first one's date separator.
    [Test]
    public async Task Date_CultureWithItsOwnDateSeparator_KeepsIt()
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("en-US");
            await AssertCoercion("1/2/2022", 44563);

            var dotted = new CultureInfo("en-US");
            dotted.DateTimeFormat.DateSeparator = ".";
            CultureInfo.CurrentCulture = dotted;
            await AssertCoercion("1.2.2022", 44563);
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }

    private static double ThisYear(int month, int day) =>
        new DateTime(DateTime.Now.Year, month, day, 0, 0, 0, DateTimeKind.Unspecified).ToOADate();

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
