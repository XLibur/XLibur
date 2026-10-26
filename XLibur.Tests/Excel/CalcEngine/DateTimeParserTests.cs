using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel.CalcEngine;

namespace XLibur.Tests.Excel.CalcEngine;

public class DateTimeParserTests
{
    // The two replaces the scan took over from, as the oracle. A test has time to spare, so the
    // timeout is long enough that a busy machine cannot trip it.
    private static readonly TimeSpan OracleTimeout = TimeSpan.FromSeconds(10);
    private static readonly Regex LeadingZeroMonth = new("(?<!M)MM(?!M)", RegexOptions.None, OracleTimeout);
    private static readonly Regex LeadingZeroDay = new("(?<!d)dd(?!d)", RegexOptions.None, OracleTimeout);

    private static string Oracle(string pattern) =>
        LeadingZeroDay.Replace(LeadingZeroMonth.Replace(pattern, "M"), "d");

    [Test]
    [Arguments("MM/dd/yyyy", "M/d/yyyy")]
    [Arguments("dd.MM.yyyy", "d.M.yyyy")]
    [Arguments("M/d/yyyy", "M/d/yyyy")]
    [Arguments("dd MMM yyyy", "d MMM yyyy")]
    [Arguments("dddd, MMMM dd, yyyy", "dddd, MMMM d, yyyy")]
    [Arguments("MMMMM", "MMMMM")]
    [Arguments("ddd", "ddd")]
    [Arguments("d 'de' MMMM 'de' yyyy", "d 'de' MMMM 'de' yyyy")]
    [Arguments("yyyy'年'MM'月'dd'日'", "yyyy'年'M'月'd'日'")]
    [Arguments("MMdd", "Md")]
    [Arguments("", "")]
    public async Task AllowSingleDigitMonthAndDay_shortens_runs_of_exactly_two(string pattern, string expected)
    {
        await Assert.That(DateTimeParser.AllowSingleDigitMonthAndDay(pattern)).IsEqualTo(expected);
    }

    /// <summary>
    /// The patterns were rewritten with two regex replaces that had a 150 ms timeout. The timeout
    /// is wall-clock time, so a thread descheduled or paused by GC threw RegexMatchTimeoutException
    /// out of the first text-to-date coercion of each culture (#825). The scan must give the same
    /// pattern as the two replaces for every date pattern of every culture.
    /// </summary>
    [Test]
    public async Task AllowSingleDigitMonthAndDay_matches_the_regex_replaces_for_every_culture()
    {
        foreach (var culture in CultureInfo.GetCultures(CultureTypes.AllCultures))
        {
            foreach (var format in new[] { 'd', 'D' })
            {
                foreach (var pattern in culture.DateTimeFormat.GetAllDateTimePatterns(format))
                {
                    var expected = Oracle(pattern);
                    var actual = DateTimeParser.AllowSingleDigitMonthAndDay(pattern);
                    if (actual != expected)
                        await Assert.That(actual).IsEqualTo(expected).Because($"'{pattern}' of {culture.Name}");
                }
            }
        }
    }

    [Test]
    public async Task AllowSingleDigitMonthAndDay_matches_the_regex_replaces_on_random_text()
    {
        const string alphabet = "Mdm D/'y\n";
        var random = new Random(825);
        for (var i = 0; i < 100_000; i++)
        {
            var chars = new char[random.Next(0, 12)];
            for (var j = 0; j < chars.Length; j++)
                chars[j] = alphabet[random.Next(alphabet.Length)];
            var pattern = new string(chars);

            var expected = Oracle(pattern);
            var actual = DateTimeParser.AllowSingleDigitMonthAndDay(pattern);
            if (actual != expected)
                await Assert.That(actual).IsEqualTo(expected).Because($"'{pattern}'");
        }
    }
}
