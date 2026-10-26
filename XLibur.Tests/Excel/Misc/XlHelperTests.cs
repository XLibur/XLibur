using System;
using System.Collections.Generic;
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.Misc;

public class XlHelperTests
{
    private static async Task CheckColumnNumber(int column)
    {
        await Assert.That(XLHelper.GetColumnNumberFromLetter(XLHelper.GetColumnLetterFromNumber(column))).IsEqualTo(column);
    }

    [Test]
    public async Task An_empty_column_letter_is_reported_as_empty_not_as_null()
    {
        await Assert.That(() => XLHelper.GetColumnNumberFromLetter("")).ThrowsExactly<ArgumentException>();
        await Assert.That(() => XLHelper.GetColumnNumberFromLetter(null!)).ThrowsExactly<ArgumentNullException>();
    }

    /// <summary>
    /// A quoted sheet name doubles each apostrophe it holds, wherever it is, as a formula writes it. The
    /// check accepted a quoted name that was one doubled apostrophe and nothing else, so
    /// <c>'Bob''s'!A1</c> was not a range address. A data validation that named such a sheet was then
    /// saved in the standard form instead of the <c>x14</c> extension, and <c>List(IXLRange)</c> stored
    /// the range as a literal list.
    /// </summary>
    [Test]
    [Arguments("'Bob''s'!A1")]
    [Arguments("'Bob''s'!$A$1:$A$3")]
    [Arguments("'''a'!A1")]
    [Arguments("'a'''!A1")]
    [Arguments("'O''Neil''s data'!$A$1:$B$2")]
    [Arguments("''''!A1")]
    [Arguments("'My Data'!A1")]
    [Arguments("Data!$A:$A")]
    [Arguments("$A$1:$B$2")]
    public async Task A_quoted_sheet_name_may_hold_a_doubled_apostrophe_anywhere(string address)
    {
        await Assert.That(XLHelper.IsValidRangeAddress(address)).IsTrue();
    }

    /// <summary>
    /// An apostrophe inside a quoted name must be doubled, the name cannot be empty, and a quoted name
    /// is closed by an apostrophe followed by <c>!</c>.
    /// </summary>
    [Test]
    [Arguments("'Bob's'!A1")]
    [Arguments("''!A1")]
    [Arguments("'Bob''s!A1")]
    [Arguments("'Bob''s'A1")]
    [Arguments("'a:b'!A1")]
    public async Task A_quoted_sheet_name_with_a_single_apostrophe_inside_or_no_name_is_not_a_range_address(
        string address)
    {
        await Assert.That(XLHelper.IsValidRangeAddress(address)).IsFalse();
    }

    /// <summary>
    /// An apostrophe around a sheet name is either matched by its pair or absent. The alternative that
    /// reads a name of letters and digits made each apostrophe optional on its own, as
    /// <c>'?\w+'?</c>, so a leading apostrophe with no closing one matched, and so did the other way
    /// round. <c>DataValidationWriter.UsesExternalSheet</c> then took everything before the last
    /// <c>!</c> as a sheet name and wrote the rule in the <c>x14</c> extension with that broken
    /// quoting, while the parser-based check refused the same text (#560).
    /// </summary>
    [Test]
    [Arguments("'Data!A1")]
    [Arguments("Data'!A1")]
    [Arguments("'Data!$A$1:$B$2")]
    [Arguments("Data'!$A$1:$B$2")]
    [Arguments("'Data!$A:$A")]
    [Arguments("Data'!1:1")]
    public async Task A_sheet_name_with_an_unterminated_apostrophe_is_not_a_range_address(string address)
    {
        await Assert.That(XLHelper.IsValidRangeAddress(address)).IsFalse();
    }

    /// <summary>
    /// The forms the letters-and-digits alternative exists for are unaffected: a name written without
    /// apostrophes, and the same name with both of them. A quoted name of word characters is read by
    /// the alternative for quoted names, which allows every one of them.
    /// </summary>
    [Test]
    [Arguments("Data!A1")]
    [Arguments("'Data'!A1")]
    [Arguments("Data!$A$1:$B$2")]
    [Arguments("'Data'!$A$1:$B$2")]
    [Arguments("Sheet_1!A1")]
    [Arguments("'Sheet_1'!A1")]
    [Arguments("Data1!1:1")]
    [Arguments("'Data1'!$A:$A")]
    public async Task A_sheet_name_of_letters_and_digits_is_read_with_both_apostrophes_or_neither(string address)
    {
        await Assert.That(XLHelper.IsValidRangeAddress(address)).IsTrue();
    }

    /// <summary>
    /// The callers of <see cref="XLHelper.IsValidRangeAddress(string)"/> that now read
    /// <c>'Bob''s'!A1:A3</c> as a range address resolve it as they resolve any other quoted name: a
    /// worksheet's <c>Range</c> takes the address, and a defined name at either scope still adds.
    /// </summary>
    [Test]
    public async Task A_range_on_a_sheet_whose_name_has_an_apostrophe_resolves()
    {
        using var wb = new XLWorkbook();
        var bobs = wb.AddWorksheet("Bob's");

        await Assert.That(wb.Range("'Bob''s'!A1:A3")!.Worksheet).IsSameReferenceAs(bobs);
        await Assert.That(bobs.Range("'Bob''s'!A1:A3")?.RangeAddress.ToString()).IsEqualTo("A1:A3");

        wb.DefinedNames.Add("Items", "'Bob''s'!$A$1:$A$3");
        bobs.DefinedNames.Add("Local", "'Bob''s'!$A$1:$A$3");

        await Assert.That(wb.DefinedName("Items")!.Ranges.Count).IsEqualTo(1);
        await Assert.That(bobs.DefinedName("Local").Ranges.Count).IsEqualTo(1);
    }

    [Test]
    [Arguments("A1\n")]
    [Arguments("Data!A1:B2\n")]
    [Arguments("'Bob''s'!$A:$A\n")]
    [Arguments("1:1\n")]
    [Arguments("A1\n\n")]
    public async Task A_range_address_ending_in_a_line_feed_is_not_a_range_address(string address)
    {
        await Assert.That(XLHelper.IsValidRangeAddress(address)).IsFalse();
    }

    [Test]
    [Arguments("A1:")]
    [Arguments("A1:B")]
    [Arguments("A:B2")]
    [Arguments("A$:B")]
    [Arguments("AAAA1")]
    [Arguments("A12345678")]
    [Arguments("12345678:1")]
    [Arguments("$$A1")]
    [Arguments("A1:B2:C3")]
    [Arguments("!A1")]
    [Arguments("Data!")]
    [Arguments("Da ta!A1")]
    public async Task Malformed_range_addresses_are_not_range_addresses(string address)
    {
        await Assert.That(XLHelper.IsValidRangeAddress(address)).IsFalse();
    }

    /// <summary>
    /// IsValidRangeAddress was a regex with a 150 ms timeout, and the timeout is wall-clock time: a
    /// thread descheduled or paused by GC mid-match threw RegexMatchTimeoutException on a valid
    /// address, from <c>Range("A1:B2")</c> and every other caller (#816). It is a scan now. This
    /// pins the scan to the language the regex accepted, with one change: the regex ended in
    /// '\Z', which also matches before a final '\n', so 'A1\n' was an address. The oracle ends in
    /// '\z' instead.
    /// </summary>
    /// <remarks>
    /// Random strings almost never form an address, so the inputs are valid addresses (with and
    /// without a sheet, quoted names with doubled apostrophes, '$' in each position, ranges, whole
    /// columns and whole rows) that are then mutated by inserting, deleting and swapping characters.
    /// The test checks that both outcomes are well represented.
    /// </remarks>
    [Test]
    public async Task IsValidRangeAddress_accepts_the_language_the_regex_accepted()
    {
        var oracle = CreateRangeAddressOracle();

        // Characters that matter to the grammar, and a few that only a Unicode-aware class treats
        // specially: '٣' is an Arabic-Indic digit, which '\d' matches, 'é' is a letter that '\w'
        // matches but '[a-zA-Z]' does not.
        const string alphabet = "'!$:AZaz019٣_é \n[]*?/\\";
        string[] sheetParts = ["Bob", "''", "My Data", "a", "'", "_", "٣", "é", "Sheet1", "[", ":", "*", "?"];
        string[] columns = ["A", "z", "AB", "XFD", "ABCD", "é"];
        string[] rows = ["1", "09", "1048576", "1234567", "12345678", "٣"];

        string Pick(string[] values, Random random) => values[random.Next(values.Length)];
        string Dollar(Random random) => random.Next(3) == 0 ? "$" : "";

        string Cell(Random random) => Dollar(random) + Pick(columns, random) + Dollar(random) + Pick(rows, random);

        string Sheet(Random random)
        {
            switch (random.Next(4))
            {
                case 0:
                    return "";
                case 1:
                    return Pick(["Data", "Sheet_1", "é", "٣", "x y"], random) + "!";
                default:
                    var name = "";
                    for (var i = random.Next(0, 4); i > 0; i--)
                        name += Pick(sheetParts, random);
                    return "'" + name + "'!";
            }
        }

        string Range(Random random)
        {
            switch (random.Next(4))
            {
                case 0:
                    return Cell(random);
                case 1:
                    return Cell(random) + ":" + Cell(random);
                case 2:
                    return Dollar(random) + Pick(rows, random) + ":" + Dollar(random) + Pick(rows, random);
                default:
                    return Dollar(random) + Pick(columns, random) + ":" + Dollar(random) + Pick(columns, random);
            }
        }

        var random = new Random(816);
        const int cases = 400_000;
        var accepted = 0;
        for (var i = 0; i < cases; i++)
        {
            var address = Sheet(random) + Range(random);
            for (var mutations = random.Next(0, 4); mutations > 0; mutations--)
            {
                var chars = new List<char>(address);
                switch (random.Next(3))
                {
                    case 0:
                        chars.Insert(random.Next(chars.Count + 1), alphabet[random.Next(alphabet.Length)]);
                        break;
                    case 1 when chars.Count > 0:
                        chars.RemoveAt(random.Next(chars.Count));
                        break;
                    case 2 when chars.Count > 1:
                        var a = random.Next(chars.Count);
                        var b = random.Next(chars.Count);
                        (chars[a], chars[b]) = (chars[b], chars[a]);
                        break;
                }

                address = new string(chars.ToArray());
            }

            var expected = !string.IsNullOrWhiteSpace(address) && oracle.IsMatch(address);
            if (expected)
                accepted++;

            if (XLHelper.IsValidRangeAddress(address) != expected)
                await Assert.That(XLHelper.IsValidRangeAddress(address)).IsEqualTo(expected).Because($"'{address}'");
        }

        Console.WriteLine($"IsValidRangeAddress fuzz: {cases} cases, the oracle accepts {accepted} ({100.0 * accepted / cases:F1}%)");

        // Both outcomes have to be common, or the comparison proves little.
        await Assert.That(accepted).IsGreaterThan(cases / 10);
        await Assert.That(accepted).IsLessThan(cases * 9 / 10);
    }

    /// <summary>
    /// '\w' is a Unicode class (letters, nonspacing marks, decimal digits, connector punctuation),
    /// which a sheet name written without apostrophes may hold. The fuzz test samples a handful of such
    /// characters, so this one tries every UTF-16 code unit, bare and inside apostrophes.
    /// </summary>
    [Test]
    public async Task IsValidRangeAddress_reads_every_character_in_a_sheet_name_as_the_regex_did()
    {
        var oracle = CreateRangeAddressOracle();
        for (var code = 0; code <= char.MaxValue; code++)
        {
            foreach (var address in new[] { (char)code + "!A1", "'" + (char)code + "'!A1", "'a" + (char)code + "'!A1" })
            {
                var expected = !string.IsNullOrWhiteSpace(address) && oracle.IsMatch(address);
                if (XLHelper.IsValidRangeAddress(address) != expected)
                    await Assert.That(XLHelper.IsValidRangeAddress(address)).IsEqualTo(expected).Because($"U+{code:X4}");
            }
        }
    }

    /// <summary>
    /// The pattern IsValidRangeAddress used to match, without the timeout, and with '\z' where it had
    /// '\Z'. NonBacktracking cannot time out, so it is an oracle that does not depend on the machine's
    /// load.
    /// </summary>
    private static System.Text.RegularExpressions.Regex CreateRangeAddressOracle() => new(
        @"\A"
        + "((" + XLHelper.QuotedSheetNamePattern + @"|\w+)!)?"
        + @"(\$?[a-zA-Z]{1,3}\$?\d{1,7}(:\$?[a-zA-Z]{1,3}\$?\d{1,7})?"
        + @"|\$?\d{1,7}:\$?\d{1,7}"
        + @"|\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})"
        + @"\z",
        System.Text.RegularExpressions.RegexOptions.NonBacktracking);

    [Test]
    public async Task InvalidA1Addresses()
    {
        await Assert.That(XLHelper.IsValidA1Address("")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("a")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("-1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AAAA1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("XFG1")).IsFalse();

        await Assert.That(XLHelper.IsValidA1Address("@A1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("@AA1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("@AAA1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("[A1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("[AA1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("[AAA1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("{A1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("{AA1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("{AAA1")).IsFalse();

        // A '$' is an anchor and is only meaningful before the column letters or before the row
        // digits. These were all reported valid while the implementation stripped every '$' and
        // validated what was left, which put IsValidA1Address in direct contradiction with
        // IsValidRangeAddress about the same string (D36).
        await Assert.That(XLHelper.IsValidA1Address("A$2$")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A2$")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("$A2$")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A$$2")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("$$A$$2")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A1$")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("$")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("$$")).IsFalse();

        // A row reference is digits and nothing else. IsValidRow used int.TryParse's default
        // NumberStyles.Integer, which also accepts surrounding whitespace and a leading sign, so
        // these reached the same contradiction with IsValidRangeAddress one layer down (D41).
        await Assert.That(XLHelper.IsValidA1Address("$A 1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A 1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A1 ")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A\t1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A+1")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("$A$+1")).IsFalse();

        // A trailing NUL survived the first fix: the number parser stops at a null terminator, so
        // NumberStyles.None still let "1\0" through. IsValidRow counts digits itself now.
        await Assert.That(XLHelper.IsValidA1Address("CC1\0")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("CC1\0\0\0\0\0\0\0\0")).IsFalse();

        // A public predicate answers, it does not throw. int.TryParse returned false for null;
        // reading .Length turned that into a NullReferenceException. IsValidColumn guards the
        // same way. Raised by CodeRabbit's review of PR #426.
        await Assert.That(XLHelper.IsValidRow(null!)).IsFalse();
        await Assert.That(XLHelper.IsValidColumn(null!)).IsFalse();

        await Assert.That(XLHelper.IsValidRow(" 1")).IsFalse();
        await Assert.That(XLHelper.IsValidRow("1 ")).IsFalse();
        await Assert.That(XLHelper.IsValidRow("+1")).IsFalse();
        await Assert.That(XLHelper.IsValidRow("1\0")).IsFalse();
        await Assert.That(XLHelper.IsValidRow("")).IsFalse();
        await Assert.That(XLHelper.IsValidRow("0")).IsFalse();
        await Assert.That(XLHelper.IsValidRow("00000001")).IsFalse();
        await Assert.That(XLHelper.IsValidRow("1")).IsTrue();
        await Assert.That(XLHelper.IsValidRow("1048576")).IsTrue();
        await Assert.That(XLHelper.IsValidRow("1048577")).IsFalse();
        await Assert.That(XLHelper.IsValidRow("9999999")).IsFalse();

        await Assert.That(XLHelper.IsValidA1Address("A1@")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AA1@")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AAA1@")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A1[")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AA1[")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AAA1[")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("A1{")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AA1{")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("AAA1{")).IsFalse();

        await Assert.That(XLHelper.IsValidA1Address("@A1@")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("@AA1@")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("@AAA1@")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("[A1[")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("[AA1[")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("[AAA1[")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("{A1{")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("{AA1{")).IsFalse();
        await Assert.That(XLHelper.IsValidA1Address("{AAA1{")).IsFalse();
    }

    [Test]
    public async Task PlusAA1_Is_Not_an_address()
    {
        await Assert.That(XLHelper.IsValidA1Address("+AA1")).IsFalse();
    }

    [Test]
    public async Task TestConvertColumnLetterToNumberAnd()
    {
        await CheckColumnNumber(1);
        await CheckColumnNumber(27);
        await CheckColumnNumber(28);
        await CheckColumnNumber(52);
        await CheckColumnNumber(53);
        await CheckColumnNumber(1000);
        await CheckColumnNumber(1353);
    }

    [Test]
    public async Task ValidA1Addresses()
    {
        await Assert.That(XLHelper.IsValidA1Address("A1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("A" + XLHelper.MaxRowNumber)).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("Z1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("Z" + XLHelper.MaxRowNumber)).IsTrue();

        await Assert.That(XLHelper.IsValidA1Address("AA1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("AA" + XLHelper.MaxRowNumber)).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("ZZ1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("ZZ" + XLHelper.MaxRowNumber)).IsTrue();

        await Assert.That(XLHelper.IsValidA1Address("AAA1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("AAA" + XLHelper.MaxRowNumber)).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address(XLHelper.MaxColumnLetter + "1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address(XLHelper.MaxColumnLetter + XLHelper.MaxRowNumber)).IsTrue();

        // The four anchor placements Excel actually allows. Tightening the rejection of a
        // misplaced '$' must not cost the legitimate ones.
        await Assert.That(XLHelper.IsValidA1Address("$A$1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("$A1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("A$1")).IsTrue();
        await Assert.That(XLHelper.IsValidA1Address("$AAA$" + XLHelper.MaxRowNumber)).IsTrue();
    }

    [Test]
    public async Task TestColumnLetterLookup()
    {
        var columnLetters = new List<string>();
        for (var c = 1; c <= XLHelper.MaxColumnNumber; c++)
        {
            var columnLetter = NaiveGetColumnLetterFromNumber(c);
            columnLetters.Add(columnLetter);

            await Assert.That(XLHelper.GetColumnLetterFromNumber(c)).IsEqualTo(columnLetter);
        }

        foreach (var cl in columnLetters)
        {
            var columnNumber = NaiveGetColumnNumberFromLetter(cl);
            await Assert.That(XLHelper.GetColumnNumberFromLetter(cl)).IsEqualTo(columnNumber);
        }
    }

    [Test]
    [Arguments("R")]
    [Arguments("C")]
    [Arguments("RC")]
    [Arguments("R111C222")]
    [Arguments("R[]C")]
    [Arguments("RC[]")]
    [Arguments("R[]C[]")]
    [Arguments("R[111]C222")]
    [Arguments("R111C[222]")]
    [Arguments("R[111]C[222]")]
    [Arguments("R[-111]C[-222]")]
    public async Task ValidRCAddresses(string address)
    {
        await Assert.That(XLHelper.IsValidRCAddress(address)).IsTrue();
    }

    [Test]
    [Arguments("RD")]
    [Arguments("CC")]
    [Arguments("R[-]C222")]
    [Arguments("R[]C[-]")]
    [Arguments("_R111C222")]
    public async Task InvalidRCAddresses(string address)
    {
        await Assert.That(XLHelper.IsValidRCAddress(address)).IsFalse();
    }

    #region Old XLHelper methods

    private static readonly string[] Letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

    /// <summary>
    /// These used to be the methods in XLHelper, but were later changed.
    /// We now use them as a check against the new methods
    /// Gets the column number of a given column letter.
    /// </summary>
    /// <param name="columnLetter"> The column letter to translate into a column number. </param>
    private static int NaiveGetColumnNumberFromLetter(string columnLetter)
    {
        if (string.IsNullOrEmpty(columnLetter)) throw new ArgumentNullException(nameof(columnLetter));

        columnLetter = columnLetter.ToUpper();

        //Extra check because we allow users to pass row col positions in as strings
        if (columnLetter[0] <= '9')
        {
            var retVal = int.Parse(columnLetter, XLHelper.NumberStyle, XLHelper.ParseCulture);
            return retVal;
        }

        var sum = 0;

        foreach (var t in columnLetter)
        {
            sum *= 26;
            sum += t - 'A' + 1;
        }

        return sum;
    }

    /// <summary>
    /// Gets the column letter of a given column number.
    /// </summary>
    /// <param name="columnNumber">The column number to translate into a column letter.</param>
    /// <param name="trimToAllowed">if set to <c>true</c> the column letter will be restricted to the allowed range.</param>
    private static string NaiveGetColumnLetterFromNumber(int columnNumber, bool trimToAllowed = false)
    {
        if (trimToAllowed) columnNumber = XLHelper.TrimColumnNumber(columnNumber);

        columnNumber--; // Adjust for start on column 1
        if (columnNumber <= 25)
        {
            return Letters[columnNumber];
        }
        var firstPart = (columnNumber) / 26;
        var remainder = ((columnNumber) % 26) + 1;
        return NaiveGetColumnLetterFromNumber(firstPart) + NaiveGetColumnLetterFromNumber(remainder);
    }

    #endregion Old XLHelper methods
}
