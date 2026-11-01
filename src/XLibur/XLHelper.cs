using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;
using XLibur.Graphics;

namespace XLibur.Excel;

/// <summary>
/// Common methods
/// </summary>
public static partial class XLHelper
{
    public const int MinRowNumber = 1;
    public const int MinColumnNumber = 1;
    public const int MaxRowNumber = 1048576;
    public const int MaxColumnNumber = 16384;
    public const string MaxColumnLetter = "XFD";
    public const double Epsilon = 1e-10;

    /// <summary>
    /// Maximum number of code units that can be stored in a cell.
    /// </summary>
    internal const int CellTextLimit = 32767;

    /// <summary>
    /// 1900 calendar serial date upper limit (exclusive).
    /// </summary>
    internal const int Calendar1900UpperLimit = 2958466;

    /// <summary>
    /// 1904 calendar serial date upper limit (exclusive).
    /// </summary>
    internal const int Calendar1904UpperLimit = 2957004;

    internal static readonly Encoding NoBomUTF8 = new UTF8Encoding(false);

    public static string LastCell => $"{MaxColumnLetter}{MaxRowNumber}";

    internal static readonly NumberStyles NumberStyle = NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign |
                                                        NumberStyles.AllowLeadingWhite |
                                                        NumberStyles.AllowTrailingWhite | NumberStyles.AllowExponent;

    internal static readonly CultureInfo ParseCulture = CultureInfo.InvariantCulture;

    /// <summary>
    /// Comparer used to compare sheet names.
    /// </summary>
    internal static readonly StringComparer SheetComparer = StringComparer.OrdinalIgnoreCase;

    /// <summary>
    /// Comparer used to compare defined names.
    /// </summary>
    internal static readonly StringComparer NameComparer = StringComparer.OrdinalIgnoreCase;

    /// <summary>
    /// Comparer of function names.
    /// </summary>
    internal static readonly StringComparer FunctionComparer = StringComparer.OrdinalIgnoreCase;

    /// <summary>
    /// A sheet name written in apostrophes. The surrounding apostrophes are required, and an
    /// apostrophe in the name itself is doubled and may appear anywhere in it, as in
    /// <c>'Bob''s'</c>.
    /// </summary>
    /// <remarks>
    /// Shared with the legacy formula shifter's own reference regex (see
    /// <c>XLCellFormulaShifter.Legacy.cs</c>), which held a second copy of this sub-pattern with the
    /// <c>+</c> bound to the negated class alone. That copy could not read a doubled apostrophe, and
    /// because it is unanchored it matched <c>'Bob''s'!A1</c> from the second apostrophe onwards, as a
    /// reference to a sheet named <c>s</c> (#570). One definition so a third copy cannot drift again.
    /// </remarks>
    internal const string QuotedSheetNamePattern = @"\'(?:[^\[\]\*/\\\?:\']|\'\')+\'";

    internal static readonly Regex NamedRangeReferenceRegex = NamedRangeReferenceRegexCompiled();

    private static readonly string[] Letters =
    [
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V",
        "W", "X", "Y", "Z"
    ];

    private static readonly string[] AllLetters;
    private static readonly Dictionary<string, int> LetterIndexes;

    static XLHelper()
    {
        AllLetters = new string[MaxColumnNumber];
        LetterIndexes = new Dictionary<string, int>(MaxColumnNumber, StringComparer.Create(ParseCulture, true));
        for (var i = 0; i < MaxColumnNumber; i++)
        {
            string letter;
            if (i < 26)
                letter = Letters[i];
            else if (i < 26 * 27)
                letter = Letters[i / 26 - 1] + Letters[i % 26];
            else
                letter = Letters[(i - 26) / 26 / 26 - 1] + Letters[(i / 26 - 1) % 26] + Letters[i % 26];
            AllLetters[i] = letter;
            LetterIndexes.Add(letter, i + 1);
        }
    }

    /// <summary>
    /// Gets the column number of a given column letter.
    /// </summary>
    /// <param name="columnLetter"> The column letter to translate into a column number. </param>
    public static int GetColumnNumberFromLetter(string columnLetter)
    {
        ArgumentException.ThrowIfNullOrEmpty(columnLetter);

        //Extra check because we allow users to pass row col positions in as strings
        if (columnLetter[0] <= '9')
        {
            return int.Parse(columnLetter, NumberStyle, ParseCulture);
        }

        if (LetterIndexes.TryGetValue(columnLetter, out var retVal))
            return retVal;

        throw new ArgumentOutOfRangeException(columnLetter + " is not recognized as a column letter");
    }

    /// <summary>
    /// Gets the column number of a given column letter, without allocating a string for the letters.
    /// The column letters are decoded as a bijective base-26 number (A=1, Z=26, AA=27, …).
    /// </summary>
    /// <param name="columnLetter"> The column letter(s) to translate into a column number. </param>
    internal static int GetColumnNumberFromLetter(ReadOnlySpan<char> columnLetter)
    {
        if (columnLetter.IsEmpty) throw new ArgumentException("Column letter must not be empty.", nameof(columnLetter));

        //Extra check because we allow users to pass row col positions in as strings
        if (columnLetter[0] <= '9')
            return int.Parse(columnLetter, NumberStyle, ParseCulture);

        var result = 0;
        foreach (var c in columnLetter)
        {
            // Uppercase ASCII letters only (matches the A1 addresses this decodes).
            var upper = (char)(c & ~0x20);
            if (upper is < 'A' or > 'Z')
                throw new ArgumentOutOfRangeException(nameof(columnLetter), columnLetter.ToString() + " is not recognized as a column letter");

            result = result * 26 + (upper - 'A' + 1);
        }

        return result;
    }

    /// <summary>
    /// Gets the column letter of a given column number.
    /// </summary>
    /// <param name="columnNumber">The column number to translate into a column letter.</param>
    /// <param name="trimToAllowed">if set to <c>true</c> the column letter will be restricted to the allowed range.</param>
    public static string GetColumnLetterFromNumber(int columnNumber, bool trimToAllowed = false)
    {
        if (trimToAllowed) columnNumber = TrimColumnNumber(columnNumber);

        if (columnNumber <= 0 || columnNumber > AllLetters.Length)
            throw new ArgumentOutOfRangeException(nameof(columnNumber));

        // Adjust for start on column 1
        return AllLetters[columnNumber - 1];
    }

    internal static int TrimColumnNumber(int columnNumber)
    {
        return Math.Max(MinColumnNumber, Math.Min(MaxColumnNumber, columnNumber));
    }

    internal static int TrimRowNumber(int rowNumber)
    {
        return Math.Max(MinRowNumber, Math.Min(MaxRowNumber, rowNumber));
    }

    public static bool IsValidColumn(string column)
    {
        if (string.IsNullOrWhiteSpace(column))
            return false;
        var length = column.Length;
        if (length > 3)
            return false;

        var theColumn = column.ToUpper();

        var isValid = IsUpperLetter(theColumn[0]);
        if (length == 1)
            return isValid;

        if (length == 2)
            return isValid && IsUpperLetter(theColumn[1]);

        return IsValidThreeLetterColumn(theColumn);
    }

    /// <summary>
    /// Whether an upper-case, three-character column name is at most XFD, the last column.
    /// </summary>
    private static bool IsValidThreeLetterColumn(string theColumn)
    {
        if (theColumn[0] >= 'A' && theColumn[0] < 'X')
            return IsUpperLetter(theColumn[1]) && IsUpperLetter(theColumn[2]);

        if (theColumn[0] != 'X') return false;

        if (theColumn[1] < 'F')
            return IsUpperLetter(theColumn[2]);

        if (theColumn[1] != 'F') return false;

        return theColumn[2] >= 'A' && theColumn[2] <= 'D';
    }

    private static bool IsUpperLetter(char c) => c >= 'A' && c <= 'Z';

    public static bool IsValidRow(string rowString)
    {
        // Digits, counted by hand, rather than int.TryParse. A row reference in A1 notation is
        // ASCII digits and nothing else — which is what IsValidRangeAddress's own pattern says
        // with `\d{1,7}` — and every attempt to express that through the number parser let
        // something else in (D41):
        //
        //   int.TryParse(s, out row)                       accepts " 1", "1 " and "+1", because
        //                                                  the default is NumberStyles.Integer
        //   int.TryParse(s, NumberStyles.None, …, out row) still accepts "1\0\0\0", because the
        //                                                  parser stops at a null terminator
        //
        // Both were found by the harness's consistency property, the second one within minutes of
        // the first being fixed. The parser is answering a different question — "is this a number
        // a user might have typed" — so it is the wrong instrument, not one to keep narrowing.
        // IsNullOrEmpty rather than a length read: int.TryParse returned false for null, so
        // reading .Length turned a public predicate that answered "no" into one that threw
        // NullReferenceException. The annotation says non-nullable, but this is a public API and
        // an external caller may not have nullable enabled. IsValidColumn guards the same way.
        // Raised by CodeRabbit's review of this PR.
        if (string.IsNullOrEmpty(rowString) || rowString.Length > MaxRowDigits)
            return false;

        var row = 0;
        foreach (var digit in rowString)
        {
            if (!char.IsAsciiDigit(digit))
                return false;

            row = (row * 10) + (digit - '0');
        }

        return row is > 0 and <= MaxRowNumber;
    }

    /// <summary>
    /// Digits in <see cref="MaxRowNumber"/>, and so the longest a row reference can be. Bounds the
    /// accumulator in <see cref="IsValidRow"/> and <see cref="Point.TryParse"/> well below overflow.
    /// </summary>
    internal const int MaxRowDigits = 7;

    public static bool IsValidA1Address(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
            return false;

        // A '$' is an anchor, and it is only meaningful in two places: before the column letters
        // and before the row digits. Its position has to be checked rather than ignored.
        //
        // This previously did `address.Replace("$", "")` and validated what was left, which
        // accepted a '$' anywhere at all — 'A$2$', 'A2$', '$$A$$2' were all reported as valid A1
        // addresses. IsValidRangeAddress, which matches a real pattern, disagreed on every one of
        // them, so the two predicates contradicted each other about the same string (D36).
        var index = 0;
        var length = address.Length;

        if (address[index] == '$')
            index++;

        // Letters only. Scanning "up to the first digit" would swallow the anchor in 'A$2' into
        // the column and leave IsValidColumn to reject 'A$'.
        var columnStart = index;
        while (index < length && char.IsAsciiLetter(address[index]))
            index++;

        var column = address[columnStart..index];
        if (column.Length == 0)
            return false;

        if (index < length && address[index] == '$')
            index++;

        return index < length
            && IsValidRow(address[index..])
            && IsValidColumn(column);
    }

    public static bool IsValidRCAddress(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
            return false;

        // A scan, not a regex. This was a regex with a 150 ms timeout, and the timeout is
        // wall-clock time: a thread descheduled or paused by GC mid-match threw
        // RegexMatchTimeoutException on a name as plain as 'OpportunityTable'. Table names are
        // checked against this on load, so a busy server failed to open workbooks with tables (#815).
        // The grammar is 'R' then 'C', each optional, each with an optional offset that is
        // either bare or in brackets.
        var index = 0;
        return TrySkipRCPart(address, ref index, 'R')
            && TrySkipRCPart(address, ref index, 'C')
            && index == address.Length;
    }

    /// <summary>
    /// Skips one part of an R1C1 address: <paramref name="letter"/> in either case, then an
    /// optional offset of digits with an optional leading '-', bare or in brackets, as in
    /// <c>R</c>, <c>R12</c>, <c>R-1</c>, <c>R[]</c> and <c>R[-1]</c>.
    /// </summary>
    /// <returns><c>false</c> if the letter is there and its offset is malformed. A part that is
    /// not there at all is skipped and is not an error.</returns>
    private static bool TrySkipRCPart(string address, ref int index, char letter)
    {
        if (index == address.Length || char.ToUpperInvariant(address[index]) != letter)
            return true;

        index++;
        var bracketed = index < address.Length && address[index] == '[';
        if (bracketed)
            index++;

        // A '-' must be followed by a digit, so 'R-' and 'R[-]' are not addresses.
        if (index < address.Length && address[index] == '-')
        {
            if (index + 1 == address.Length || !char.IsDigit(address[index + 1]))
                return false;

            index++;
        }

        // char.IsDigit, not IsAsciiDigit, as the regex's '\d' matched every Unicode digit.
        while (index < address.Length && char.IsDigit(address[index]))
            index++;

        if (!bracketed)
            return true;

        if (index == address.Length || address[index] != ']')
            return false;

        index++;
        return true;
    }

    public static bool IsValidRangeAddress(string rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(rangeAddress))
            return false;

        // A scan, not a regex. This was a regex with a 150 ms timeout, and the timeout is
        // wall-clock time: a thread descheduled or paused by GC mid-match threw
        // RegexMatchTimeoutException on a valid address, from Range("A1:B2") and every other
        // caller (#816). The grammar is an optional sheet name and '!', then one of 'A1',
        // 'A1:B2', '1:2' and 'A:B', each reference with an optional '$' before its letters and
        // before its digits.
        // The regex's '\Z' also let the address end in one line feed. INDIRECT relied on that to
        // read "A1:B2"&CHAR(10) as Excel does, and now drops the whitespace itself (#820).
        var index = 0;
        TrySkipSheetName(rangeAddress, ref index);
        return TrySkipRange(rangeAddress, ref index) && index == rangeAddress.Length;
    }

    /// <summary>
    /// Skips a sheet name and its '!': either written in apostrophes, where an apostrophe in the
    /// name is doubled and the characters <c>[ ] * / \ ? :</c> are not allowed, as in
    /// <c>'Bob''s'!</c>, or a run of word characters, as in <c>Sheet_1!</c>.
    /// </summary>
    /// <returns><c>false</c>, with <paramref name="index"/> unchanged, if there is no sheet name.
    /// The sheet name is optional, so that is not an error here. A malformed one returns <c>false</c> too,
    /// deliberately: it mirrors the regex's optional sheet group, and the range itself can start with
    /// word characters, as in <c>A1:B2</c>, so the range check is what refuses what is left.</returns>
    private static bool TrySkipSheetName(string address, ref int index)
    {
        var position = index;
        if (address[position] == '\'')
        {
            position++;
            var nameStart = position;
            while (position < address.Length)
            {
                var c = address[position];
                if (c == '\'')
                {
                    // A doubled apostrophe is part of the name, a single one closes it. The
                    // closing one has to be followed by '!', so there is nothing to back out of.
                    if (position + 1 == address.Length || address[position + 1] != '\'')
                        break;

                    position += 2;
                }
                else if (c is '[' or ']' or '*' or '/' or '\\' or '?' or ':')
                {
                    return false;
                }
                else
                {
                    position++;
                }
            }

            // The name cannot be empty, and the closing apostrophe has to be there.
            if (position == nameStart || position == address.Length)
                return false;

            position++;
        }
        else
        {
            // Each apostrophe is required or absent as a pair, never one on its own, so
            // 'Data!A1 and Data'!A1 are not addresses (#560). The quoted form of a name of word
            // characters, as in 'Data'!A1, is read by the branch above.
            var nameStart = position;
            while (position < address.Length && IsWordChar(address[position]))
                position++;

            if (position == nameStart)
                return false;
        }

        if (position == address.Length || address[position] != '!')
            return false;

        index = position + 1;
        return true;
    }

    /// <summary>
    /// The characters the regex's <c>\w</c> matched: letters, nonspacing marks, decimal digits
    /// and connector punctuation, in every script.
    /// </summary>
    private static bool IsWordChar(char c) => char.GetUnicodeCategory(c) is
        UnicodeCategory.UppercaseLetter or UnicodeCategory.LowercaseLetter or UnicodeCategory.TitlecaseLetter
        or UnicodeCategory.ModifierLetter or UnicodeCategory.OtherLetter or UnicodeCategory.NonSpacingMark
        or UnicodeCategory.DecimalDigitNumber or UnicodeCategory.ConnectorPunctuation;

    /// <summary>
    /// Skips a reference with no sheet: <c>A1</c>, <c>A1:B2</c>, <c>1:2</c> or <c>A:B</c>, each
    /// reference with an optional '$' before its letters and before its digits.
    /// </summary>
    private static bool TrySkipRange(string address, ref int index)
    {
        SkipDollar(address, ref index);
        if (TrySkipLetters(address, ref index))
        {
            // 'A:B'. A '$' cannot come between the letters and the ':'.
            if (index < address.Length && address[index] == ':')
            {
                index++;
                SkipDollar(address, ref index);
                return TrySkipLetters(address, ref index);
            }

            // 'A1' or 'A1:B2'.
            SkipDollar(address, ref index);
            if (!TrySkipDigits(address, ref index))
                return false;

            // 'A1', unless something follows that the caller will refuse.
            if (index == address.Length || address[index] != ':')
                return true;

            index++;
            SkipDollar(address, ref index);
            if (!TrySkipLetters(address, ref index))
                return false;

            SkipDollar(address, ref index);
            return TrySkipDigits(address, ref index);
        }

        // '1:2'.
        if (!TrySkipDigits(address, ref index)
            || index == address.Length
            || address[index] != ':')
            return false;

        index++;
        SkipDollar(address, ref index);
        return TrySkipDigits(address, ref index);
    }

    private static void SkipDollar(string address, ref int index)
    {
        if (index < address.Length && address[index] == '$')
            index++;
    }

    /// <summary>
    /// Skips one to three ASCII letters, the longest a column can be. A fourth is not skipped, so
    /// what follows it is not valid either.
    /// </summary>
    private static bool TrySkipLetters(string address, ref int index)
    {
        var start = index;
        while (index < address.Length && index - start < 3 && char.IsAsciiLetter(address[index]))
            index++;

        return index > start;
    }

    /// <summary>
    /// Skips one to <see cref="MaxRowDigits"/> digits, in any script as the regex's <c>\d</c> matched
    /// them, the longest a row can be. An eighth is not skipped, so what follows it is not valid either.
    /// </summary>
    private static bool TrySkipDigits(string address, ref int index)
    {
        var start = index;
        while (index < address.Length && index - start < MaxRowDigits && char.IsDigit(address[index]))
            index++;

        return index > start;
    }

    public static bool IsValidRangeAddress(IXLRangeAddress rangeAddress)
    {
        return rangeAddress is
        {
            IsValid: true, FirstAddress: { RowNumber: >= 1, ColumnNumber: >= 1 },
            LastAddress: { RowNumber: <= MaxRowNumber, ColumnNumber: <= MaxColumnNumber }
        }
               && rangeAddress.FirstAddress.RowNumber <= rangeAddress.LastAddress.RowNumber
               && rangeAddress.FirstAddress.ColumnNumber <= rangeAddress.LastAddress.ColumnNumber;
    }

    public static int GetColumnNumberFromAddress(string cellAddressString)
    {
        var rowPos = 0;
        while (cellAddressString[rowPos] > '9')
            rowPos++;

        return GetColumnNumberFromLetter(cellAddressString.AsSpan(0, rowPos));
    }

    /// <summary>
    /// Split one entry of a row or column list, such as "A", "A:C", "3" or "1-3", into its first
    /// and last part. An entry without a ':' or '-' separator is both its first and its last part.
    /// </summary>
    /// <remarks>
    /// Neither part is validated or converted; callers decide whether the parts are letters,
    /// numbers or cell addresses. Anything after a second separator is ignored.
    /// </remarks>
    internal static (string First, string Last) SplitRangePair(string pair)
    {
        if (!pair.Contains(':') && !pair.Contains('-'))
            return (pair, pair);

        var parts = pair.Split(':', '-');
        return (parts[0], parts[1]);
    }

    private static readonly Regex A1RegexRelative = A1RegexRelativeCompiled(); // A:A

    private static string Evaluator(Match match, int row, string column)
    {
        if (match.Groups["one"].Success)
        {
            var split = match.Groups["one"].Value.Split('$');
            if (split.Length == 1) return column + row; // A1
            if (split.Length == 3) return match.Groups["one"].Value; // $A$1
            var a = XLAddress.Create(match.Groups["one"].Value);
            if (split[0].Length == 0) return "$" + a.ColumnLetter + row; // $A1
            return column + "$" + a.RowNumber;
        }

        if (match.Groups["two"].Success)
            return ReplaceGroup(match.Groups["two"].Value, row.ToString());

        return ReplaceGroup(match.Groups["three"].Value, column);
    }

    private static string ReplaceGroup(string value, string item)
    {
        var split = value.Split(':');
        var ret1 = split[0].StartsWith('$') ? split[0] : item;
        var ret2 = split[1].StartsWith('$') ? split[1] : item;
        return ret1 + ":" + ret2;
    }

    internal static string ReplaceRelative(string value, int row, string column)
    {
        var oldValue = ">" + value + "<";
        var newVal = A1RegexRelative.Replace(oldValue, m => Evaluator(m, row, column));
        return newVal.Substring(1, newVal.Length - 2);
    }

    public static bool AreEqual(double d1, double d2)
    {
        return Math.Abs(d1 - d2) < Epsilon;
    }

    public static DateTime GetDate(object v)
    {
        // handle dates
        if (v is DateTime time)
        {
            return time;
        }

        // handle doubles
        if (v is double d && d.IsValidOADateNumber())
        {
            return DateTime.FromOADate(d);
        }

        // handle everything else
        return (DateTime)Convert.ChangeType(v, typeof(DateTime));
    }

    // ReSharper disable once InconsistentNaming
    internal static bool IsValidOADateNumber(this double d)
    {
        return d is >= -657435 and < 2958466;
    }

    /// <summary>
    /// <para>
    /// An alternative to <see cref="TimeSpan.FromDays(double)"/>. In NetFx, it returned a value
    /// rounded to milliseconds. In .Net Core 3.0 the behavior has changed and conversion doesn't
    /// round at all (=precision down to ticks). To avoid problems with a different behavior on
    /// NetFx and Core (saving value 1:12:30 on NetFx machine could become 1:12:29.999999 on Core
    /// one machine), we use instead this method for both runtimes (behaves as on Core).
    /// </para>
    /// <para>
    /// TimeSpan has a resolution of 0.1 us (1.15e-12 as a serial date). ~12 digits of precision
    /// are needed to accurately represent one day as a serial date time in that resolution. Double
    /// has ~15 digits of precision, so it should be able to represent up to ~100 days in a ticks
    /// precision.
    /// </para>
    /// </summary>
    internal static TimeSpan GetTimeSpan(double totalDays)
    {
        var ticks = Math.Round(totalDays * TimeSpan.TicksPerDay, MidpointRounding.AwayFromZero);
        if (ticks is > long.MaxValue or < long.MinValue)
            throw new OverflowException("The serial date time value is too large to be represented in a TimeSpan.");

        return TimeSpan.FromTicks(checked((long)ticks));
    }

    internal static bool ValidateName(string objectType, string newName, string oldName,
        IEnumerable<string> existingNames, out string message)
    {
        if (!ValidateName(objectType, newName, out message))
            return false;

        // Table names are case insensitive
        if (!string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase)
            && existingNames.Contains(newName, StringComparer.OrdinalIgnoreCase))
        {
            message = $"There is already a {objectType} named '{newName}'";
            return false;
        }

        return true;
    }

    internal static bool ValidateName(string objectType, string newName, out string message)
    {
        message = "";
        if (string.IsNullOrWhiteSpace(newName))
        {
            message = $"The {objectType} name '{newName}' is invalid";
            return false;
        }

        var allowedFirstCharacters = new[] { '_', '\\' };
        if (!allowedFirstCharacters.Contains(newName[0]) && !char.IsLetter(newName[0]))
        {
            message = $"The {objectType} name '{newName}' does not begin with a letter, an underscore or a backslash.";
            return false;
        }

        if (newName.Length > 255)
        {
            message = $"The {objectType} name is more than 255 characters";
            return false;
        }

        if (newName.Equals("C", StringComparison.OrdinalIgnoreCase) || newName.Equals("R", StringComparison.OrdinalIgnoreCase))
        {
            message = $"The {objectType} name '{newName}' is invalid";
            return false;
        }

        return true;
    }

    internal static double PixelsToPoints(double pixels, double dpi) => pixels * 72d / dpi;

    internal static double PointsToPixels(double points, double dpi) => points * dpi / 72d;

    /// <summary>
    /// The padding of a column, in pixels: 2 px each side for Calibri 11 at 96 DPI, scaled with the
    /// digit width, plus 1 px for the grid line. <c>2 * ceil(mdw / 4) + 1</c>, 5 px for a 7 px
    /// digit and 7 px for a 9 px one.
    /// </summary>
    /// <param name="mdw">Maximum digit width in whole pixels.</param>
    internal static int ColumnPaddingPx(int mdw) => 2 * (int)Math.Ceiling(mdw / 4.0) + 1;

    /// <summary>
    /// Convert size in pixels to a size in NoC (number of characters).
    /// </summary>
    /// <param name="px">Size in pixels, not necessarily whole.</param>
    /// <param name="mdw">Size of maximum digit width in pixels.</param>
    /// <returns>Size in NoC.</returns>
    internal static double PixelToNoC(double px, int mdw)
    {
        var pp = ColumnPaddingPx(mdw);

        // NoC scales linearly with MDW, if size is at least 1 char (+padding)
        if (px >= (mdw + pp))
            return (px - pp) / mdw;

        // smaller sizes are scaled to the 1 NoC size
        return px / (mdw + pp);
    }

    /// <summary>
    /// Convert size in NoC to size in pixels.
    /// </summary>
    /// <param name="noc">Size in number of characters.</param>
    /// <param name="mdw">Maximum digit width in pixels.</param>
    /// <returns>Size in pixels (not rounded).</returns>
    internal static double NoCToPixels(double noc, int mdw)
    {
        var pp = ColumnPaddingPx(mdw);
        if (noc < 1)
            return noc * (mdw + pp);

        return noc * mdw + pp;
    }

    /// <summary>
    /// The <c>&lt;col width&gt;</c> of a column that shows <paramref name="noc"/> digits: the width
    /// with its padding, in digits of <paramref name="mdw"/>. The padding is
    /// <c>2 * ceil(mdw / 4) + 1</c> px, 5/7 of a digit at 96 DPI and 7/9 at 120 DPI, and below one
    /// digit it scales with the width. Not rounded to whole pixels, unlike Excel (#803).
    /// </summary>
    /// <param name="noc">Width as Excel shows it, in number of characters.</param>
    /// <param name="mdw">Maximum digit width of the workbook's first font, in whole pixels.</param>
    internal static double NoCToFileWidth(double noc, int mdw) => NoCToPixels(noc, mdw) / mdw;

    /// <summary>
    /// The width Excel shows for a <c>&lt;col width&gt;</c> of <paramref name="width"/>: the inverse
    /// of <see cref="NoCToFileWidth"/>. Unlike <see cref="ConvertWidthToNoC"/> it snaps to whole
    /// pixels only a width Excel wrote (#803).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Excel writes the width of a column it draws <c>n</c> px wide as <c>n / mdw</c> truncated to
    /// 1/256: 64 px of 7 px digits is 9.142857, written 9.140625. Such a width loads as the
    /// <c>n</c> px it was truncated from, 59/7 digits as Excel shows it (8.43), not 8.426339.
    /// </para>
    /// <para>
    /// Any other width loads exactly, rounded to the six decimals a save writes a width with, so a
    /// width of up to six decimals that XLibur saved loads as exactly the width that was set: 20 at
    /// 120 DPI is saved as 20.777778, which is 20.0000003 less its padding. A width XLibur saved
    /// loads as Excel shows it when the file holds just what Excel writes, as a whole number of
    /// pixels does. Rarely it is not a whole number: 139.426339 at 96 DPI saves as 140.140625, which
    /// Excel writes for 981 px, so it loads as 139.428571. About 1 in 200,000 widths of six decimals
    /// does this, and no width of two decimals at a digit of 5 to 20 px.
    /// </para>
    /// </remarks>
    /// <param name="width">Width from the source file, padding included.</param>
    /// <param name="mdw">Maximum digit width of the workbook's first font, in whole pixels.</param>
    internal static double FileWidthToNoC(double width, int mdw)
    {
        var px = TryGetExcelWholePixels(width, mdw, out var wholePx) ? wholePx : width * mdw;
        return Math.Round(PixelToNoC(px, mdw), 6);
    }

    /// <summary>
    /// Whether <paramref name="width"/> is the whole-pixel width Excel writes for a column
    /// <paramref name="px"/> pixels wide: <c>px / mdw</c> truncated to 1/256.
    /// </summary>
    private static bool TryGetExcelWholePixels(double width, int mdw, out int px)
    {
        // Scaling by 256 is exact in binary, so a whole 1/256 scales to a whole number. Truncating
        // to 1/256 takes off less than a pixel, so the pixels are the width's, rounded up.
        var in256ths = width * 256;
        var whole256ths = Math.Round(in256ths);
        px = (int)Math.Ceiling(width * mdw);
        return Math.Abs(in256ths - whole256ths) < Epsilon
               && px * 256L / mdw == (long)whole256ths;
    }

    /// <summary>
    /// Maximum digit width (MDW) of a font in whole pixels. MDW is the unit of OOXML column widths
    /// (ECMA-376 Part 1, 18.3.1.13), so Calibri 11pt at 96 DPI is 7 pixels.
    /// </summary>
    /// <remarks>
    /// This is the one place that rounds the engine's fractional width. It rounds half away from
    /// zero, the same as every other pixel snap of a font metric here (<c>GlyphBox.EmSize</c>,
    /// <c>GlyphBox.Descent</c>), and as a renderer snaps a positive width to the pixel grid.
    /// Banker's rounding would make a width of 8.5 px 8 but 7.5 px 8 too. A digit is at least
    /// 1 px: every column width is in digits, and a font engine that reports a digit of under half
    /// a pixel would otherwise give a width of NaN.
    /// </remarks>
    internal static int GetMdw(IXLFontEngine engine, IXLFontBase font, double dpiX)
        => Math.Max(1, engine.GetMaxDigitWidth(font, dpiX).RoundToInt());

    /// <inheritdoc cref="GetMdw(IXLFontEngine, IXLFontBase, double)"/>
    internal static int GetMdw(IXLFontBase font, XLWorkbook workbook)
        => GetMdw(workbook.FontEngine, font, workbook.DpiX);

    /// <summary>
    /// Convert width to pixels.
    /// </summary>
    /// <param name="width">Width from the source file, not NoC that is displayed in Excel as a width.</param>
    /// <param name="mdw"></param>
    /// <returns>Number of pixels.</returns>
    internal static int WidthToPixels(double width, int mdw)
    {
        return (width * mdw).RoundToInt();
    }

    internal static double PixelsToWidth(double width, int mdw)
    {
        return Math.Truncate(width * mdw * 256) / 256d;
    }

    /// <summary>
    /// Convert width (as a multiple of MDWs) into a NoCs (number displayed in Excel).
    /// </summary>
    /// <param name="width">Width in MDWs to convert.</param>
    /// <param name="mdw">Maximum digit width in whole pixels.</param>
    /// <returns>Width as a number of NoC.</returns>
    internal static double ConvertWidthToNoC(double width, int mdw)
    {
        var pixelsWidth = WidthToPixels(width, mdw);
        var columnWidth = PixelToNoC(pixelsWidth, mdw);
        return columnWidth;
    }

    /// <summary>
    /// Convert degrees to radians.
    /// </summary>
    internal static double DegToRad(double angle) => Math.PI * angle / 180.0;

    // The first sheet alternative is a quoted name that doubles each apostrophe it holds, as in
    // 'Bob''s'!A1, which IsValidRangeAddress accepts too. Without it such a reference fell through
    // XLDefinedNames' check for a sheet-qualified address.
    [GeneratedRegex(@"^(('(?<Sheet>([^']|'')+)'|'?(?<Sheet>[^'!]+)'?)!(?<Range>.+))|((?<Table>[^\[]+)\[(?<Column>[^\]]+)\])$",
        RegexOptions.ExplicitCapture | RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex NamedRangeReferenceRegexCompiled();

    [GeneratedRegex(
        @"(?<=\W)(?<one>\$?[a-zA-Z]{1,3}\$?\d{1,7})(?=\W)|(?<=\W)(?<two>\$?\d{1,7}:\$?\d{1,7})(?=\W)|(?<=\W)(?<three>\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})(?=\W)",
        RegexOptions.Compiled)]
    private static partial Regex A1RegexRelativeCompiled();
}
