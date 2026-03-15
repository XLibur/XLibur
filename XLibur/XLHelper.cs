using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using XLibur.Extensions;

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

    public static readonly Encoding NoBomUTF8 = new UTF8Encoding(false);

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

    internal static readonly Regex RCSimpleRegex = new(
        @"^(r(((-\d)?\d*)|\[(-\d)?\d*\]))?(c(((-\d)?\d*)|\[(-\d)?\d*\]))?$"
        , RegexOptions.IgnoreCase | RegexOptions.Compiled);

    internal static readonly Regex A1SimpleRegex = new(
        @"\A"
        + "(?<Reference>" // Start Group to pick
        + "(?<Sheet>" // Start Sheet Name, optional
        + "("
        + @"\'([^\[\]\*/\\\?:\']+|\'\')\'"
        // Sheet name with special characters, surrounding apostrophes are required
        + "|"
        + @"\'?\w+\'?" // Sheet name with letters and numbers, surrounding apostrophes are optional
        + ")"
        + "!)?" // End Sheet Name, optional
        + "(?<Range>" // Start range
        + @"\$?[a-zA-Z]{1,3}\$?\d{1,7}" // A1 Address 1
        + @"(?<RangeEnd>:\$?[a-zA-Z]{1,3}\$?\d{1,7})?" // A1 Address 2, optional
        + "|"
        + @"(?<ColumnNumbers>\$?\d{1,7}:\$?\d{1,7})" // 1:1
        + "|"
        + @"(?<ColumnLetters>\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})" // A:A
        + ")" // End Range
        + ")" // End Group to pick
        + @"\Z"
        , RegexOptions.Compiled);

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
        if (string.IsNullOrEmpty(columnLetter)) throw new ArgumentNullException(nameof(columnLetter));

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

        var isValid = theColumn[0] >= 'A' && theColumn[0] <= 'Z';
        if (length == 1)
            return isValid;

        if (length == 2)
            return isValid && theColumn[1] >= 'A' && theColumn[1] <= 'Z';

        if (theColumn[0] >= 'A' && theColumn[0] < 'X')
            return theColumn[1] >= 'A' && theColumn[1] <= 'Z'
                                       && theColumn[2] >= 'A' && theColumn[2] <= 'Z';

        if (theColumn[0] != 'X') return false;

        if (theColumn[1] < 'F')
            return theColumn[2] >= 'A' && theColumn[2] <= 'Z';

        if (theColumn[1] != 'F') return false;

        return theColumn[2] >= 'A' && theColumn[2] <= 'D';
    }

    public static bool IsValidRow(string rowString)
    {
        if (int.TryParse(rowString, out var row))
            return row is > 0 and <= MaxRowNumber;
        return false;
    }

    public static bool IsValidA1Address(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
            return false;

        address = address.Replace("$", "");
        var rowPos = 0;
        var addressLength = address.Length;
        while (rowPos < addressLength && (address[rowPos] > '9' || address[rowPos] < '0'))
            rowPos++;

        return
            rowPos < addressLength
            && IsValidRow(address[rowPos..])
            && IsValidColumn(address[..rowPos]);
    }

    public static bool IsValidRCAddress(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
            return false;

        return RCSimpleRegex.IsMatch(address);
    }

    public static bool IsValidRangeAddress(string rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(rangeAddress))
            return false;

        return A1SimpleRegex.IsMatch(rangeAddress);
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

        return GetColumnNumberFromLetter(cellAddressString.Substring(0, rowPos));
    }

    internal static string[] SplitRange(string range)
    {
        return range.Contains('-') ? range.Replace('-', ':').Split(':') : range.Split(':');
    }

    internal static IXLTableRows InsertRowsWithoutEvents(Func<int, bool, IXLRangeRows> insertFunc,
        XLTableRange tableRange, int numberOfRows,
        bool expandTable)
    {
        var ws = tableRange.Worksheet;
        var rows = new XLTableRows(ws.Style);
        var inserted = insertFunc(numberOfRows, false);
        inserted.ForEach(r => rows.Add(new XLTableRow(tableRange, (XLRangeRow)r)));

        if (expandTable)
            tableRange.Table.ExpandTableRows(numberOfRows);

        return rows;
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

        if (new[] { 'C', 'R' }.Any(c => newName.ToUpper().Equals(c.ToString())))
        {
            message = $"The {objectType} name '{newName}' is invalid";
            return false;
        }

        return true;
    }

    internal static double PixelsToPoints(double pixels, double dpi) => pixels * 72d / dpi;

    internal static double PointsToPixels(double points, double dpi) => points * dpi / 72d;

    /// <summary>
    /// Convert size in pixels to a size in NoC (number of characters).
    /// </summary>
    /// <param name="px">Size in pixels.</param>
    /// <param name="mdw">Size of maximum digit width in pixels.</param>
    /// <returns>Size in NoC.</returns>
    internal static double PixelToNoC(int px, int mdw)
    {
        // Pixel padding. Each side should have 2px for Calibri at 11pt plus 1 pixel for the grid line.
        var pp = 2 * (int)Math.Ceiling(mdw / 4.0) + 1;

        // NoC scales linearly with MDW, if size is at least 1 char (+padding)
        if (px >= (mdw + pp))
            return (px - pp) / (double)mdw;

        // smaller sizes are scaled to the 1 NoC size
        return px / (double)(mdw + pp);
    }

    /// <summary>
    /// Convert size in NoC to size in pixels.
    /// </summary>
    /// <param name="noc">Size in number of characters.</param>
    /// <param name="mdw">Maximum digit width in pixels.</param>
    /// <returns>Size in pixels (not rounded).</returns>
    internal static double NoCToPixels(double noc, int mdw)
    {
        var pp = 2 * (int)Math.Ceiling(mdw / 4.0) + 1;
        if (noc < 1)
            return noc * (mdw + pp);

        return noc * mdw + pp;
    }

    /// <summary>
    /// Convert size in number of characters to pixels.
    /// </summary>
    /// <param name="noc">Width</param>
    /// <param name="font">Font used to determine mdw.</param>
    /// <param name="workbook">Workbook for dpi and graphic engine.</param>
    /// <returns>Width in pixels.</returns>
    internal static int NoCToPixels(double noc, IXLFont font, XLWorkbook workbook)
    {
        var mdw = workbook.GraphicEngine.GetMaxDigitWidth(font, workbook.DpiX).RoundToInt();
        return NoCToPixels(noc, mdw).RoundToInt();
    }

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
    /// <param name="font">Font used to determine MDW.</param>
    /// <param name="workbook">Workbook</param>
    /// <returns>Width as a number of NoC.</returns>
    internal static double ConvertWidthToNoC(double width, IXLFont font, XLWorkbook workbook)
    {
        var mdw = workbook.GraphicEngine.GetMaxDigitWidth(font, workbook.DpiX).RoundToInt();
        var pixelsWidth = WidthToPixels(width, mdw);
        var columnWidth = PixelToNoC(pixelsWidth, mdw);
        return columnWidth;
    }

    /// <summary>
    /// Convert degrees to radians.
    /// </summary>
    internal static double DegToRad(double angle) => Math.PI * angle / 180.0;

    [GeneratedRegex(@"^('?(?<Sheet>[^'!]+)'?!(?<Range>.+))|((?<Table>[^\[]+)\[(?<Column>[^\]]+)\])$",
        RegexOptions.ExplicitCapture | RegexOptions.Compiled | RegexOptions.CultureInvariant)]
    private static partial Regex NamedRangeReferenceRegexCompiled();

    [GeneratedRegex(
        @"(?<=\W)(?<one>\$?[a-zA-Z]{1,3}\$?\d{1,7})(?=\W)|(?<=\W)(?<two>\$?\d{1,7}:\$?\d{1,7})(?=\W)|(?<=\W)(?<three>\$?[a-zA-Z]{1,3}:\$?[a-zA-Z]{1,3})(?=\W)",
        RegexOptions.Compiled)]
    private static partial Regex A1RegexRelativeCompiled();
}
