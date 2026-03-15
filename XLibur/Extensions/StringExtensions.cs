using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using XLibur.Excel;

namespace XLibur.Extensions;

internal static partial class StringExtensions
{
    private static readonly Regex RegexNewLine = RegexNewLineGenerated();

    extension(string instance)
    {
        public int CharCount(char c)
        {
            return instance.Length - instance.Replace(c.ToString(), "").Length;
        }

        public string RemoveSpecialCharacters()
        {
            var sb = new StringBuilder();
            foreach (var c in instance.Where(c => char.IsLetterOrDigit(c) || c == '.' || c == '_'))
            {
                sb.Append(c);
            }
            return sb.ToString();
        }

        internal string EscapeSheetName()
        {
            if (string.IsNullOrEmpty(instance)) return instance;

            var needEscape = (!char.IsLetter(instance[0]) && instance[0] != '_') ||
                             XLHelper.IsValidA1Address(instance) ||
                             XLHelper.IsValidRCAddress(instance) ||
                             StartsLikeCellReference(instance) ||
                             instance.Any(c => (char.IsPunctuation(c) && c != '.' && c != '_') ||
                                               char.IsSeparator(c) ||
                                               char.IsControl(c) ||
                                               char.IsSymbol(c));
            return needEscape ? string.Concat('\'', instance.Replace("'", "''"), '\'') : instance;
        }

        internal string FixNewLines()
        {
            return RegexNewLine.Replace(instance, Environment.NewLine);
        }

        internal bool PreserveSpaces()
        {
            return instance.StartsWith(' ') || instance.EndsWith(' ') || instance.AsSpan().IndexOfAny('\n', '\r', '\t') >= 0;
        }

        internal string ToCamel()
        {
            return instance.Length switch
            {
                0 => instance,
                1 => instance.ToLower(),
                _ => string.Concat(instance[..1].ToLower(), instance.AsSpan(1))
            };
        }

        internal string ToProper()
        {
            return instance.Length switch
            {
                0 => instance,
                1 => instance.ToUpper(),
                _ => instance[..1].ToUpper() + instance[1..]
            };
        }

        internal string UnescapeSheetName()
        {
            return instance
                .Trim('\'')
                .Replace("''", "'");
        }

        internal string WithoutLast(int length)
        {
            return length < instance.Length ? instance[..^length] : string.Empty;
        }
    }

    /// <summary>
    /// Convert a string (containing code units) into code points.
    /// Surrogate pairs of code units are joined to code points.
    /// </summary>
    /// <param name="text">UTF-16 code units to convert.</param>
    /// <param name="output">Output containing code points. Must always be able to fit whole <paramref name="text"/>.</param>
    /// <returns>Number of code points in the <paramref name="output"/>.</returns>
    internal static int ToCodePoints(this ReadOnlySpan<char> text, Span<int> output)
    {
        var j = 0;
        for (var i = 0; i < text.Length; ++i, ++j)
        {
            if (i + 1 < text.Length && char.IsSurrogatePair(text[i], text[i + 1]))
            {
                output[j] = char.ConvertToUtf32(text[i], text[i + 1]);
                i++;
            }
            else
            {
                output[j] = text[i];
            }
        }

        return j;
    }

    /// <summary>
    /// Is the string a new line of any kind (widnows/unix/Mac)?
    /// </summary>
    /// <param name="text">Input text to check for EOL at the beginning.</param>
    /// <param name="length">Length of EOL chars.</param>
    /// <returns>True, if the text has EOL at the beginning.</returns>
    internal static bool TrySliceNewLine(this ReadOnlySpan<char> text, out int length)
    {
        switch (text.Length)
        {
            case >= 2 when text[0] == '\r' && text[1] == '\n':
                length = 2;
                return true;
            case >= 1 when (text[0] == '\n' || text[0] == '\r'):
                length = 1;
                return true;
            default:
                length = 0;
                return false;
        }
    }

    /// <summary>
    /// Convert a magic text to a number, where the first letter is in the highest byte of the number.
    /// </summary>
    /// <exception cref="ArgumentException"></exception>
    internal static uint ToMagicNumber(this string magic)
    {
        if (magic.Length > 4)
        {
            throw new ArgumentException();
        }

        return Encoding.ASCII.GetBytes(magic).Select(x => (uint)x).Aggregate((acc, cur) => acc * 256 + cur);
    }

    internal static string TrimFormulaEqual(this string text)
    {
        var trimmed = text.AsSpan().Trim();
        if (trimmed.Length > 1 && trimmed[0] == '=')
            return trimmed[1..].TrimStart().ToString();

        return text;
    }

    /// <summary>
    /// Does the name start with a pattern that could be confused with a cell reference?
    /// E.g. "C05A" starts with column "C" followed by digit "0", which is ambiguous
    /// to Excel's formula parser even though "C05A" isn't a complete valid A1 address.
    /// </summary>
    private static bool StartsLikeCellReference(string name)
    {
        var i = 0;
        while (i < name.Length && char.IsLetter(name[i]))
            i++;

        // 1-3 letters followed by at least one digit, where the letters form a valid column
        return i >= 1 && i <= 3 && i < name.Length && char.IsDigit(name[i])
            && XLHelper.IsValidColumn(name[..i]);
    }

    [GeneratedRegex(@"((?<!\r)\n|\r\n)", RegexOptions.Compiled)]
    private static partial Regex RegexNewLineGenerated();
}
