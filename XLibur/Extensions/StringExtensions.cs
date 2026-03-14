#nullable disable


using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace XLibur.Excel;

internal static partial class StringExtensions
{
    private static readonly Regex RegexNewLine = RegexNewLineGenerated();

    public static int CharCount(this string instance, char c)
    {
        return instance.Length - instance.Replace(c.ToString(), "").Length;
    }

    public static string RemoveSpecialCharacters(this string str)
    {
        var sb = new StringBuilder();
        foreach (var c in str.Where(c => char.IsLetterOrDigit(c) || c == '.' || c == '_'))
        {
            sb.Append(c);
        }
        return sb.ToString();
    }

    internal static string EscapeSheetName(this string sheetName)
    {
        if (string.IsNullOrEmpty(sheetName)) return sheetName;

        var needEscape = (!char.IsLetter(sheetName[0]) && sheetName[0] != '_') ||
                         XLHelper.IsValidA1Address(sheetName) ||
                         XLHelper.IsValidRCAddress(sheetName) ||
                         sheetName.Any(c => (char.IsPunctuation(c) && c != '.' && c != '_') ||
                                            char.IsSeparator(c) ||
                                            char.IsControl(c) ||
                                            char.IsSymbol(c));
        return needEscape ? string.Concat('\'', sheetName.Replace("'", "''"), '\'') : sheetName;
    }

    internal static string FixNewLines(this string value)
    {
        return RegexNewLine.Replace(value, Environment.NewLine);
    }

    internal static bool PreserveSpaces(this string value)
    {
        return value.StartsWith(' ') || value.EndsWith(' ') || value.AsSpan().IndexOfAny('\n', '\r', '\t') >= 0;
    }

    internal static string ToCamel(this string value)
    {
        return value.Length switch
        {
            0 => value,
            1 => value.ToLower(),
            _ => value.Substring(0, 1).ToLower() + value.Substring(1)
        };
    }

    internal static string ToProper(this string value)
    {
        return value.Length switch
        {
            0 => value,
            1 => value.ToUpper(),
            _ => value.Substring(0, 1).ToUpper() + value.Substring(1)
        };
    }

    internal static string UnescapeSheetName(this string sheetName)
    {
        return sheetName
            .Trim('\'')
            .Replace("''", "'");
    }

    internal static string WithoutLast(this string value, int length)
    {
        return length < value.Length ? value.Substring(0, value.Length - length) : string.Empty;
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
    /// Is the string a new line of any kind (widnows/unix/mac)?
    /// </summary>
    /// <param name="text">Input text to check for EOL at the beginning.</param>
    /// <param name="length">Length of EOL chars.</param>
    /// <returns>True, if text has EOL at the beginning.</returns>
    internal static bool TrySliceNewLine(this ReadOnlySpan<char> text, out int length)
    {
        if (text.Length >= 2 && text[0] == '\r' && text[1] == '\n')
        {
            length = 2;
            return true;
        }

        if (text.Length >= 1 && (text[0] == '\n' || text[0] == '\r'))
        {
            length = 1;
            return true;
        }

        length = default;
        return false;
    }

    /// <summary>
    /// Convert a magic text to a number, where the first letter is in the highest byte of the number.
    /// </summary>
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

    [GeneratedRegex(@"((?<!\r)\n|\r\n)", RegexOptions.Compiled)]
    private static partial Regex RegexNewLineGenerated();
}
