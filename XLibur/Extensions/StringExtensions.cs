using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using XLibur.Excel;
using XLibur.Parser;

namespace XLibur.Extensions;

internal static partial class StringExtensions
{
    private static readonly Regex RegexNewLine = RegexNewLineGenerated();

    extension(string instance)
    {
        public int CharCount(char c)
        {
            return instance.AsSpan().Count(c);
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

        /// <summary>
        /// The sheet name as a formula writes it before the <c>!</c>: quoted when the parser's
        /// <see cref="NameUtils.ShouldQuote"/> says so, with its apostrophes doubled. The parser's rule
        /// is the one Excel stores, so every formula XLibur writes names a sheet the same way (#651).
        /// </summary>
        internal string EscapeSheetName()
        {
            if (string.IsNullOrEmpty(instance)) return instance;

            if (!NameUtils.ShouldQuote(instance.AsSpan()))
                return instance;

            var escaped = instance.Contains('\'') ? instance.Replace("'", "''") : instance;
            return string.Concat("'", escaped, "'");
        }

        internal string FixNewLines()
        {
            // The regex only ever matches sequences containing '\n'; skip its match machinery when there is none.
            return instance.AsSpan().IndexOf('\n') < 0
                ? instance
                : RegexNewLine.Replace(instance, Environment.NewLine);
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
        var i = 0;
        while (i < text.Length)
        {
            if (i + 1 < text.Length && char.IsSurrogatePair(text[i], text[i + 1]))
            {
                output[j] = char.ConvertToUtf32(text[i], text[i + 1]);
                i += 2;
            }
            else
            {
                output[j] = text[i];
                i++;
            }

            j++;
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
            throw new ArgumentException("Magic text must be at most 4 characters.", nameof(magic));
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
