using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelNumberFormat;
using XLibur.Extensions;
using static XLibur.Excel.CalcEngine.Functions.SignatureAdapter;

#pragma warning disable S1244 // Intentional exact float comparison for Excel formula compatibility

namespace XLibur.Excel.CalcEngine.Functions;

internal static class Text
{
    /// <summary>
    /// Characters 0x80 to 0xFF of win-1252 encoding. Core doesn't include win-1252 encoding,
    /// so keep the conversion table in this string.
    /// </summary>
    private const string Windows1252 =
        "\u20AC\u0081\u201A\u0192\u201E\u2026\u2020\u2021\u02C6\u2030\u0160\u2039\u0152\u008D\u017D\u008F" +
        "\u0090\u2018\u2019\u201C\u201D\u2022\u2013\u2014\u02DC\u2122\u0161\u203A\u0153\u009D\u017E\u0178" +
        "\u00A0\u00A1\u00A2\u00A3\u00A4\u00A5\u00A6\u00A7\u00A8\u00A9\u00AA\u00AB\u00AC\u00AD\u00AE\u00AF" +
        "\u00B0\u00B1\u00B2\u00B3\u00B4\u00B5\u00B6\u00B7\u00B8\u00B9\u00BA\u00BB\u00BC\u00BD\u00BE\u00BF" +
        "\u00C0\u00C1\u00C2\u00C3\u00C4\u00C5\u00C6\u00C7\u00C8\u00C9\u00CA\u00CB\u00CC\u00CD\u00CE\u00CF" +
        "\u00D0\u00D1\u00D2\u00D3\u00D4\u00D5\u00D6\u00D7\u00D8\u00D9\u00DA\u00DB\u00DC\u00DD\u00DE\u00DF" +
        "\u00E0\u00E1\u00E2\u00E3\u00E4\u00E5\u00E6\u00E7\u00E8\u00E9\u00EA\u00EB\u00EC\u00ED\u00EE\u00EF" +
        "\u00F0\u00F1\u00F2\u00F3\u00F4\u00F5\u00F6\u00F7\u00F8\u00F9\u00FA\u00FB\u00FC\u00FD\u00FE\u00FF";

    private static readonly Lazy<Dictionary<int, string>> Windows1252Char = new(static () =>
        Enumerable.Range(0, 0x80).Select(static i => (Char: (char)i, Code: i))
            .Concat(Windows1252.Select(static (c, i) => (Char: c, Code: i + 0x80)))
            .ToDictionary(x => x.Code, x => char.ToString(x.Char)));

    private static readonly Lazy<Dictionary<char, int>> Windows1252Code = new(static () =>
        Windows1252Char.Value.ToDictionary(x => x.Value[0], x => x.Key));

    public static void Register(FunctionRegistry ce)
    {
        ce.RegisterFunction("ARRAYTOTEXT", 1, 2, ArrayToText, FunctionFlags.Range | FunctionFlags.Future,
            AllowRange.Only, 0); // Renders a range or array as text
        ce.RegisterFunction("ASC", 1, 1, Adapt(Asc),
            FunctionFlags
                .Scalar); // Changes full-width (double-byte) English letters or katakana within a character string to half-width (single-byte) characters
        //ce.RegisterFunction("BAHTTEXT	Converts a number to text, using the ß (baht) currency format
        ce.RegisterFunction("CHAR", 1, 1, Adapt(Char),
            FunctionFlags.Scalar); // Returns the character specified by the code number
        ce.RegisterFunction("CLEAN", 1, 1, Adapt(Clean),
            FunctionFlags.Scalar); //	Removes all nonprintable characters from text
        ce.RegisterFunction("CODE", 1, 1, Adapt(Code),
            FunctionFlags.Scalar); // Returns a numeric code for the first character in a text string
        ce.RegisterFunction("CONCAT", 1, 255, Adapt(Concat), FunctionFlags.Future | FunctionFlags.Range,
            AllowRange.All); // Joins several text items into one text item
        ce.RegisterFunction("CONCATENATE", 1, 255, Adapt(Concatenate),
            FunctionFlags.Scalar); //	Joins several text items into one text item
        ce.RegisterFunction("DBCS", 1, 1, Adapt(Dbcs),
            FunctionFlags.Scalar); // Changes half-width (single-byte) characters to full-width (double-byte) characters
        ce.RegisterFunction("DOLLAR", 1, 2, AdaptLastOptional(Dollar, 2),
            FunctionFlags.Scalar); // Converts a number to text, using the $ (dollar) currency format
        ce.RegisterFunction("ENCODEURL", 1, 1, Adapt(EncodeUrl),
            FunctionFlags.Scalar | FunctionFlags.Future); // Percent-encodes a string for use in a URL
        ce.RegisterFunction("EXACT", 2, 2, Adapt(Exact),
            FunctionFlags.Scalar); // Checks to see if two text values are identical
        ce.RegisterFunction("FIND", 2, 3, AdaptLastOptional(Find),
            FunctionFlags.Scalar); //Finds one text value within another (case-sensitive)
        ce.RegisterFunction("FIXED", 1, 3, AdaptLastTwoOptional(Fixed, 2, false),
            FunctionFlags.Scalar); // Formats a number as text with a fixed number of decimals
        //ce.RegisterFunction("JIS	Changes half-width (single-byte) English letters or katakana within a character string to full-width (double-byte) characters
        ce.RegisterFunction("LEFT", 1, 2, AdaptLastOptional(Left, 1),
            FunctionFlags.Scalar); // Returns the leftmost characters from a text value
        //ce.RegisterFunction("LEFTB", 1, 2, AdaptLastOptional(Leftb, 1), FunctionFlags.Scalar); // Returns the leftmost bytes from a text value
        ce.RegisterFunction("LEN", 1, 1, Adapt(Len),
            FunctionFlags.Scalar); //, Returns the number of characters in a text string
        ce.RegisterFunction("LOWER", 1, 1, Adapt(Lower), FunctionFlags.Scalar); //	Converts text to lowercase
        ce.RegisterFunction("MID", 3, 3, Adapt(Mid),
            FunctionFlags
                .Scalar); // Returns a specific number of characters from a text string starting at the position you specify
        ce.RegisterFunction("NUMBERVALUE", 1, 3, AdaptNumberValue(NumberValue),
            FunctionFlags.Scalar | FunctionFlags.Future); // Converts a text argument to a number
        //ce.RegisterFunction("PHONETIC	Extracts the phonetic (furigana) characters from a text string
        ce.RegisterFunction("PROPER", 1, 1, Adapt(Proper),
            FunctionFlags.Scalar); // Capitalizes the first letter in each word of a text value
        ce.RegisterFunction("REPLACE", 4, 4, Adapt(Replace), FunctionFlags.Scalar); // Replaces characters within text
        ce.RegisterFunction("REPT", 2, 2, Adapt(Rept), FunctionFlags.Scalar); // Repeats text a given number of times
        ce.RegisterFunction("RIGHT", 1, 2, AdaptLastOptional(Right, 1),
            FunctionFlags.Scalar); // Returns the rightmost characters from a text value
        ce.RegisterFunction("SEARCH", 2, 3, AdaptLastOptional(Search),
            FunctionFlags.Scalar); // Finds one text value within another (not case-sensitive)
        ce.RegisterFunction("SUBSTITUTE", 3, 4, AdaptSubstitute(Substitute),
            FunctionFlags.Scalar); // Substitutes new text for old text in a text string
        ce.RegisterFunction("T", 1, 1, Adapt(T), FunctionFlags.Range | FunctionFlags.ReturnsArray,
            AllowRange.All); // Converts its arguments to text
        ce.RegisterFunction("TEXT", 2, 2, Adapt(_Text),
            FunctionFlags.Scalar); // Formats a number and converts it to text
        ce.RegisterFunction("TEXTAFTER", 2, 6, TextAfter, FunctionFlags.Range | FunctionFlags.Future,
            AllowRange.Only, 1); // Returns the text after a delimiter
        ce.RegisterFunction("TEXTBEFORE", 2, 6, TextBefore, FunctionFlags.Range | FunctionFlags.Future,
            AllowRange.Only, 1); // Returns the text before a delimiter
        ce.RegisterFunction("TEXTJOIN", 3, 255, Adapt(TextJoin), FunctionFlags.Range | FunctionFlags.Future,
            AllowRange.Except, 0, 1); // Joins text via delimiter
        // AllowRange.All keeps the array-formula engine from broadcasting the arguments element by
        // element: TEXTSPLIT produces the array itself, and reduces its own arguments.
        ce.RegisterFunction("TEXTSPLIT", 2, 6, TextSplit,
            FunctionFlags.Range | FunctionFlags.Future | FunctionFlags.ReturnsArray,
            AllowRange.All); // Splits text into a grid on column and row delimiters
        ce.RegisterFunction("TRIM", 1, 1, Adapt(Trim), FunctionFlags.Scalar); // Removes spaces from text
        ce.RegisterFunction("UNICHAR", 1, 1, Adapt(UniChar),
            FunctionFlags.Scalar | FunctionFlags.Future); // Returns the character for a Unicode code point
        ce.RegisterFunction("UNICODE", 1, 1, Adapt(Unicode),
            FunctionFlags.Scalar | FunctionFlags.Future); // Returns the Unicode code point of the first character
        ce.RegisterFunction("UPPER", 1, 1, Adapt(Upper), FunctionFlags.Scalar); // Converts text to uppercase
        ce.RegisterFunction("VALUE", 1, 1, Adapt(Value), FunctionFlags.Scalar); // Converts a text argument to a number
        ce.RegisterFunction("VALUETOTEXT", 1, 2, ValueToText, FunctionFlags.Range | FunctionFlags.Future,
            AllowRange.Only, 0); // Renders any value as text
    }

    #region Unicode and URL encoding

    private static ScalarValue UniChar(CalcContext ctx, double number)
    {
        var codePoint = (int)Math.Truncate(number);
        if (codePoint < 1 || codePoint > 0x10FFFF)
            return XLError.IncompatibleValue;

        // Lone surrogates are not valid code points, but Excel hands them back as the raw UTF-16
        // unit rather than refusing, and a formula can use that to build a pair by hand.
        if (codePoint is >= 0xD800 and <= 0xDFFF)
            return ((char)codePoint).ToString();

        return char.ConvertFromUtf32(codePoint);
    }

    private static ScalarValue Unicode(CalcContext ctx, string text)
    {
        if (text.Length == 0)
            return XLError.IncompatibleValue;

        // A surrogate pair is one character to Excel, and its code point is the combined value.
        if (char.IsHighSurrogate(text[0]) && text.Length > 1 && char.IsLowSurrogate(text[1]))
            return (double)char.ConvertToUtf32(text[0], text[1]);

        return (double)text[0];
    }

    /// <summary>
    /// Percent-encode for use in a URL. Everything outside the RFC 3986 unreserved set is escaped
    /// as the uppercase hex of its UTF-8 bytes, so "/" and ":" are encoded too — Excel escapes a
    /// whole URL, not just the part after the host.
    /// </summary>
    private static ScalarValue EncodeUrl(CalcContext ctx, string text)
    {
        var sb = new StringBuilder(text.Length);
        foreach (var b in Encoding.UTF8.GetBytes(text))
        {
            var c = (char)b;
            if (c is >= 'A' and <= 'Z' or >= 'a' and <= 'z' or >= '0' and <= '9' or '-' or '_' or '.' or '~')
                sb.Append(c);
            else
                sb.Append('%').Append(b.ToString("X2", CultureInfo.InvariantCulture));
        }

        return sb.ToString();
    }

    /// <summary>
    /// The inverse of <see cref="Asc"/>: half-width ASCII and katakana become their full-width
    /// forms. Like ASC, real Excel only does this when the authoring language is East Asian; the
    /// mapping is applied unconditionally here, which is what makes DBCS(ASC(x)) an identity.
    /// </summary>
#pragma warning disable S3776 // Half-width katakana recombination then a flat per-character mapping
#pragma warning disable S127 // The extra i++ consumes the second half of a combining pair
    private static ScalarValue Dbcs(CalcContext ctx, string text)
    {
        const char dakuten = 'ﾞ';
        const char handakuten = 'ﾟ';
        var inverse = HalfToFullKatakana.Value;

        var sb = new StringBuilder(text.Length);
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];

            // A voiced katakana is written half-width as a base plus a combining mark, so the two
            // have to be recombined into one full-width character before the base is translated.
            if (i + 1 < text.Length)
            {
                var marked = text[i + 1] switch
                {
                    dakuten => inverse.Voiced,
                    handakuten => inverse.SemiVoiced,
                    _ => null,
                };

                if (marked is not null && marked.TryGetValue(c, out var composed))
                {
                    sb.Append(composed);
                    i++;
                    continue;
                }
            }

            if (c is >= '!' and <= '~')
                sb.Append((char)(c - 0x0021 + 0xFF01));
            else if (c == ' ')
                sb.Append('　');
            else if (inverse.Plain.TryGetValue(c, out var katakana))
                sb.Append(katakana);
            else
                sb.Append(c);
        }

        return sb.ToString();
    }
#pragma warning restore S3776

#pragma warning restore S127

    /// <summary>
    /// The half-width to full-width katakana mapping, derived by running every full-width katakana
    /// through <see cref="ToHalfWidth"/> and inverting what comes back. Deriving it rather than
    /// writing a second table by hand is what keeps DBCS and ASC exact inverses of each other.
    /// </summary>
    private static readonly Lazy<KatakanaInverse> HalfToFullKatakana = new(static () =>
    {
        var plain = new Dictionary<char, string>();
        var voiced = new Dictionary<char, string>();
        var semiVoiced = new Dictionary<char, string>();

        for (var codePoint = 0x30A1; codePoint <= 0x30FC; codePoint++)
        {
            var full = ((char)codePoint).ToString();
            var half = ToHalfWidth(full);
            switch (half.Length)
            {
                case 1 when half[0] != codePoint:
                    plain.TryAdd(half[0], full);
                    break;
                case 2:
                    var target = half[1] == 'ﾟ' ? semiVoiced : voiced;
                    target.TryAdd(half[0], full);
                    break;
            }
        }

        return new KatakanaInverse(plain, voiced, semiVoiced);
    });

    private sealed record KatakanaInverse(
        Dictionary<char, string> Plain,
        Dictionary<char, string> Voiced,
        Dictionary<char, string> SemiVoiced);

    #endregion

    #region Splitting on delimiters

    private static AnyValue TextBefore(CalcContext ctx, Span<AnyValue> args)
        => TextAround(ctx, args, before: true);

    private static AnyValue TextAfter(CalcContext ctx, Span<AnyValue> args)
        => TextAround(ctx, args, before: false);

    /// <summary>
    /// TEXTBEFORE/TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end],
    /// [if_not_found]) — cut the text at the nth occurrence of a delimiter and return one side of
    /// it. A negative instance counts occurrences from the end.
    /// </summary>
    private static AnyValue TextAround(CalcContext ctx, Span<AnyValue> args, bool before)
    {
        if (!TryGetText(ctx, args[0], out var text, out var textError))
            return textError;

        if (!TryGetDelimiters(ctx, args[1], out var delimiters, out var delimiterError))
            return delimiterError;

        var instance = 1;
        if (args.Length > 2 && !TryOptionalInt(ctx, args[2], 1, out instance, out var instanceError))
            return instanceError;
        if (instance == 0)
            return XLError.IncompatibleValue;

        var ignoreCase = false;
        if (args.Length > 3 && !TryOptionalFlag(ctx, args[3], out ignoreCase, out var matchModeError))
            return matchModeError;

        var endCounts = false;
        if (args.Length > 4 && !TryOptionalFlag(ctx, args[4], out endCounts, out var matchEndError))
            return matchEndError;

        var notFound = args.Length > 5 ? args[5] : AnyValue.From(XLError.NoValueAvailable);

        var matches = FindDelimiters(text, delimiters, ignoreCase);

        // With match_end set, the very end of the text counts as one more delimiter, which is how
        // TEXTBEFORE(text, delim, -1, , 1) returns the whole text when there is no trailing delimiter.
        if (endCounts)
            matches.Add((text.Length, 0));

        var index = instance > 0 ? instance - 1 : matches.Count + instance;
        if (index < 0 || index >= matches.Count)
            return notFound;

        var (start, length) = matches[index];
        return before ? text[..start] : text[(start + length)..];
    }

    /// <summary>
    /// TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]) —
    /// split into a grid, rows first and then columns within each row, and pad the short rows.
    /// </summary>
#pragma warning disable S3776 // Six optional arguments to read before splitting; each guard is independent
    private static AnyValue TextSplit(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetText(ctx, args[0], out var text, out var textError))
            return textError;

        var hasColumnDelimiters = !IsOmitted(args, 1);
        List<string> columnDelimiters = [];
        if (hasColumnDelimiters && !TryGetDelimiters(ctx, args[1], out columnDelimiters, out var columnError))
            return columnError;

        var hasRowDelimiters = !IsOmitted(args, 2);
        List<string> rowDelimiters = [];
        if (hasRowDelimiters && !TryGetDelimiters(ctx, args[2], out rowDelimiters, out var rowError))
            return rowError;

        if (!hasColumnDelimiters && !hasRowDelimiters)
            return XLError.IncompatibleValue;

        var ignoreEmpty = false;
        if (args.Length > 3 && !TryOptionalFlag(ctx, args[3], out ignoreEmpty, out var ignoreError))
            return ignoreError;

        var ignoreCase = false;
        if (args.Length > 4 && !TryOptionalFlag(ctx, args[4], out ignoreCase, out var matchModeError))
            return matchModeError;

        var padding = args.Length > 5 && !IsOmitted(args, 5)
            ? ToScalar(ctx, args[5])
            : ScalarValue.From(XLError.NoValueAvailable);

        var rows = new List<List<string>>();
        foreach (var line in Split(text, rowDelimiters, ignoreCase, ignoreEmpty))
            rows.Add(Split(line, columnDelimiters, ignoreCase, ignoreEmpty));

        if (ignoreEmpty)
            rows.RemoveAll(static row => row.Count == 0);

        if (rows.Count == 0)
            return XLError.IncompatibleValue;

        var width = 0;
        foreach (var row in rows)
            width = Math.Max(width, row.Count);

        var data = new ScalarValue[rows.Count, Math.Max(width, 1)];
        for (var y = 0; y < rows.Count; y++)
        {
            for (var x = 0; x < data.GetLength(1); x++)
                data[y, x] = x < rows[y].Count ? rows[y][x] : padding;
        }

        return new ConstArray(data);
    }
#pragma warning restore S3776

    /// <summary>Split on any of the delimiters; no delimiters at all leaves the text in one piece.</summary>
    private static List<string> Split(string text, List<string> delimiters, bool ignoreCase, bool ignoreEmpty)
    {
        var pieces = new List<string>();
        if (delimiters.Count == 0)
        {
            pieces.Add(text);
            return pieces;
        }

        var position = 0;
        foreach (var (start, length) in FindDelimiters(text, delimiters, ignoreCase))
        {
            pieces.Add(text[position..start]);
            position = start + length;
        }

        pieces.Add(text[position..]);

        if (ignoreEmpty)
            pieces.RemoveAll(string.IsNullOrEmpty);

        return pieces;
    }

    /// <summary>
    /// Every non-overlapping occurrence of any delimiter, left to right. When two delimiters could
    /// match at the same place the longer one wins, so splitting on both "&lt;br&gt;" and "&lt;b&gt;"
    /// does not leave a stray "r&gt;".
    /// </summary>
#pragma warning disable S3776 // Longest-match-wins delimiter scan; the tie-breaking is the point of the method
    private static List<(int Start, int Length)> FindDelimiters(string text, List<string> delimiters, bool ignoreCase)
    {
        var comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        var matches = new List<(int Start, int Length)>();

        var position = 0;
        while (position <= text.Length)
        {
            var bestStart = -1;
            var bestLength = 0;
            foreach (var delimiter in delimiters)
            {
                if (delimiter.Length == 0)
                    continue;

                var found = text.IndexOf(delimiter, position, comparison);
                if (found < 0)
                    continue;

                if (bestStart < 0 || found < bestStart || (found == bestStart && delimiter.Length > bestLength))
                {
                    bestStart = found;
                    bestLength = delimiter.Length;
                }
            }

            if (bestStart < 0)
                break;

            matches.Add((bestStart, bestLength));
            position = bestStart + bestLength;
        }

        return matches;
    }
#pragma warning restore S3776

    /// <summary>Read a delimiter argument, which Excel lets you write as an array of alternatives.</summary>
    private static bool TryGetDelimiters(CalcContext ctx, in AnyValue value, out List<string> delimiters, out XLError error)
    {
        delimiters = [];
        error = default;

        if (value.TryPickScalar(out var scalar, out _))
        {
            if (!scalar.ToText(ctx.Culture).TryPickT0(out var single, out error))
                return false;

            delimiters.Add(single);
            return true;
        }

        if (!value.TryPickCollectionArray(out var array, ctx))
        {
            error = XLError.IncompatibleValue;
            return false;
        }

        foreach (var item in array!)
        {
            if (item.IsBlank)
                continue;

            if (!item.ToText(ctx.Culture).TryPickT0(out var text, out error))
                return false;

            delimiters.Add(text);
        }

        return true;
    }

    private static bool IsOmitted(Span<AnyValue> args, int index)
        => args.Length <= index || (args[index].TryPickScalar(out var scalar, out _) && scalar.IsBlank);

    private static bool TryOptionalInt(CalcContext ctx, in AnyValue value, int fallback, out int result, out XLError error)
    {
        result = fallback;
        if (!value.TryReduceToScalar(ctx, out var scalar, out error))
            return false;

        if (scalar.IsBlank)
            return true;

        if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out error))
            return false;

        result = (int)Math.Truncate(number);
        return true;
    }

    /// <summary>Read one of the 0/1 mode arguments the modern text functions use as booleans.</summary>
    private static bool TryOptionalFlag(CalcContext ctx, in AnyValue value, out bool flag, out XLError error)
    {
        if (!TryOptionalInt(ctx, value, 0, out var mode, out error))
        {
            flag = false;
            return false;
        }

        if (mode is not (0 or 1))
        {
            flag = false;
            error = XLError.IncompatibleValue;
            return false;
        }

        flag = mode == 1;
        return true;
    }

    /// <summary>
    /// Reduce an argument to the single value a scalar parameter wants. Registering a function with
    /// <see cref="AllowRange.All"/> stops the engine from doing this, which these functions need so
    /// that the array-formula path hands them their arguments whole rather than one element at a
    /// time — so they have to do the reduction themselves.
    /// </summary>
    private static ScalarValue ToScalar(CalcContext ctx, in AnyValue value)
        => value.TryReduceToScalar(ctx, out var scalar, out var error) ? scalar : error;

    private static bool TryGetText(CalcContext ctx, in AnyValue value, out string text, out XLError error)
    {
        return ToScalar(ctx, value).ToText(ctx.Culture).TryPickT0(out text!, out error);
    }

    #endregion

    #region Value rendering

    private static AnyValue ValueToText(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetFormat(ctx, args, 1, out var strict, out var formatError))
            return formatError;

        return Render(ctx, ToScalar(ctx, args[0]), strict);
    }

#pragma warning disable S3776 // Concise and strict rendering differ only in separators, chosen inline per position
    private static AnyValue ArrayToText(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetFormat(ctx, args, 1, out var strict, out var formatError))
            return formatError;

        if (!args[0].TryPickCollectionArray(out var array, ctx))
            return Render(ctx, ToScalar(ctx, args[0]), strict);

        // Concise form is a flat comma-separated list; strict form reproduces the array literal,
        // with commas between columns and semicolons between rows.
        var sb = new StringBuilder();
        if (strict)
            sb.Append('{');

        for (var y = 0; y < array!.Height; y++)
        {
            if (y > 0)
                sb.Append(strict ? ";" : ", ");

            for (var x = 0; x < array.Width; x++)
            {
                if (x > 0)
                    sb.Append(strict ? "," : ", ");

                sb.Append(Render(ctx, array[y, x], strict));
            }
        }

        if (strict)
            sb.Append('}');

        return sb.ToString();
    }
#pragma warning restore S3776

    /// <summary>
    /// Read the shared <c>format</c> argument of VALUETOTEXT and ARRAYTOTEXT: 0 is the concise form
    /// a cell would display, 1 the strict form that could be pasted back into a formula.
    /// </summary>
    private static bool TryGetFormat(CalcContext ctx, Span<AnyValue> args, int index, out bool strict, out XLError error)
    {
        strict = false;
        error = default;
        if (args.Length <= index)
            return true;

        if (!args[index].TryPickScalar(out var scalar, out _))
        {
            error = XLError.IncompatibleValue;
            return false;
        }

        if (scalar.IsBlank)
            return true;

        if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out error))
            return false;

        var format = (int)Math.Truncate(number);
        if (format is not (0 or 1))
        {
            error = XLError.IncompatibleValue;
            return false;
        }

        strict = format == 1;
        return true;
    }

    private static string Render(CalcContext ctx, in ScalarValue value, bool strict)
    {
        if (value.TryPickError(out var error))
            return error.ToDisplayString();

        if (value.TryPickText(out var text, out _))
            return strict ? "\"" + text!.Replace("\"", "\"\"") + "\"" : text!;

        // Numbers, logicals and blanks read the same either way.
        return value.ToText(ctx.Culture).Match(t => t!, e => e.ToDisplayString());
    }

    #endregion

    private static ScalarValue Asc(CalcContext ctx, string text) => ToHalfWidth(text);

    private static string ToHalfWidth(string text)
    {
        // Excel version only works when the authoring language is set to specific languages (e.g., Japanese).
        // Function doesn't do anything when Excel is set to most locales (e.g., English). There is no further
        // info. For practical purposes, it converts full-width characters from Halfwidth and Fullwidth Forms
        // Unicode block to half-width variants.

        // Because fullwidth code points are in the base multilingual plane, I just skip over surrogates.
        // Voiced/semi-voiced katakana map to two half-width chars (base + combining mark),
        // so the result can be longer than the input.
        var sb = new StringBuilder(text.Length);
        foreach (int c in text)
            AppendHalfForm(sb, c);

        return sb.ToString();

        // Per ODS specification https://docs.oasis-open.org/office/v1.2/os/OpenDocument-v1.2-os-part2.html#ASC
        static void AppendHalfForm(StringBuilder sb, int c)
        {
            if (c is >= 0x30A1 and <= 0x30F4)
                AppendKatakanaHalfWidth(sb, c);
            else if (c is >= 0xFF01 and <= 0xFF5E)
                sb.Append((char)(c - 0xFF01 + 0x0021)); // Fullwidth ASCII to ASCII
            else
                sb.Append((char)PunctuationToHalfWidth(c));
        }

        static void AppendKatakanaHalfWidth(StringBuilder sb, int c)
        {
            const char dakuten = '\uFF9E';
            const char handakuten = '\uFF9F';

            switch (c)
            {
                // a-o vowels (ア-オ) and their small forms (ァ-ォ)
                case >= 0x30A1 and <= 0x30AA when c % 2 == 0:
                    sb.Append((char)((c - 0x30A2) / 2 + 0xFF71));
                    break;
                case >= 0x30A1 and <= 0x30AA when c % 2 == 1:
                    sb.Append((char)((c - 0x30A1) / 2 + 0xFF67));
                    break;

                // ka-chi (カ-チ) unvoiced
                case >= 0x30AB and <= 0x30C2 when c % 2 == 1:
                    sb.Append((char)((c - 0x30AB) / 2 + 0xFF76));
                    break;
                // ga-dhi (ガ-ヂ) voiced = base + dakuten
                case >= 0x30AB and <= 0x30C2 when c % 2 == 0:
                    sb.Append((char)((c - 0x30AC) / 2 + 0xFF76));
                    sb.Append(dakuten);
                    break;

                // small tsu (ッ)
                case 0x30C3:
                    sb.Append('\uFF6F');
                    break;

                // tsu-to (ツ-ト) unvoiced
                case >= 0x30C4 and <= 0x30C9 when c % 2 == 0:
                    sb.Append((char)((c - 0x30C4) / 2 + 0xFF82));
                    break;
                // du-do (ヅ-ド) voiced = base + dakuten
                case >= 0x30C4 and <= 0x30C9 when c % 2 == 1:
                    sb.Append((char)((c - 0x30C5) / 2 + 0xFF82));
                    sb.Append(dakuten);
                    break;

                // na-no (ナ-ノ)
                case >= 0x30CA and <= 0x30CE:
                    sb.Append((char)(c - 0x30CA + 0xFF85));
                    break;

                // ha-ho (ハ-ホ) group: unvoiced, voiced (dakuten), semi-voiced (handakuten)
                case >= 0x30CF and <= 0x30DD:
                    AppendHaHoGroup(sb, c, dakuten, handakuten);
                    break;

                // ma-mo (マ-モ)
                case >= 0x30DE and <= 0x30E2:
                    sb.Append((char)(c - 0x30DE + 0xFF8F));
                    break;

                // ya-yo (ヤ-ヨ) and small forms (ャ-ョ)
                case >= 0x30E3 and <= 0x30E8 when c % 2 == 0:
                    sb.Append((char)((c - 0x30E4) / 2 + 0xFF94));
                    break;
                case >= 0x30E3 and <= 0x30E8 when c % 2 == 1:
                    sb.Append((char)((c - 0x30E3) / 2 + 0xFF6C));
                    break;

                // ra-ro (ラ-ロ)
                case >= 0x30E9 and <= 0x30ED:
                    sb.Append((char)(c - 0x30E9 + 0xFF97));
                    break;

                case 0x30EF: sb.Append('\uFF9C'); break; // wa (ワ)
                case 0x30F2: sb.Append('\uFF66'); break; // wo (ヲ)
                case 0x30F3: sb.Append('\uFF9D'); break; // n (ン)

                // vu (ヴ) voiced = ｳ + dakuten
                case 0x30F4:
                    sb.Append('\uFF73');
                    sb.Append(dakuten);
                    break;

                default:
                    sb.Append((char)c);
                    break;
            }
        }

        static void AppendHaHoGroup(StringBuilder sb, int c, char dakuten, char handakuten)
        {
            if (c % 3 == 0)
            {
                sb.Append((char)((c - 0x30CF) / 3 + 0xFF8A));
            }
            else if (c % 3 == 1)
            {
                sb.Append((char)((c - 0x30D0) / 3 + 0xFF8A));
                sb.Append(dakuten);
            }
            else
            {
                sb.Append((char)((c - 0x30D1) / 3 + 0xFF8A));
                sb.Append(handakuten);
            }
        }

        static int PunctuationToHalfWidth(int c) => c switch
        {
            0x2015 => 0xFF70, // HORIZONTAL BAR => HALFWIDTH PROLONGED SOUND MARK
            0x2018 => 0x0060, // LEFT SINGLE QUOTATION MARK => GRAVE ACCENT
            0x2019 => 0x0027, // RIGHT SINGLE QUOTATION MARK => APOSTROPHE
            0x201D => 0x0022, // RIGHT DOUBLE QUOTATION MARK => QUOTATION MARK
            0x3001 => 0xFF64, // IDEOGRAPHIC COMMA
            0x3002 => 0xFF61, // IDEOGRAPHIC FULL STOP
            0x300C => 0xFF62, // LEFT CORNER BRACKET
            0x300D => 0xFF63, // RIGHT CORNER BRACKET
            0x309B => 0xFF9E, // KATAKANA-HIRAGANA VOICED SOUND MARK
            0x309C => 0xFF9F, // KATAKANA-HIRAGANA SEMI-VOICED SOUND MARK
            0x30FB => 0xFF65, // KATAKANA MIDDLE DOT
            0x30FC => 0xFF70, // KATAKANA-HIRAGANA PROLONGED SOUND MARK
            0xFFE5 => 0x005C, // FULLWIDTH YEN SIGN => REVERSE SOLIDUS
            _ => c
        };
    }

    private static ScalarValue Char(double number)
    {
        number = Math.Truncate(number);
        if (number is < 1 or > 255)
            return XLError.IncompatibleValue;

        // Spec says to interpret numbers as values encoded in iso-8859-1. The actual
        // encoding depends on authoring language, e.g. JP uses JIS X 0201. Fun fact,
        // JP has values 253-255 from iso-8859-1, not JIS. EN/CZ/RU uses win-1252.
        // Anyway, there is no way to get a map of all encodings, so let's use one.
        // Win-1252 is probably the best default choice, because this function is
        // pre-unicode and Excel was mostly sold in US/EU.
        var value = checked((int)number);

        return Windows1252Char.Value[value];
    }

    private static ScalarValue Clean(CalcContext ctx, string text)
    {
        // Although the standard says it removes only 0..1F, the real one removes other characters as
        // well. Based on `LEN(CLEAN(UNICHAR(A1))) = 0`, it removes 1-1F and 0x80-0x9F. ODF
        // says to remove Cc and Cn, but Excel doesn't seem to remove Cn.
        var result = new StringBuilder(text.Length);
        foreach (char c in text)
        {
            int codePoint = c;
            if (codePoint is >= 0 and <= 0x1F)
                continue;

            if (codePoint is >= 0x80 and <= 0x9F)
                continue;

            result.Append(c);
        }

        return result.ToString();
    }

    private static ScalarValue Code(CalcContext ctx, string text)
    {
        // CODE should be an inverse function to CHAR
        if (text.Length == 0)
            return XLError.IncompatibleValue;

        if (!Windows1252Code.Value.TryGetValue(text[0], out var code))
            return Windows1252Code.Value['?'];

        return code;
    }

    private static ScalarValue Concat(CalcContext ctx, List<Array> texts)
    {
        var sb = new StringBuilder();
        foreach (var array in texts)
        {
            foreach (var scalar in array)
            {
                ctx.ThrowIfCancelled();
                if (!scalar.ToText(ctx.Culture).TryPickT0(out var text, out var error))
                    return error;

                sb.Append(text);
                if (sb.Length > 32767)
                    return XLError.IncompatibleValue;
            }
        }

        return sb.ToString();
    }

    private static ScalarValue Concatenate(CalcContext ctx, List<string> texts)
    {
        var totalLength = texts.Sum(static x => x.Length);
        var sb = new StringBuilder(totalLength);
        foreach (var text in texts)
        {
            sb.Append(text);
            if (sb.Length > 32767)
                return XLError.IncompatibleValue;
        }

        return sb.ToString();
    }

    private static AnyValue Find(CalcContext ctx, string findText, string withinText, OneOf<double, Blank> startNum)
    {
        var startIndex = startNum.TryPickT0(out var startNumber, out _) ? (int)Math.Truncate(startNumber) - 1 : 0;
        if (startIndex < 0 || startIndex > withinText.Length)
            return XLError.IncompatibleValue;

        var text = withinText.AsSpan(startIndex);
        var index = text.IndexOf(findText.AsSpan());
        return index == -1
            ? XLError.IncompatibleValue
            : index + startIndex + 1;
    }

    private static ScalarValue Fixed(CalcContext ctx, double number, double numDecimals, bool suppressComma)
    {
        numDecimals = Math.Truncate(numDecimals);

        // Excel allows up to 127 decimal digits. The .NET Core 8+ allows it, but older Core and
        // Fx are more limited. To keep code sane, use 99, so N99 formatting string works everywhere.
        if (numDecimals > 99)
            return XLError.IncompatibleValue;

        var culture = ctx.Culture;
        if (suppressComma)
        {
            culture = (CultureInfo)culture.Clone();
            culture.NumberFormat.NumberGroupSeparator = string.Empty;
        }

        var rounded = XLMath.Round(number, numDecimals);

        // Number rounded to tens, hundreds... should be displayed without any decimal places
        var digits = Math.Max(numDecimals, 0);
        return rounded.ToString("N" + digits, culture);
    }

    private static ScalarValue Left(CalcContext ctx, string text, double numChars)
    {
        if (numChars < 0)
            return XLError.IncompatibleValue;

        numChars = Math.Truncate(numChars);
        if (numChars >= text.Length)
            return text;

        // StringInfo.LengthInTextElements returns a length in graphemes, regardless of
        // how is grapheme stored (e.g. denormalized family emoji is 7 code points long,
        // with 4 emoji and 3 zero width joiners).
        // Generally we should return number of codepoints, at least that's how Excel and
        // LibreOffice do it (at least for LEFT).
        var i = 0;
        while (numChars > 0 && i < text.Length)
        {
            // Most C# text API will happily ignore invalid surrogate pairs, so do we
            i += char.IsSurrogatePair(text, i) ? 2 : 1;
            numChars--;
        }

        return text[..i];
    }

    private static ScalarValue Len(CalcContext ctx, string text)
    {
        // Excel counts code units, not codepoints, e.g. it returns 2 for emoji in astral
        // plane. LibreOffice returns 1 and most other functions (e.g. LEFT) use codepoints,
        // not code units. Sanity says count codepoints, but compatibility says code units.
        return text.Length;
    }

    private static ScalarValue Lower(CalcContext ctx, string text)
    {
        // Spec says "by doing a character-by-character conversion"
        // so don't do the whole string at once.
        var sb = new StringBuilder(text.Length);
        for (var i = 0; i < text.Length; ++i)
        {
            var c = text[i];
            char lowercase;
            if (i == text.Length - 1 && c == 'Σ')
            {
                // Spec: when Σ (U+03A3) is found in a word-final position, it is converted
                // to ς (U+03C2) instead of σ (U+03C3).
                lowercase = 'ς';
            }
            else
            {
                lowercase = char.ToLower(c, ctx.Culture);
            }

            sb.Append(lowercase);
        }

        return sb.ToString();
    }

    private static ScalarValue Mid(CalcContext ctx, string text, double startPos, double numChars)
    {
        // Unlike LEFT, MID uses code units and even cuts off half of surrogates,
        // e.g. LEN(MID("😊😊",1,3)) = 3. Also, spec has parameters at wrong places.
        if (startPos is < 1 or >= int.MaxValue + 1d || numChars is < 0 or >= int.MaxValue + 1d)
            return XLError.IncompatibleValue;

        var start = checked((int)Math.Truncate(startPos)) - 1;
        var length = checked((int)Math.Truncate(numChars));
        if (start >= text.Length - 1)
            return string.Empty;

        if (start + length >= text.Length)
            return text[start..];

        return text.Substring(start, length);
    }

    private static ScalarValue Proper(CalcContext ctx, string text)
    {
        if (text.Length == 0)
            return string.Empty;

        var culture = ctx.Culture;
        var sb = new StringBuilder(text.Length);
        var prevWasLetter = false;
        foreach (var c in text)
        {
            var casedChar = prevWasLetter
                ? char.ToLower(c, culture)
                : char.ToUpper(c, culture);
            sb.Append(casedChar);
            prevWasLetter = char.IsLetter(c);
        }

        return sb.ToString();
    }

    private static ScalarValue Replace(CalcContext ctx, string oldText, double startPos, double numChars,
        string replacement)
    {
        if (startPos is < 1 or >= XLHelper.CellTextLimit + 1)
            return XLError.IncompatibleValue;

        if (numChars is < 0 or >= XLHelper.CellTextLimit + 1)
            return XLError.IncompatibleValue;

        var prefixLength = checked((int)startPos) - 1;
        if (prefixLength > oldText.Length)
            prefixLength = oldText.Length;

        var deletedLength = checked((int)numChars);
        if (prefixLength + deletedLength > oldText.Length)
            deletedLength = oldText.Length - prefixLength;

        // Excel does everything is in code units, produces invalid surrogate pairs and everything.
        var sb = new StringBuilder(oldText.Length - deletedLength + replacement.Length);
        var text = oldText.AsSpan();
        sb.Append(text[..prefixLength]);
        sb.Append(replacement);
        sb.Append(text[(prefixLength + deletedLength)..]);

        return sb.ToString();
    }

    private static ScalarValue Rept(string text, double replicationCount)
    {
        if (replicationCount is < 0 or >= int.MaxValue + 1d)
            return XLError.IncompatibleValue;

        // If text is empty, loop could run too many times
        if (text.Length == 0)
            return string.Empty;

        var count = checked((int)replicationCount);
        var resultLength = text.Length * count;
        if (resultLength > XLHelper.CellTextLimit)
            return XLError.IncompatibleValue;

        var sb = new StringBuilder(resultLength);
        for (var i = 0; i < count; ++i)
            sb.Append(text);

        return sb.ToString();
    }

    private static ScalarValue Right(CalcContext ctx, string text, double numChars)
    {
        // Unlike MID, RIGHT uses codepoint semantic
        if (numChars < 0)
            return XLError.IncompatibleValue;

        numChars = Math.Truncate(numChars);
        if (numChars >= text.Length)
            return text;

        var i = text.Length;
        while (numChars > 0 && i > 0)
        {
            i -= i > 1 && char.IsSurrogatePair(text[i - 2], text[i - 1]) ? 2 : 1;
            numChars--;
        }

        return text[i..];
    }

    private static AnyValue Search(CalcContext ctx, string findText, string withinText, OneOf<double, Blank> startNum)
    {
        if (withinText.Length == 0)
            return XLError.IncompatibleValue;

        var startIndex = startNum.TryPickT0(out var startNumber, out _) ? (int)Math.Truncate(startNumber) : 1;
        startIndex -= 1;
        if (startIndex < 0 || startIndex >= withinText.Length)
            return XLError.IncompatibleValue;

        var wildcard = new Wildcard(findText);
        ReadOnlySpan<char> text = withinText.AsSpan().Slice(startIndex);
        var firstIdx = wildcard.Search(text);
        if (firstIdx < 0)
            return XLError.IncompatibleValue;

        return firstIdx + startIndex + 1;
    }

    private static ScalarValue Substitute(CalcContext ctx, string text, string oldText, string newText,
        double? occurrenceOrMissing)
    {
        // Replace is case sensitive
        if (occurrenceOrMissing is < 1 or >= 2147483647)
            return XLError.IncompatibleValue;

        if (text.Length == 0 || oldText.Length == 0)
            return text;

        if (occurrenceOrMissing is null)
            return text.Replace(oldText, newText);

        // There must be at least one loop (>=1), so `pos` will be set to an index or returned as not found
        var pos = -1;
        var occurrence = (int)occurrenceOrMissing.Value;
        for (var i = 0; i < occurrence; ++i)
        {
            pos = text.IndexOf(oldText, pos + 1, StringComparison.Ordinal);
            if (pos < 0)
                return text;
        }

        var textSpan = text.AsSpan();
        var sb = new StringBuilder(text.Length - oldText.Length + newText.Length);
        sb.Append(textSpan[..pos]);
        sb.Append(newText);
        sb.Append(textSpan[(pos + oldText.Length)..]);
        return sb.ToString();
    }

    private static AnyValue T(CalcContext ctx, AnyValue value)
    {
        if (value.TryPickScalar(out var scalar, out var collection))
        {
            if (scalar.TryPickError(out var scalarError))
                return scalarError;

            return scalar.IsText ? scalar.GetText() : string.Empty;
        }

        if (collection.TryPickT0(out var array, out var reference))
            return TArray(ctx, array);

        var area = reference[0];
        var cellValue = ctx.GetCellValue(area.Worksheet, area.FirstAddress.RowNumber, area.FirstAddress.ColumnNumber);
        if (cellValue.TryPickError(out var cellError))
            return cellError;

        return cellValue.IsText ? cellValue.GetText() : string.Empty;
    }

    /// <remarks>
    /// Lazy, for the reason <see cref="Array.Apply(Func{ScalarValue, ScalarValue})"/> is: this used
    /// to fill a <c>ScalarValue[array.Height, array.Width]</c> before returning. <c>T</c> accepts a
    /// range, so a 458-column operand made that 480 million elements — the fuzzer reached it as
    /// <c>T(B1:C1/V+AM/U/+QU:B%+1)</c>, 11 GB and 60 seconds for a result the caller reduced to one
    /// cell (D38).
    /// </remarks>
    private static AnyValue TArray(CalcContext ctx, Array array)
        => array.Apply(element => ToTextElement(ctx, element));

    private static ScalarValue ToTextElement(CalcContext ctx, ScalarValue element)
    {
        ctx.ThrowIfCancelled();
        if (element.TryPickError(out var elementError))
            return elementError;

        return element.IsText ? element.GetText() : string.Empty;
    }

    private static ScalarValue _Text(CalcContext ctx, ScalarValue value, string format)
    {
        // Non-convertible values are turned to string
        if (!value.ToNumber(ctx.Culture).TryPickT0(out var number, out _) || value.IsLogical)
        {
            return value
                .ToText(ctx.Culture)
                .Match<ScalarValue>(static x => x!, static x => x);
        }

        // Library doesn't format whitespace formats
        if (string.IsNullOrWhiteSpace(format))
            return format;

        var nf = new NumberFormat(format);

        // Values formated as date/time must be in the limit for dates
        var isDateFormat = nf.IsDateTimeFormat || nf.IsTimeSpanFormat;
        if (isDateFormat && number < 0 || number >= ctx.DateSystemUpperLimit)
            return XLError.IncompatibleValue;

        try
        {
            return nf.Format(number, ctx.Culture);
        }
        catch
        {
            return XLError.IncompatibleValue;
        }
    }

    private static ScalarValue TextJoin(CalcContext ctx, string delimiter, bool ignoreEmpty, List<AnyValue> texts)
    {
        var first = true;
        var sb = new StringBuilder();
        foreach (var textValue in texts)
        {
            var result = TextJoinAppend(ctx, delimiter, ignoreEmpty, textValue, sb, ref first);
            if (result.TryPickError(out var error))
                return error;
        }

        return sb.ToString();
    }

    private static ScalarValue TextJoinAppend(CalcContext ctx, string delimiter, bool ignoreEmpty, AnyValue textValue,
        StringBuilder sb, ref bool first)
    {
        var textElements = ignoreEmpty
            ? ctx.GetNonBlankValues(textValue)
            : ctx.GetAllValues(textValue);
        foreach (var scalar in textElements)
        {
            ctx.ThrowIfCancelled();
            if (!scalar.ToText(ctx.Culture).TryPickT0(out var text, out var error))
                return error;

            if (ignoreEmpty && text.Length == 0)
                continue;

            if (first)
            {
                sb.Append(text);
                first = false;
            }
            else
            {
                sb.Append(delimiter).Append(text);
            }

            if (sb.Length > XLHelper.CellTextLimit)
                return XLError.IncompatibleValue;
        }

        return ScalarValue.Blank;
    }

    private static ScalarValue Trim(CalcContext ctx, string text)
    {
        const char space = ' ';
        var span = text.AsSpan().Trim(space);
        var sb = new StringBuilder(span.Length);
        var i = 0;
        while (i < span.Length)
        {
            sb.Append(span[i]);
            if (span[i] == space)
            {
                while (i < span.Length - 1 && span[i + 1] == space)
                    i++;
            }

            i++;
        }

        return sb.ToString();
    }

    private static ScalarValue Upper(CalcContext ctx, string text)
    {
        return text.ToUpper(ctx.Culture);
    }

    private static AnyValue Value(CalcContext ctx, ScalarValue arg)
    {
        // Specification is vague/misleading:
        // * function accepts significantly more diverse range of inputs e.g. result of "($100)" is -100
        //   despite braces not being part of any default number format.
        // * Different cultures work weird, e.g. 7:30 PM is detected as 19:30 in cs locale despite "PM" designator being "odp."
        // * Formats 14 and 22 differ depending on the locale (that is why in dialogue are with a '*' sign)
        if (arg.IsBlank)
            return 0;

        if (arg.TryPickNumber(out var number))
            return number;

        if (!arg.TryPickText(out var text, out var error))
            return error;

        const string percentSign = "%";
        var isPercent = text!.Contains(percentSign, StringComparison.Ordinal);
        var textWithoutPercent = isPercent ? text.Replace(percentSign, string.Empty) : text;
        if (double.TryParse(textWithoutPercent, NumberStyles.Any, ctx.Culture, out var parsedNumber))
            return isPercent ? parsedNumber / 100d : parsedNumber;

        // fraction isn't parsed, maybe in the future
        // No idea how Date/Time parsing works, good enough for initial approach
        var dateTimeFormats = new[]
        {
            ctx.Culture.DateTimeFormat.ShortDatePattern,
            ctx.Culture.DateTimeFormat.YearMonthPattern,
            ctx.Culture.DateTimeFormat.ShortTimePattern,
            ctx.Culture.DateTimeFormat.LongTimePattern,
            "mm-dd-yy", // format 14
            "d-MMMM-yy", // format 15
            "d-MMMM", // format 16
            "d-MMM-yyyy",
            "H:mm", // format 20
            "H:mm:ss" // format 21
        };
        const DateTimeStyles dateTimeStyle = DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault;
        if (DateTime.TryParseExact(text, dateTimeFormats, ctx.Culture, dateTimeStyle, out var parsedDate))
            return parsedDate.ToOADate();

        return XLError.IncompatibleValue;
    }

    private static ScalarValue NumberValue(CalcContext ctx, string text, string decimalSeparator, string groupSeparator)
    {
        if (decimalSeparator.Length == 0)
            return XLError.IncompatibleValue;

        if (groupSeparator.Length == 0)
            return XLError.IncompatibleValue;

        if (text.Length == 0)
            return 0;

        var decimalSep = decimalSeparator[0];
        var groupSep = groupSeparator[0];
        if (decimalSep == groupSep)
            return XLError.IncompatibleValue;

        // Protect against taking up too much stack in stackalloc
        if (text.Length >= 256)
            return XLError.IncompatibleValue;

        // Process by ODF specification. Add one character for optional 0 before decimal.
        Span<char> textSpan = stackalloc char[text.Length + 1];
        var newLength = NormalizeNumberText(text, textSpan, decimalSep, groupSep);

        if (textSpan.Length > 0 && textSpan[0] == '.')
        {
            textSpan[..newLength].CopyTo(textSpan[1..]);
            textSpan[0] = '0';
            newLength++;
        }

        textSpan = textSpan[..newLength];

        // Count percent signs at the end
        var percentCount = 0;
        while (textSpan.Length > 0 && textSpan[^1] == '%')
        {
            textSpan = textSpan[..^1];
            percentCount++;
        }

        if (!double.TryParse(textSpan.ToString(), NumberStyles.Float | NumberStyles.AllowParentheses,
                CultureInfo.InvariantCulture, out var number))
            return XLError.IncompatibleValue;

        return ValidateNumberValue(number, percentCount);
    }

    private static int NormalizeNumberText(string text, Span<char> textSpan, char decimalSep, char groupSep)
    {
        var newLength = 0;
        var decimalSeen = false;
        foreach (var c in text)
        {
            if (c == decimalSep)
            {
                textSpan[newLength++] = !decimalSeen ? '.' : c;
                decimalSeen = true;
            }
            else if (c == groupSep && !decimalSeen)
            {
                // Skip all group separators before first encounter of decimal one
            }
            else if (!char.IsWhiteSpace(c))
            {
                textSpan[newLength++] = c;
            }
        }

        return newLength;
    }

    private static ScalarValue ValidateNumberValue(double number, int percentCount)
    {
        if (double.IsInfinity(number))
            return XLError.NumberInvalid;

        for (var i = 0; i < percentCount; ++i)
            number /= 100.0;

        if (number is <= -1e308 or >= 1e308)
            return XLError.IncompatibleValue;

        if (number is >= -1e-309 and <= 1e-309 && number != 0)
            return XLError.IncompatibleValue;

        if (number is >= -1e-308 and <= 1e-308)
            number = 0d;

        return number;
    }

    private static ScalarValue Dollar(CalcContext ctx, double number, double decimals)
    {
        // Excel has limit of 127 decimal places, but C# has limit of 99.
        decimals = Math.Truncate(decimals);
        if (decimals > 99)
            return XLError.IncompatibleValue;

        if (decimals >= 0)
            return number.ToString("C" + decimals, ctx.Culture);

        var factor = Math.Pow(10, -decimals);
        var rounded = Math.Round(number / factor, 0, MidpointRounding.AwayFromZero);
        if (rounded != 0)
            rounded *= factor;

        return rounded.ToString("C0", ctx.Culture);
    }

    private static ScalarValue Exact(string lhs, string rhs)
    {
        return lhs == rhs;
    }
}
