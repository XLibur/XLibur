using System.Text;
using System.Xml;

namespace XLibur.Utils;

/// <summary>
/// The <c>_xHHHH_</c> escape of the spreadsheet text types: a character XML cannot carry is written
/// as <c>_xHHHH_</c>, and a literal underscore that would otherwise open such an escape is itself
/// written as <c>_x005F_</c>.
/// </summary>
/// <remarks>
/// Decoding scans left to right and consumes each escape whole, so the underscore that closes one
/// escape cannot open the next. Encoding has to agree with that: it looks ahead at the original
/// text without consuming anything, so overlapping lookalikes such as <c>_x005F_x0041_</c> are
/// escaped at every underscore that opens one.
/// </remarks>
internal static class XmlEncoder
{
    private const int EscapeLength = 7; // _xHHHH_
    private const int LongEscapeLength = 11; // _xHHHHHHHH_

    public static string EncodeString(string encodeStr)
    {
        // Fast-path: if the string has no characters that need encoding, return as-is.
        // Most shared strings are plain text (letters, digits, punctuation) and don't
        // need any XML encoding or _xHHHH_ escape processing.
        if (!NeedsEncoding(encodeStr))
            return encodeStr;

        var sb = new StringBuilder(encodeStr.Length + 16);
        var len = encodeStr.Length;

        // Index in sb of the latest underscore written as itself, i.e. one that opens nothing in
        // the original text. If an escape follows within the next five characters it can turn out
        // to open one in the output.
        var literalUnderscore = -1;
        var i = 0;
        while (i < len)
        {
            var currentChar = encodeStr[i];
            if (currentChar == '_')
            {
                if (OpensEscape(encodeStr, i))
                {
                    sb.Append("_x005F_");
                }
                else
                {
                    literalUnderscore = sb.Length;
                    sb.Append('_');
                }

                i++;
            }
            else if (XmlConvert.IsXmlChar(currentChar))
            {
                sb.Append(currentChar);
                i++;
            }
            else if (i + 1 < len && XmlConvert.IsXmlSurrogatePair(encodeStr[i + 1], currentChar))
            {
                sb.Append(currentChar);
                sb.Append(encodeStr[i + 1]);
                i += 2;
            }
            else
            {
                // The escape starts with an underscore, which may close "_xHHHH" written just
                // before it as plain characters into a different escape. Escape that opener.
                if (literalUnderscore >= 0 && EndsWithEscapeBody(sb, literalUnderscore))
                {
                    sb.Remove(literalUnderscore, 1);
                    sb.Insert(literalUnderscore, "_x005F_");
                }

                sb.Append("_x").Append(((int)currentChar).ToString("X4")).Append('_');
                i++;
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Checks whether a string contains any characters that require encoding.
    /// Returns false (no encoding needed) for the common case of plain text.
    /// </summary>
    private static bool NeedsEncoding(string s)
    {
        var len = s.Length;
        var i = 0;
        while (i < len)
        {
            var c = s[i];

            // An underscore that opens a _xHHHH_ pattern needs to be escaped itself.
            if (c == '_' && OpensEscape(s, i))
                return true;

            // Check for non-XML characters that need encoding.
            if (!XmlConvert.IsXmlChar(c))
            {
                // Valid surrogate pair doesn't need encoding.
                if (i + 1 < len && XmlConvert.IsXmlSurrogatePair(s[i + 1], c))
                {
                    i += 2; // Skip the surrogate pair
                    continue;
                }

                return true;
            }

            i++;
        }

        return false;
    }

    /// <summary>
    /// Whether the underscore at <paramref name="index"/> is followed by <c>xHHHH_</c> or
    /// <c>xHHHHHHHH_</c>. ECMA-376 (ST_Xstring) defines only the first; the second is what
    /// <see cref="XmlConvert.DecodeName"/> accepted for a character outside the BMP, and is read
    /// so that text from a producer using it keeps loading.
    /// </summary>
    private static bool OpensEscape(string s, int index) => EscapeLengthAt(s, index) > 0;

    /// <summary>
    /// The length of the escape that starts at the underscore at <paramref name="index"/> (7 or 11),
    /// or 0 when none starts there.
    /// </summary>
    private static int EscapeLengthAt(string s, int index)
    {
        if (index + 1 >= s.Length || s[index + 1] != 'x')
            return 0;

        if (HexRunEndsWithUnderscore(s, index + 2, 4))
            return EscapeLength;

        return HexRunEndsWithUnderscore(s, index + 2, 8) ? LongEscapeLength : 0;
    }

    private static bool HexRunEndsWithUnderscore(string s, int start, int hexCount)
    {
        if (start + hexCount >= s.Length || s[start + hexCount] != '_')
            return false;

        for (var k = 0; k < hexCount; k++)
        {
            if (!IsHexDigit(s[start + k]))
                return false;
        }

        return true;
    }

    /// <summary>
    /// Whether <paramref name="sb"/> ends, after the underscore at <paramref name="underscore"/>, with
    /// exactly <c>xHHHH</c> or <c>xHHHHHHHH</c>: an escape that a following underscore would close.
    /// </summary>
    private static bool EndsWithEscapeBody(StringBuilder sb, int underscore)
    {
        var hexCount = sb.Length - underscore - 2;
        if (hexCount is not (4 or 8) || sb[underscore + 1] != 'x')
            return false;

        for (var k = 0; k < hexCount; k++)
        {
            if (!IsHexDigit(sb[underscore + 2 + k]))
                return false;
        }

        return true;
    }

    private static bool IsHexDigit(char c)
    {
        return (uint)(c - '0') <= 9 || (uint)(c - 'a') <= 5 || (uint)(c - 'A') <= 5;
    }

    /// <summary>
    /// Appends the character an escape stands for. An 8-digit escape that is not a scalar value
    /// (beyond U+10FFFF, or a surrogate) is not one, and is left as the text it is.
    /// </summary>
    private static bool TryDecodeEscape(string s, int index, int escapeLength, StringBuilder sb)
    {
        long code = 0;
        for (var k = index + 2; k < index + escapeLength - 1; k++)
            code = (code << 4) | (uint)HexValue(s[k]);

        if (escapeLength == EscapeLength)
        {
            sb.Append((char)code);
            return true;
        }

        if (code > 0x10FFFF || code is >= 0xD800 and <= 0xDFFF)
            return false;

        sb.Append(char.ConvertFromUtf32((int)code));
        return true;
    }

    private static int HexValue(char c)
    {
        return c <= '9' ? c - '0' : (c | 0x20) - 'a' + 10;
    }

    public static string DecodeString(string? decodeStr)
    {
        if (string.IsNullOrEmpty(decodeStr)) return string.Empty;

        // Fast-path: if the string contains no underscore it cannot have any
        // _xHHHH_ escape sequences, so return as-is.
        // The vast majority of shared strings are plain text.
        var first = decodeStr.IndexOf('_');
        if (first < 0)
            return decodeStr;

        // Strings "escaped" with _X (capital X) are not escapes.
        // Example: _Xceed_Something
        // https://github.com/XLibur/XLibur/issues/1154
        var sb = new StringBuilder(decodeStr.Length);
        sb.Append(decodeStr, 0, first);
        var i = first;
        while (i < decodeStr.Length)
        {
            var c = decodeStr[i];
            var escapeLength = c == '_' ? EscapeLengthAt(decodeStr, i) : 0;
            if (escapeLength > 0 && TryDecodeEscape(decodeStr, i, escapeLength, sb))
            {
                i += escapeLength;
            }
            else
            {
                sb.Append(c);
                i++;
            }
        }

        return sb.ToString();
    }
}
