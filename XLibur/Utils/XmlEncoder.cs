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
                if (literalUnderscore >= 0 && sb.Length - literalUnderscore == EscapeLength - 1
                    && IsEscapeBody(sb, literalUnderscore + 1))
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
    /// Whether the underscore at <paramref name="index"/> is followed by <c>xHHHH_</c>.
    /// </summary>
    private static bool OpensEscape(string s, int index)
    {
        return index + EscapeLength <= s.Length
               && s[index + 1] == 'x'
               && IsHexDigit(s[index + 2]) && IsHexDigit(s[index + 3])
               && IsHexDigit(s[index + 4]) && IsHexDigit(s[index + 5])
               && s[index + 6] == '_';
    }

    /// <summary>
    /// Whether the five characters of <paramref name="sb"/> from <paramref name="start"/> are
    /// <c>xHHHH</c>.
    /// </summary>
    private static bool IsEscapeBody(StringBuilder sb, int start)
    {
        return sb[start] == 'x'
               && IsHexDigit(sb[start + 1]) && IsHexDigit(sb[start + 2])
               && IsHexDigit(sb[start + 3]) && IsHexDigit(sb[start + 4]);
    }

    private static bool IsHexDigit(char c)
    {
        return (uint)(c - '0') <= 9 || (uint)(c - 'a') <= 5 || (uint)(c - 'A') <= 5;
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
            if (c == '_' && OpensEscape(decodeStr, i))
            {
                var code = (HexValue(decodeStr[i + 2]) << 12) | (HexValue(decodeStr[i + 3]) << 8)
                           | (HexValue(decodeStr[i + 4]) << 4) | HexValue(decodeStr[i + 5]);
                sb.Append((char)code);
                i += EscapeLength;
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
