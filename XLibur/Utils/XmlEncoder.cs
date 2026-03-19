using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace XLibur.Utils;

internal static partial class XmlEncoder
{
    private static readonly Regex xHHHHRegex = xHHHHRegexCompiled();
    private static readonly Regex Uppercase_X_HHHHRegex = Uppercase_X_HHHHRegexCompiled();

    public static string EncodeString(string encodeStr)
    {
        // Fast-path: if the string has no characters that need encoding, return as-is.
        // Most shared strings are plain text (letters, digits, punctuation) and don't
        // need any XML encoding or _xHHHH_ escape processing.
        if (!NeedsEncoding(encodeStr))
            return encodeStr;

        encodeStr = xHHHHRegex.Replace(encodeStr, "_x005F_$1_");

        var sb = new StringBuilder(encodeStr.Length);
        var len = encodeStr.Length;
        var i = 0;
        while (i < len)
        {
            var currentChar = encodeStr[i];
            if (XmlConvert.IsXmlChar(currentChar))
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
                sb.Append(XmlConvert.EncodeName(currentChar.ToString()));
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

            // Check for _xHHHH_ escape pattern that needs to be escaped itself.
            // Pattern: underscore followed by 'x', 4 hex digits, underscore.
            if (c == '_' && i + 6 < len && s[i + 1] == 'x' && s[i + 6] == '_'
                && IsHexDigit(s[i + 2]) && IsHexDigit(s[i + 3])
                && IsHexDigit(s[i + 4]) && IsHexDigit(s[i + 5]))
            {
                return true;
            }

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

    private static bool IsHexDigit(char c)
    {
        return (uint)(c - '0') <= 9 || (uint)(c - 'a') <= 5 || (uint)(c - 'A') <= 5;
    }

    public static string DecodeString(string? decodeStr)
    {
        if (string.IsNullOrEmpty(decodeStr)) return string.Empty;

        // Fast-path: if the string contains no underscore it cannot have any
        // _xHHHH_ or _XHHHH_ escape sequences, so return as-is.
        // The vast majority of shared strings are plain text.
        if (!decodeStr.Contains('_'))
            return decodeStr;

        // Strings "escaped" with _X (capital X) should not be treated as escaped
        // Example: _Xceed_Something
        // https://github.com/XLibur/XLibur/issues/1154
        decodeStr = Uppercase_X_HHHHRegex.Replace(decodeStr, "_x005F_$1_");

        return XmlConvert.DecodeName(decodeStr);
    }

    [GeneratedRegex("_(x[\\dA-Fa-f]{4})_", RegexOptions.Compiled)]
    private static partial Regex xHHHHRegexCompiled();
    [GeneratedRegex("_(X[\\dA-Fa-f]{4})_", RegexOptions.Compiled)]
    private static partial Regex Uppercase_X_HHHHRegexCompiled();
}
