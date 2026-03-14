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
        encodeStr = xHHHHRegex.Replace(encodeStr, "_x005F_$1_");

        var sb = new StringBuilder(encodeStr.Length);
        var len = encodeStr.Length;
        for (var i = 0; i < len; ++i)
        {
            var currentChar = encodeStr[i];
            if (XmlConvert.IsXmlChar(currentChar))
            {
                sb.Append(currentChar);
            }
            else if (i + 1 < len && XmlConvert.IsXmlSurrogatePair(encodeStr[i + 1], currentChar))
            {
                sb.Append(currentChar);
                sb.Append(encodeStr[++i]);
            }
            else
            {
                sb.Append(XmlConvert.EncodeName(currentChar.ToString()));
            }
        }

        return sb.ToString();
    }

    public static string DecodeString(string? decodeStr)
    {
        if (string.IsNullOrEmpty(decodeStr)) return string.Empty;

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
