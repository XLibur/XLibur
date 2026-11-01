using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// Reads the colours of <c>a:CT_ShapeProperties</c> and its effects into <see cref="XLColor"/>:
/// a literal <c>a:srgbClr</c>, or a theme colour from <c>a:schemeClr</c>.
/// </summary>
/// <remarks>
/// The read side of <see cref="ShapePropertiesWriter"/>'s colours, shared by charts and pictures.
/// A colour XLibur does not model reads as <c>null</c>, "no explicit colour", rather than
/// throwing, and is left in the file as it is.
/// </remarks>
internal static class DrawingColorReader
{
    /// <summary>
    /// The colour of the <c>a:solidFill</c> child of <paramref name="parent"/>, or <c>null</c> when
    /// there is none or its colour is neither a literal RGB nor a theme colour.
    /// </summary>
    internal static XLColor? ReadSolidFillColor(OpenXmlCompositeElement? parent)
    {
        var solidFill = parent?.Elements<A.SolidFill>().FirstOrDefault();
        if (solidFill == null)
            return null;

        var rgb = solidFill.Elements<A.RgbColorModelHex>().FirstOrDefault()?.Val?.Value;
        if (IsHexRgb(rgb))
            return XLColor.FromHexRgb(rgb!);

        var scheme = solidFill.Elements<A.SchemeColor>().FirstOrDefault()?.Val;
        if (scheme != null && TryMapThemeColor(scheme.Value, out var themeColor))
            return XLColor.FromTheme(themeColor);

        return null;
    }

    /// <summary>
    /// The colour one colour element names: <c>a:srgbClr</c>, <c>a:schemeClr</c>, <c>a:scrgbClr</c>
    /// (percentages, which Excel writes for an unset shadow colour) or <c>a:sysClr</c> (by its last
    /// computed value). <c>null</c> for anything else.
    /// </summary>
    internal static XLColor? ReadColor(OpenXmlElement? color) => color switch
    {
        A.RgbColorModelHex { Val.Value: var rgb } when IsHexRgb(rgb) => XLColor.FromHexRgb(rgb!),
        A.SchemeColor { Val: { } scheme } when TryMapThemeColor(scheme.Value, out var theme) => XLColor.FromTheme(theme),
        A.RgbColorModelPercentage percentage => XLColor.FromArgb(
            Channel(percentage.RedPortion?.Value), Channel(percentage.GreenPortion?.Value), Channel(percentage.BluePortion?.Value)),
        A.SystemColor { LastColor.Value: var last } when IsHexRgb(last) => XLColor.FromHexRgb(last!),
        _ => null,
    };

    /// <summary>
    /// How transparent a colour element's <c>a:alpha</c> makes it, 0 to 100 percent; 0 when it has
    /// none.
    /// </summary>
    internal static double ReadTransparency(OpenXmlElement? color)
    {
        var alpha = color?.Elements<A.Alpha>().FirstOrDefault()?.Val?.Value;
        return alpha is { } opacity ? 100 - opacity / 1000.0 : 0;
    }

    /// <summary>One channel of <c>a:scrgbClr</c>, in thousandths of a percent, as a byte.</summary>
    private static int Channel(int? portion) =>
        (int)System.Math.Round(System.Math.Clamp(portion ?? 0, 0, 100000) * 255 / 100000.0);

    /// <summary>
    /// Whether the value is a six digit hexadecimal RGB colour. Malformed XML must be read as "no
    /// explicit colour" rather than throwing.
    /// </summary>
    private static bool IsHexRgb(string? value)
    {
        if (value is not { Length: 6 })
            return false;

        foreach (var c in value)
        {
            var isHexDigit = c is >= '0' and <= '9' or >= 'a' and <= 'f' or >= 'A' and <= 'F';
            if (!isHexDigit)
                return false;
        }

        return true;
    }

    /// <summary>
    /// The scheme colours XLibur models, paired with the theme slot each maps onto. The first four
    /// slots answer to two names apiece: <c>bg1</c>/<c>lt1</c> and <c>tx1</c>/<c>dk1</c> are the same
    /// slot addressed from a colour map or from the theme itself, and a producer may write either.
    /// </summary>
    /// <remarks>
    /// A plain array rather than a dictionary: <c>SchemeColorValues</c> is an Open XML SDK 3 enum
    /// value type whose equality is defined by its <c>==</c> operator, not by GetHashCode, so a scan
    /// over sixteen entries is both correct and cheaper than hashing.
    /// </remarks>
    private static readonly (A.SchemeColorValues Scheme, XLThemeColor Theme)[] ThemeColorMap =
    [
        (A.SchemeColorValues.Background1, XLThemeColor.Background1),
        (A.SchemeColorValues.Light1, XLThemeColor.Background1),
        (A.SchemeColorValues.Text1, XLThemeColor.Text1),
        (A.SchemeColorValues.Dark1, XLThemeColor.Text1),
        (A.SchemeColorValues.Background2, XLThemeColor.Background2),
        (A.SchemeColorValues.Light2, XLThemeColor.Background2),
        (A.SchemeColorValues.Text2, XLThemeColor.Text2),
        (A.SchemeColorValues.Dark2, XLThemeColor.Text2),
        (A.SchemeColorValues.Accent1, XLThemeColor.Accent1),
        (A.SchemeColorValues.Accent2, XLThemeColor.Accent2),
        (A.SchemeColorValues.Accent3, XLThemeColor.Accent3),
        (A.SchemeColorValues.Accent4, XLThemeColor.Accent4),
        (A.SchemeColorValues.Accent5, XLThemeColor.Accent5),
        (A.SchemeColorValues.Accent6, XLThemeColor.Accent6),
        (A.SchemeColorValues.Hyperlink, XLThemeColor.Hyperlink),
        (A.SchemeColorValues.FollowedHyperlink, XLThemeColor.FollowedHyperlink),
    ];

    internal static bool TryMapThemeColor(A.SchemeColorValues value, out XLThemeColor themeColor)
    {
        foreach (var (scheme, theme) in ThemeColorMap)
        {
            if (value == scheme)
            {
                themeColor = theme;
                return true;
            }
        }

        themeColor = default;
        return false;
    }
}
