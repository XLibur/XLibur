using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using XLibur.Excel.IO;
using XLibur.Extensions;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

#pragma warning disable S1244 // Intentional exact float comparison for Excel formula compatibility

namespace XLibur.Utils;

internal static class OpenXmlHelper
{
    #region Public Methods

    /// <summary>
    /// Convert color in XLibur representation to specified OpenXML type.
    /// </summary>
    /// <typeparam name="T">The descendant of <see cref="ColorType"/>.</typeparam>
    /// <param name="openXMLColor">The existing instance of ColorType.</param>
    /// <param name="xlColor">Color in XLibur format.</param>
    /// <param name="isDifferential">Flag specifying that the color should be saved in
    /// differential format (affects the transparent color processing).</param>
    /// <returns>The original color in OpenXML format.</returns>
    public static T FromXLiburColor<T>(this ColorType openXMLColor, XLColor xlColor, bool isDifferential = false)
        where T : ColorType
    {
        var adapter = new ColorTypeAdapter(openXMLColor);
        FillFromXLiburColor(adapter, xlColor, isDifferential);
        return (T)adapter.ColorType;
    }

    /// <summary>
    /// Convert color in XLibur representation to specified OpenXML type.
    /// </summary>
    /// <typeparam name="T">The descendant of <see cref="X14.ColorType"/>.</typeparam>
    /// <param name="openXMLColor">The existing instance of ColorType.</param>
    /// <param name="xlColor">Color in XLibur format.</param>
    /// <param name="isDifferential">Flag specifying that the color should be saved in
    /// differential format (affects the transparent color processing).</param>
    /// <returns>The original color in OpenXML format.</returns>
    public static T FromXLiburColor<T>(this X14.ColorType openXMLColor, XLColor xlColor, bool isDifferential = false)
        where T : X14.ColorType
    {
        var adapter = new X14ColorTypeAdapter(openXMLColor);
        FillFromXLiburColor(adapter, xlColor, isDifferential);
        return (T)adapter.ColorType;
    }

    public static BooleanValue? GetBooleanValue(bool value, bool? defaultValue = null)
    {
        return (defaultValue.HasValue && value == defaultValue.Value) ? null : new BooleanValue(value);
    }

    public static bool GetBooleanValueAsBool(BooleanValue? value, bool defaultValue)
    {
        return (value?.HasValue ?? false) ? value.Value : defaultValue;
    }

    /// <summary>
    /// Convert color in OpenXML representation to XLibur type.
    /// </summary>
    /// <param name="openXMLColor">Color in OpenXML format.</param>
    /// <returns>The color in XLibur format.</returns>
    public static XLColor ToXLiburColor(this ColorType openXMLColor)
    {
        return ConvertToXLiburColor(new ColorTypeAdapter(openXMLColor));
    }

    /// <summary>
    /// Convert color in OpenXML representation to XLibur type.
    /// </summary>
    /// <param name="openXMLColor">Color in OpenXML format.</param>
    /// <returns>The color in XLibur format.</returns>
    public static XLColor ToXLiburColor(this X14.ColorType openXMLColor)
    {
        return ConvertToXLiburColor(new X14ColorTypeAdapter(openXMLColor));
    }

    internal static void LoadNumberFormat(NumberingFormat? nfSource, IXLNumberFormat nf)
    {
        if (nfSource == null) return;

        if (nfSource.NumberFormatId != null && nfSource.NumberFormatId.Value < XLConstants.NumberOfBuiltInStyles)
            nf.NumberFormatId = (int)nfSource.NumberFormatId.Value;
        else if (nfSource.FormatCode != null)
            nf.Format = nfSource.FormatCode.Value!;
    }

    internal static void LoadAlignment(Alignment? alignmentSource, IXLAlignment alignment)
    {
        if (alignmentSource is null) return;

        var horizontal = alignmentSource.Horizontal.ToXLiburOrNull();
        if (horizontal is not null)
            alignment.Horizontal = horizontal.Value;
        var vertical = alignmentSource.Vertical.ToXLiburOrNull();
        if (vertical is not null)
            alignment.Vertical = vertical.Value;
        if (alignmentSource.Indent?.Value is not null)
            alignment.Indent = checked((int)alignmentSource.Indent.Value);
        if (alignmentSource.ReadingOrder?.Value is not null)
            alignment.ReadingOrder = alignmentSource.ReadingOrder.Value.ToXLibur();
        if (alignmentSource.WrapText?.Value is not null)
            alignment.WrapText = alignmentSource.WrapText.Value;
        if (alignmentSource.TextRotation is not null)
            alignment.TextRotation = GetXLiburTextRotation(alignmentSource);
        if (alignmentSource.ShrinkToFit?.Value is not null)
            alignment.ShrinkToFit = alignmentSource.ShrinkToFit.Value;
        if (alignmentSource.RelativeIndent?.Value is not null)
            alignment.RelativeIndent = alignmentSource.RelativeIndent.Value;
        if (alignmentSource.JustifyLastLine?.Value is not null)
            alignment.JustifyLastLine = alignmentSource.JustifyLastLine.Value;
    }

    internal static void LoadBorder(Border? borderSource, IXLBorder border)
    {
        if (borderSource == null) return;

        LoadBorderValues(borderSource.DiagonalBorder, border.SetDiagonalBorder, border.SetDiagonalBorderColor);

        if (borderSource.DiagonalUp != null)
            border.DiagonalUp = borderSource.DiagonalUp.Value;
        if (borderSource.DiagonalDown != null)
            border.DiagonalDown = borderSource.DiagonalDown.Value;

        LoadBorderValues(borderSource.LeftBorder, border.SetLeftBorder, border.SetLeftBorderColor);
        LoadBorderValues(borderSource.RightBorder, border.SetRightBorder, border.SetRightBorderColor);
        LoadBorderValues(borderSource.TopBorder, border.SetTopBorder, border.SetTopBorderColor);
        LoadBorderValues(borderSource.BottomBorder, border.SetBottomBorder, border.SetBottomBorderColor);
    }

    private static void LoadBorderValues(BorderPropertiesType? source, Func<XLBorderStyleValues, IXLStyle> setBorder, Func<XLColor, IXLStyle> setColor)
    {
        if (source != null)
        {
            if (source.Style != null)
                setBorder(source.Style.Value.ToXLibur());
            if (source.Color != null)
                setColor(source.Color.ToXLiburColor());
        }
    }

    // Differential fills store the patterns differently than other fills. Actually,
    //  differential fills make more sense. bg is bg, and fg is fg
    // 'Other' fills store the bg color in the fg field when the pattern type is solid
    internal static void LoadFill(Fill? openXMLFill, IXLFill XLiburFill, bool differentialFillFormat)
    {
        if (openXMLFill?.PatternFill == null) return;

        if (openXMLFill.PatternFill.PatternType != null)
            XLiburFill.PatternType = openXMLFill.PatternFill.PatternType.Value.ToXLibur();
        else
            XLiburFill.PatternType = XLFillPatternValues.Solid;

        switch (XLiburFill.PatternType)
        {
            case XLFillPatternValues.None:
                break;

            case XLFillPatternValues.Solid:
                LoadSolidFill(openXMLFill.PatternFill, XLiburFill, differentialFillFormat);
                break;

            default:
                LoadPatternedFill(openXMLFill.PatternFill, XLiburFill);
                break;
        }
    }

    private static void LoadSolidFill(PatternFill patternFill, IXLFill XLiburFill, bool differentialFillFormat)
    {
        if (differentialFillFormat)
        {
            if (patternFill.BackgroundColor != null)
                XLiburFill.BackgroundColor = patternFill.BackgroundColor.ToXLiburColor();
            else
                XLiburFill.BackgroundColor = XLColor.FromIndex(64);
        }
        else
        {
            // yes, source is foreground!
            if (patternFill.ForegroundColor != null)
                XLiburFill.BackgroundColor = patternFill.ForegroundColor.ToXLiburColor();
            else
                XLiburFill.BackgroundColor = XLColor.FromIndex(64);
        }
    }

    private static void LoadPatternedFill(PatternFill patternFill, IXLFill XLiburFill)
    {
        if (patternFill.ForegroundColor != null)
            XLiburFill.PatternColor = patternFill.ForegroundColor.ToXLiburColor();

        if (patternFill.BackgroundColor != null)
            XLiburFill.BackgroundColor = patternFill.BackgroundColor.ToXLiburColor();
        else
            XLiburFill.BackgroundColor = XLColor.FromIndex(64);
    }

    internal static void LoadFont(OpenXmlElement? fontSource, IXLFontBase fontBase)
    {
        if (fontSource == null) return;

        fontBase.Bold = GetBoolean(fontSource.Elements<Bold>().FirstOrDefault());
        var fontColor = fontSource.Elements<Color>().FirstOrDefault();
        if (fontColor != null)
            fontBase.FontColor = fontColor.ToXLiburColor();

        LoadFontFamilyNumbering(fontSource, fontBase);
        LoadFontName(fontSource, fontBase);
        LoadFontSize(fontSource, fontBase);

        fontBase.Italic = GetBoolean(fontSource.Elements<Italic>().FirstOrDefault());
        fontBase.Shadow = GetBoolean(fontSource.Elements<Shadow>().FirstOrDefault());
        fontBase.Strikethrough = GetBoolean(fontSource.Elements<Strike>().FirstOrDefault());

        LoadFontUnderline(fontSource, fontBase);
        LoadFontVerticalAlignment(fontSource, fontBase);
        LoadFontScheme(fontSource, fontBase);
    }

    private static void LoadFontFamilyNumbering(OpenXmlElement fontSource, IXLFontBase fontBase)
    {
        var fontFamilyNumbering = fontSource.Elements<FontFamily>().FirstOrDefault();
        if (fontFamilyNumbering != null && fontFamilyNumbering.Val != null)
            fontBase.FontFamilyNumbering =
                (XLFontFamilyNumberingValues)int.Parse(fontFamilyNumbering.Val.ToString()!);
    }

    private static void LoadFontName(OpenXmlElement fontSource, IXLFontBase fontBase)
    {
        var runFont = fontSource.Elements<RunFont>().FirstOrDefault();
        if (runFont?.Val != null)
            // The guard proves the StringValue is non-null, but the implicit StringValue -> string
            // conversion returns string?, so the bang suppresses CS8601 on the assignment, not on
            // the guarded member access. S8969 only sees the latter.
#pragma warning disable S8969
            fontBase.FontName = runFont.Val!;
#pragma warning restore S8969
    }

    private static void LoadFontSize(OpenXmlElement fontSource, IXLFontBase fontBase)
    {
        var fontSize = fontSource.Elements<FontSize>().FirstOrDefault();
        if (fontSize?.Val != null)
            fontBase.FontSize = fontSize.Val;
    }

    private static void LoadFontUnderline(OpenXmlElement fontSource, IXLFontBase fontBase)
    {
        var underline = fontSource.Elements<Underline>().FirstOrDefault();
        if (underline != null)
            fontBase.Underline = underline.Val != null ? underline.Val.Value.ToXLibur() : XLFontUnderlineValues.Single;
    }

    private static void LoadFontVerticalAlignment(OpenXmlElement fontSource, IXLFontBase fontBase)
    {
        var verticalTextAlignment = fontSource.Elements<VerticalTextAlignment>().FirstOrDefault();
        if (verticalTextAlignment is not null)
            fontBase.VerticalAlignment = verticalTextAlignment.Val is not null ? verticalTextAlignment.Val.Value.ToXLibur() : XLFontVerticalTextAlignmentValues.Baseline;
    }

    private static void LoadFontScheme(OpenXmlElement fontSource, IXLFontBase fontBase)
    {
        var fontScheme = fontSource.Elements<FontScheme>().FirstOrDefault();
        if (fontScheme is not null)
            fontBase.FontScheme = fontScheme.Val is not null ? fontScheme.Val.Value.ToXLibur() : XLFontScheme.None;
    }

    internal static bool GetBoolean(BooleanPropertyType? property)
    {
        if (property != null)
        {
            if (property.Val != null)
                return property.Val;
            return true;
        }

        return false;
    }

    #endregion Public Methods

    #region Private Methods

    /// <summary>
    /// Here we perform the actual conversion from OpenXML color to XLibur color.
    /// </summary>
    /// <param name="openXMLColor">OpenXML color. Must be either <see cref="ColorType"/> or <see cref="X14.ColorType"/>.
    /// Since these types do not implement a common interface, we use dynamic.</param>
    /// <returns>The color in XLibur format.</returns>
    private static XLColor ConvertToXLiburColor(IColorTypeAdapter openXMLColor)
    {
        XLColor? retVal = null;

        // auto="1" is the explicit spelling of the automatic color (ECMA-376 CT_Color/@auto). It is
        // checked first because it wins over any other attribute present on the same element, and
        // the fall-through at the end resolves to the same color for an element that states nothing.
        if (openXMLColor.Auto?.Value == true)
            return XLColor.Automatic;

        if (openXMLColor.Rgb?.Value is not null)
        {
            var thisColor = ColorStringParser.ParseFromArgb(openXMLColor.Rgb.Value.AsSpan());
            retVal = XLColor.FromColor(thisColor);
        }
        else if (openXMLColor.Indexed is not null && openXMLColor.Indexed <= 64)
            retVal = XLColor.FromIndex((int)openXMLColor.Indexed.Value);
        else if (openXMLColor.Theme is not null)
        {
            retVal = openXMLColor.Tint is not null
                ? XLColor.FromTheme((XLThemeColor)openXMLColor.Theme.Value, openXMLColor.Tint.Value)
                : XLColor.FromTheme((XLThemeColor)openXMLColor.Theme.Value);
        }
        // An element stating none of rgb/indexed/theme is automatic, same as auto="1".
        return retVal ?? XLColor.Automatic;
    }

    /// <summary>
    /// Initialize properties of the existing instance of the color in OpenXML format basing on properties of the color
    /// in XLibur format.
    /// </summary>
    /// <param name="openXMLColor">OpenXML color. Must be either <see cref="ColorType"/> or <see cref="X14.ColorType"/>.
    /// Since these types do not implement a common interface we use dynamic.</param>
    /// <param name="xlColor">Color in XLibur format.</param>
    /// <param name="isDifferential">Flag specifying that the color should be saved in
    /// differential format (affects the transparent color processing).</param>
    private static void FillFromXLiburColor(IColorTypeAdapter openXMLColor, XLColor xlColor, bool isDifferential)
    {
        ArgumentNullException.ThrowIfNull(openXMLColor);
        ArgumentNullException.ThrowIfNull(xlColor);

        switch (xlColor.ColorType)
        {
            case XLColorType.Automatic:
                openXMLColor.Auto = true;
                break;

            case XLColorType.Color:
                openXMLColor.Rgb = xlColor.Color.ToHex();
                break;

            case XLColorType.Indexed:
                // 64 is 'transparent' and should be ignored for differential formats
                if (!isDifferential || xlColor.Indexed != 64)
                    openXMLColor.Indexed = (uint)xlColor.Indexed;
                break;

            case XLColorType.Theme:
                openXMLColor.Theme = (uint)xlColor.ThemeColor;

                if (xlColor.ThemeTint != 0)
                    openXMLColor.Tint = xlColor.ThemeTint;
                break;
        }
    }

    internal static int GetXLiburTextRotation(Alignment alignment)
    {
        if (alignment.TextRotation is null)
            return 0;

        var textRotation = (int)alignment.TextRotation.Value;
        return textRotation switch
        {
            255 => 255,
            > 90 => 90 - textRotation,
            _ => textRotation
        };
    }

    #endregion Private Methods
}
