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
