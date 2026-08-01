using System;
using System.Drawing;
using System.Globalization;
using XLibur.Extensions;

namespace XLibur.Excel;

public enum XLColorType
{
    /// <summary>
    /// Automatic color, serialized as <c>&lt;color auto="1"/&gt;</c> or by omitting the color
    /// element entirely. The actual color is chosen by the application from the context it is used
    /// in - generally black for a font or border, white for a fill. <see cref="XLColor.Color"/> has
    /// no bearing on the resolved color and must not be read.
    /// <para>
    /// Deliberately the first member so that a default <c>XLColorKey</c> is automatic rather than a
    /// fully transparent RGB black, which is not a color any Excel file means to express.
    /// </para>
    /// </summary>
    Automatic,

    /// <summary>
    /// An RGB color, stored directly in <see cref="XLColor.Color"/>. It can carry an alpha
    /// component, but Excel ignores it and treats every color as fully opaque.
    /// </summary>
    Color,

    /// <summary>
    /// A theme color. The value depends on the workbook theme.
    /// </summary>
    Theme,

    /// <summary>
    /// An indexed color into the legacy palette, from the days when a fixed palette was the norm.
    /// The only semi-current uses are the system foreground (64) and background (65) colors.
    /// </summary>
    Indexed
}

public enum XLThemeColor
{
    Background1,
    Text1,
    Background2,
    Text2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink
}

public sealed partial class XLColor : IEquatable<XLColor>
{
    public bool HasValue { get; private set; }

    public XLColorType ColorType => Key.ColorType;

    /// <summary>
    /// Whether this is the <see cref="XLColorType.Automatic"/> color, i.e. no color was stated and
    /// the application resolves one from context.
    /// </summary>
    public bool IsAutomatic => ColorType == XLColorType.Automatic;

    public Color Color
    {
        get
        {
            if (ColorType == XLColorType.Color)
                return Key.Color;

            if (ColorType == XLColorType.Indexed)
                return IndexedColors[Indexed].Color;

            throw new InvalidOperationException($"Cannot convert {ColorType} color to Color.");
        }
    }

    public int Indexed
    {
        get
        {
            if (ColorType == XLColorType.Indexed)
                return Key.Indexed;

            throw new InvalidOperationException($"Cannot convert {ColorType} color to indexed color.");
        }
    }

    public XLThemeColor ThemeColor
    {
        get
        {
            if (ColorType == XLColorType.Theme)
                return Key.ThemeColor;

            throw new InvalidOperationException($"Cannot convert {ColorType} color to theme color.");
        }
    }

    public double ThemeTint
    {
        get
        {
            if (ColorType == XLColorType.Theme)
                return Key.ThemeTint;

            if (ColorType == XLColorType.Indexed)
                throw new InvalidOperationException("Cannot extract theme tint from an indexed color.");

            return Color.A / 255.0;
        }
    }

    #region IEquatable<XLColor> Members

    public bool Equals(XLColor? other)
    {
        return other is not null && Key == other.Key;
    }

    #endregion IEquatable<XLColor> Members

    public override bool Equals(object? obj)
    {
        return obj is XLColor color && Equals(color);
    }

    public override int GetHashCode()
    {
        var hashCode = 229333804;
        hashCode = hashCode * -1521134295 + HasValue.GetHashCode();
        hashCode = hashCode * -1521134295 + Key.GetHashCode();
        return hashCode;
    }

    public override string ToString()
    {
        if (ColorType == XLColorType.Automatic)
            return "Automatic";

        if (ColorType == XLColorType.Color)
            return Color.ToHex();

        if (ColorType == XLColorType.Theme)
            return $"Color Theme: {ThemeColor.ToString()}, Tint: {ThemeTint.ToString(CultureInfo.InvariantCulture)}";

        return "Color Index: " + Indexed;
    }

    public static bool operator ==(XLColor? left, XLColor? right)
    {
        // If both are null, or both are same instance, return true.
        if (ReferenceEquals(left, right)) return true;

        // If one is null, but not both, return false.
        if ((left as object) == null || (right as object) == null) return false;

        return left.Equals(right);
    }

    public static bool operator !=(XLColor? left, XLColor? right)
    {
        return !(left == right);
    }
}
