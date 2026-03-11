using System;
using System.Drawing;
using System.Globalization;

namespace XLibur.Excel;

public enum XLColorType
{
    Color,
    Theme,
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

public partial class XLColor : IEquatable<XLColor>
{
    public bool HasValue { get; private set; }

    public XLColorType ColorType
    {
        get { return Key.ColorType; }
    }

    public Color Color
    {
        get
        {
            if (ColorType == XLColorType.Theme)
                throw new InvalidOperationException("Cannot convert theme color to Color.");

            if (ColorType == XLColorType.Indexed)
                return IndexedColors[Indexed].Color;

            return Key.Color;
        }
    }

    public int Indexed
    {
        get
        {
            if (ColorType == XLColorType.Theme)
                throw new InvalidOperationException("Cannot convert theme color to indexed color.");

            if (ColorType == XLColorType.Indexed)
                return Key.Indexed;

            throw new InvalidOperationException("Cannot convert Color to indexed color.");
        }
    }

    public XLThemeColor ThemeColor
    {
        get
        {
            if (ColorType == XLColorType.Theme)
                return Key.ThemeColor;

            if (ColorType == XLColorType.Indexed)
                throw new InvalidOperationException("Cannot convert indexed color to theme color.");

            throw new InvalidOperationException("Cannot convert Color to theme color.");
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
