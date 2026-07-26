using System;
using System.Drawing;

namespace XLibur.Excel;

internal readonly struct XLColorKey : IEquatable<XLColorKey>
{
    public XLColorType ColorType { get; init; }

    public Color Color { get; init; }

    public int Indexed { get; init; }

    public XLThemeColor ThemeColor { get; init; }

    public double ThemeTint { get; init; }

    /// <summary>
    /// True when no color was stated and the application resolves one from context. Unlike
    /// <see cref="XLColor.IsTransparent"/> this does not match indexed color 64, which is an
    /// explicitly written value and must round-trip as such.
    /// </summary>
    public bool IsAutomatic => ColorType == XLColorType.Automatic;

    public override int GetHashCode()
    {
        unchecked
        {
            var hash = (int)ColorType;

            switch (ColorType)
            {
                case XLColorType.Automatic:
                    // Carries no value of its own - the color type alone identifies it.
                    break;

                case XLColorType.Indexed:
                    hash = (hash * 397) ^ Indexed;
                    break;

                case XLColorType.Theme:
                    hash = (hash * 397) ^ (int)ThemeColor;
                    var tintHash = (int)(ThemeTint * 100000);
                    hash = (hash * 397) ^ tintHash;
                    break;

                case XLColorType.Color:
                    hash = (hash * 397) ^ Color.ToArgb();
                    break;
                default:
                    hash = (hash * 397) ^ Indexed;
                    break;
            }

            return hash;
        }
    }

    public bool Equals(XLColorKey other)
    {
        if (ColorType != other.ColorType) return false;
        switch (ColorType)
        {
            case XLColorType.Automatic:
                // Carries no value of its own - matching color types is the whole comparison.
                return true;

            case XLColorType.Color:
                return Color.ToArgb() == other.Color.ToArgb();
            case XLColorType.Theme:
                if (ThemeColor != other.ThemeColor)
                    return false;

                // Fast path for identical stored double values without floating-point ==.
                if (BitConverter.DoubleToInt64Bits(ThemeTint) == BitConverter.DoubleToInt64Bits(other.ThemeTint))
                    return true;

                return Math.Abs(ThemeTint - other.ThemeTint) < XLHelper.Epsilon;

            case XLColorType.Indexed:
            default:
                return Indexed == other.Indexed;
        }
    }

    public override bool Equals(object? obj)
    {
        if (obj is XLColorKey key)
            return Equals(key);
        return base.Equals(obj);
    }

    public override string ToString()
    {
        return ColorType switch
        {
            XLColorType.Automatic => "Automatic",
            XLColorType.Color => Color.ToString(),
            XLColorType.Theme => $"{ThemeColor} ({ThemeTint})",
            XLColorType.Indexed => $"Indexed: {Indexed}",
            _ => base.ToString()!
        };
    }

    public static bool operator ==(XLColorKey left, XLColorKey right) => left.Equals(right);

    public static bool operator !=(XLColorKey left, XLColorKey right) => !(left.Equals(right));
}
