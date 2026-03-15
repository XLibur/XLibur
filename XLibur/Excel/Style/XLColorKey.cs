using System;

namespace XLibur.Excel;

internal readonly struct XLColorKey : IEquatable<XLColorKey>
{
    public XLColorType ColorType { get; init; }

    public System.Drawing.Color Color { get; init; }

    public int Indexed { get; init; }

    public XLThemeColor ThemeColor { get; init; }

    public double ThemeTint { get; init; }

    public override int GetHashCode()
    {
        unchecked
        {
            var hash = (int)ColorType;

            switch (ColorType)
            {
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
                    throw new ArgumentOutOfRangeException();
            }

            return hash;
        }
    }

    public bool Equals(XLColorKey other)
    {
        if (ColorType != other.ColorType) return false;
        switch (ColorType)
        {
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
            XLColorType.Color => Color.ToString(),
            XLColorType.Theme => $"{ThemeColor} ({ThemeTint})",
            XLColorType.Indexed => $"Indexed: {Indexed}",
            _ => base.ToString()!
        };
    }

    public static bool operator ==(XLColorKey left, XLColorKey right) => left.Equals(right);

    public static bool operator !=(XLColorKey left, XLColorKey right) => !(left.Equals(right));
}
