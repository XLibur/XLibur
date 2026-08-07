using System;
using System.Drawing;

namespace XLibur.Excel;

/// <summary>
/// The identity of a colour: which kind of colour it is, plus the one payload that kind carries.
/// </summary>
/// <remarks>
/// The four colour kinds are mutually exclusive - <see cref="Equals(XLColorKey)"/> and
/// <see cref="GetHashCode"/> switch on <see cref="ColorType"/> and read exactly one payload - so the
/// payload is stored once, in a single 8-byte field reinterpreted per kind, rather than as one field
/// per kind. Only <see cref="XLColorType.Theme"/> needs a second component, and
/// <see cref="ThemeColor"/> is small enough to sit outside the union.
/// <para>
/// That matters because this struct is embedded five times in <c>XLBorderKey</c>, twice in
/// <c>XLFillKey</c> and once in <c>XLFontKey</c>, so its size is multiplied through the whole style
/// key. Holding a <see cref="System.Drawing.Color"/> - 24 bytes, of which the key only ever consumes
/// the 4 ARGB bytes - alongside separate <c>Indexed</c>, <c>ThemeColor</c> and <c>ThemeTint</c>
/// fields cost 48 bytes and made <c>XLStyleKey</c> 536 bytes. Storing the ARGB value instead, over a
/// union, brings this type to 16 bytes and <c>XLStyleKey</c> to 240.
/// </para>
/// <para>
/// Because a payload write depends on the kind, instances are built through the factory methods
/// rather than an object initialiser: two initialiser assignments would write the same field and the
/// second would silently win.
/// </para>
/// </remarks>
internal readonly struct XLColorKey : IEquatable<XLColorKey>
{
    /// <summary>
    /// The ARGB value, the palette index, or the raw bits of the theme tint, selected by
    /// <see cref="_colorType"/>. Unused for <see cref="XLColorType.Automatic"/>, which carries no
    /// value of its own.
    /// </summary>
    private readonly ulong _payload;

    /// <summary>
    /// <see cref="XLColorType"/> narrowed to a byte. The enum is public and stays <c>int</c>-backed;
    /// widening happens at the property boundary.
    /// </summary>
    private readonly byte _colorType;

    /// <summary><see cref="XLThemeColor"/> narrowed to a byte, on the same terms as <see cref="_colorType"/>.</summary>
    private readonly byte _themeColor;

    private XLColorKey(ulong payload, XLColorType colorType, XLThemeColor themeColor)
    {
        _payload = payload;
        _colorType = (byte)colorType;
        _themeColor = (byte)themeColor;
    }

    /// <summary>
    /// The automatic colour. Identical to <c>default</c>: <see cref="XLColorType.Automatic"/> is the
    /// first enum member, so a zeroed key already describes itself as automatic.
    /// </summary>
    public static XLColorKey Automatic => default;

    public static XLColorKey FromColor(Color color) => FromArgb(unchecked((uint)color.ToArgb()));

    public static XLColorKey FromArgb(uint argb) => new(argb, XLColorType.Color, default);

    public static XLColorKey FromIndex(int index) => new(unchecked((uint)index), XLColorType.Indexed, default);

    public static XLColorKey FromTheme(XLThemeColor themeColor) => FromTheme(themeColor, 0d);

    public static XLColorKey FromTheme(XLThemeColor themeColor, double themeTint) =>
        new(unchecked((ulong)BitConverter.DoubleToInt64Bits(themeTint)), XLColorType.Theme, themeColor);

    public XLColorType ColorType => (XLColorType)_colorType;

    /// <summary>
    /// The colour as an ARGB value. Only meaningful when <see cref="ColorType"/> is
    /// <see cref="XLColorType.Color"/>.
    /// </summary>
    public uint Argb => (uint)_payload;

    /// <summary>
    /// The colour. Only meaningful when <see cref="ColorType"/> is <see cref="XLColorType.Color"/>.
    /// </summary>
    /// <remarks>
    /// Rebuilt from the stored ARGB value, so it is never a <em>known</em> colour:
    /// <c>Color.Red</c> in gives a colour equal to <c>Color.FromArgb(255, 255, 0, 0)</c> back, whose
    /// <see cref="Color.Name"/> is its hex digits and whose <see cref="Color.IsNamedColor"/> is
    /// false. Everything the library does with a colour goes through the ARGB bytes, and a colour
    /// read out of a spreadsheet has no name to preserve in the first place.
    /// </remarks>
    public Color Color => Color.FromArgb(unchecked((int)(uint)_payload));

    /// <summary>
    /// The palette index. Only meaningful when <see cref="ColorType"/> is
    /// <see cref="XLColorType.Indexed"/>.
    /// </summary>
    public int Indexed => unchecked((int)(uint)_payload);

    /// <summary>
    /// The theme colour. Only meaningful when <see cref="ColorType"/> is
    /// <see cref="XLColorType.Theme"/>.
    /// </summary>
    public XLThemeColor ThemeColor => (XLThemeColor)_themeColor;

    /// <summary>
    /// The theme tint. Only meaningful when <see cref="ColorType"/> is
    /// <see cref="XLColorType.Theme"/>.
    /// </summary>
    public double ThemeTint => BitConverter.Int64BitsToDouble(unchecked((long)_payload));

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
                    hash = (hash * 397) ^ (int)Argb;
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
        if (_colorType != other._colorType) return false;
        switch (ColorType)
        {
            case XLColorType.Automatic:
                // Carries no value of its own - matching color types is the whole comparison.
                return true;

            case XLColorType.Color:
                return Argb == other.Argb;
            case XLColorType.Theme:
                if (_themeColor != other._themeColor)
                    return false;

                // Fast path for identical stored double values without floating-point ==. The tint
                // is held as its raw bits, so this is the payload comparison itself.
                if (_payload == other._payload)
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
