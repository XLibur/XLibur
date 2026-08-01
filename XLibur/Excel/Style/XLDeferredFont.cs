using System;

namespace XLibur.Excel;

/// <summary>
/// Lightweight <see cref="IXLFont"/> that accumulates property changes into an <see cref="XLFontKey"/>
/// without triggering repository lookups or style-slice writes. Used by <see cref="XLDeferredStyle"/>
/// to support batch style mutations.
/// </summary>
internal sealed class XLDeferredFont : IXLFont
{
    private readonly XLDeferredStyle _style;
    internal XLFontKey Key;

    internal XLDeferredFont(XLDeferredStyle style, XLFontKey key)
    {
        _style = style;
        Key = key;
    }

    public bool Bold
    {
        get => Key.Bold;
        set => Key = Key with { Bold = value };
    }

    public bool Italic
    {
        get => Key.Italic;
        set => Key = Key with { Italic = value };
    }

    public XLFontUnderlineValues Underline
    {
        get => Key.Underline;
        set => Key = Key with { Underline = value };
    }

    public bool Strikethrough
    {
        get => Key.Strikethrough;
        set => Key = Key with { Strikethrough = value };
    }

    public XLFontVerticalTextAlignmentValues VerticalAlignment
    {
        get => Key.VerticalAlignment;
        set => Key = Key with { VerticalAlignment = value };
    }

    public bool Shadow
    {
        get => Key.Shadow;
        set => Key = Key with { Shadow = value };
    }

    public double FontSize
    {
        get => Key.FontSize;
        set => Key = Key with { FontSize = value };
    }

    public XLColor FontColor
    {
        get
        {
            var colorKey = Key.FontColor;
            return XLColor.FromKey(ref colorKey);
        }
        set
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value), "Color cannot be null");
            Key = Key with { FontColor = value.Key };
        }
    }

    public string FontName
    {
        get => Key.FontName;
        set => Key = Key with { FontName = value };
    }

    public XLFontFamilyNumberingValues FontFamilyNumbering
    {
        get => Key.FontFamilyNumbering;
        set => Key = Key with { FontFamilyNumbering = value };
    }

    public XLFontCharSet FontCharSet
    {
        get => Key.FontCharSet;
        set => Key = Key with { FontCharSet = value };
    }

    public XLFontScheme FontScheme
    {
        get => Key.FontScheme;
        set => Key = Key with { FontScheme = value };
    }

    public IXLStyle SetBold() { Bold = true; return _style; }
    public IXLStyle SetBold(bool value) { Bold = value; return _style; }
    public IXLStyle SetItalic() { Italic = true; return _style; }
    public IXLStyle SetItalic(bool value) { Italic = value; return _style; }
    public IXLStyle SetUnderline() { Underline = XLFontUnderlineValues.Single; return _style; }
    public IXLStyle SetUnderline(XLFontUnderlineValues value) { Underline = value; return _style; }
    public IXLStyle SetStrikethrough() { Strikethrough = true; return _style; }
    public IXLStyle SetStrikethrough(bool value) { Strikethrough = value; return _style; }
    public IXLStyle SetVerticalAlignment(XLFontVerticalTextAlignmentValues value) { VerticalAlignment = value; return _style; }
    public IXLStyle SetShadow() { Shadow = true; return _style; }
    public IXLStyle SetShadow(bool value) { Shadow = value; return _style; }
    public IXLStyle SetFontSize(double value) { FontSize = value; return _style; }
    public IXLStyle SetFontColor(XLColor value) { FontColor = value; return _style; }
    public IXLStyle SetFontName(string value) { FontName = value; return _style; }
    public IXLStyle SetFontFamilyNumbering(XLFontFamilyNumberingValues value) { FontFamilyNumbering = value; return _style; }
    public IXLStyle SetFontCharSet(XLFontCharSet value) { FontCharSet = value; return _style; }
    public IXLStyle SetFontScheme(XLFontScheme value) { FontScheme = value; return _style; }

    public bool Equals(IXLFont? other) => other is XLDeferredFont df ? Key == df.Key : false;
}
