using XLibur.Excel.Caching;
using System.Collections.Generic;

namespace XLibur.Excel;

internal sealed class XLFontValue
{
    private static readonly XLFontRepository Repository = new(key => new XLFontValue(key));

    public static XLFontValue FromKey(ref XLFontKey key)
    {
        return Repository.GetOrCreate(ref key);
    }

    private static readonly XLFontKey DefaultKey = new()
    {
        Bold = false,
        Italic = false,
        Underline = XLFontUnderlineValues.None,
        Strikethrough = false,
        VerticalAlignment = XLFontVerticalTextAlignmentValues.Baseline,
        Shadow = false,
        FontSize = 11,
        FontColor = XLColor.FromArgb(0, 0, 0).Key,
        FontName = "Calibri",
        FontFamilyNumbering = XLFontFamilyNumberingValues.Swiss,
        FontCharSet = XLFontCharSet.Default,
        FontScheme = XLFontScheme.None
    };
    internal static readonly XLFontValue Default = FromKey(ref DefaultKey);

    public XLFontKey Key { get; private set; }

    public bool Bold => Key.Bold;

    public bool Italic => Key.Italic;

    public XLFontUnderlineValues Underline => Key.Underline;

    public bool Strikethrough => Key.Strikethrough;

    public XLFontVerticalTextAlignmentValues VerticalAlignment => Key.VerticalAlignment;

    public bool Shadow => Key.Shadow;

    public double FontSize => Key.FontSize;

    public XLColor FontColor { get; private set; }

    public string FontName => Key.FontName;

    public XLFontFamilyNumberingValues FontFamilyNumbering => Key.FontFamilyNumbering;

    public XLFontCharSet FontCharSet => Key.FontCharSet;

    public XLFontScheme FontScheme => Key.FontScheme;

    private XLFontValue(XLFontKey key)
    {
        Key = key;
        var fontColorKey = Key.FontColor;
        FontColor = XLColor.FromKey(ref fontColorKey);
    }

    public override bool Equals(object? obj)
    {
        var cached = obj as XLFontValue;
        return cached != null &&
               Key.Equals(cached.Key);
    }

    public override int GetHashCode()
    {
        return -280332839 + EqualityComparer<XLFontKey>.Default.GetHashCode(Key);
    }
}
