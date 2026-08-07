using System;
using System.Collections.Generic;
using XLibur.Excel.Caching;

namespace XLibur.Excel;

internal sealed class XLFontValue : IEquatable<XLFontValue?>
{
    private static readonly XLRepositoryBase<XLFontKey, XLFontValue> Repository = new(key => new XLFontValue(key));

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

    public XLFontKey Key { get; }

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

    /// <inheritdoc cref="XLBorderValue._hashCode"/>
    /// <remarks>The font key's hash includes the hash of the font name.</remarks>
    private readonly int _hashCode;

    private XLFontValue(XLFontKey key)
    {
        Key = key;
        _hashCode = -280332839 + EqualityComparer<XLFontKey>.Default.GetHashCode(key);
        var fontColorKey = Key.FontColor;
        FontColor = XLColor.FromKey(ref fontColorKey);
    }

    public override bool Equals(object? obj)
    {
        return ReferenceEquals(this, obj) || Equals(obj as XLFontValue);
    }

    /// <inheritdoc cref="XLBorderValue.Equals(XLBorderValue)"/>
    public bool Equals(XLFontValue? other)
    {
        if (other is null)
            return false;

        return ReferenceEquals(this, other) || (_hashCode == other._hashCode && Key.Equals(other.Key));
    }

    public override int GetHashCode() => _hashCode;
}
