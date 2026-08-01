using System;

namespace XLibur.Excel;

/// <summary>
/// Lightweight <see cref="IXLFill"/> that accumulates property changes into an <see cref="XLFillKey"/>
/// without triggering repository lookups or style-slice writes. Used by <see cref="XLDeferredStyle"/>
/// to support batch style mutations.
/// </summary>
internal sealed class XLDeferredFill : IXLFill
{
    private readonly XLDeferredStyle _style;
    internal XLFillKey Key;

    internal XLDeferredFill(XLDeferredStyle style, XLFillKey key)
    {
        _style = style;
        Key = key;
    }

    public XLColor BackgroundColor
    {
        get
        {
            var colorKey = Key.BackgroundColor;
            return XLColor.FromKey(ref colorKey);
        }
        set
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value), "Color cannot be null");

            if ((PatternType == XLFillPatternValues.None ||
                 PatternType == XLFillPatternValues.Solid)
                && XLColor.IsNullOrTransparent(BackgroundColor))
            {
                var patternType = value.HasValue ? XLFillPatternValues.Solid : XLFillPatternValues.None;
                Key = Key with { BackgroundColor = value.Key, PatternType = patternType };
            }
            else
            {
                Key = Key with { BackgroundColor = value.Key };
            }
        }
    }

    public XLColor PatternColor
    {
        get
        {
            var colorKey = Key.PatternColor;
            return XLColor.FromKey(ref colorKey);
        }
        set
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value), "Color cannot be null");
            Key = Key with { PatternColor = value.Key };
        }
    }

    public XLFillPatternValues PatternType
    {
        get => Key.PatternType;
        set
        {
            if (Key.PatternType == XLFillPatternValues.None && value != XLFillPatternValues.None)
                Key = Key with { BackgroundColor = XLColor.FromTheme(XLThemeColor.Text1).Key, PatternType = value };
            else
                Key = Key with { PatternType = value };
        }
    }

    public IXLStyle SetBackgroundColor(XLColor value) { BackgroundColor = value; return _style; }
    public IXLStyle SetPatternColor(XLColor value) { PatternColor = value; return _style; }
    public IXLStyle SetPatternType(XLFillPatternValues value) { PatternType = value; return _style; }

    public bool Equals(IXLFill? other) => other is XLDeferredFill df ? Key == df.Key : false;
}
