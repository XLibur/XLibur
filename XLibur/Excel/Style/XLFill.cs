using System;

namespace XLibur.Excel;

internal sealed class XLFill : IXLFill
{
    #region static members

    internal static XLFillKey GenerateKey(IXLFill? defaultFill) => defaultFill switch
    {
        null => XLFillValue.Default.Key,
        XLFill fill => fill.Key,
        _ => new XLFillKey
        {
            PatternType = defaultFill.PatternType,
            BackgroundColor = defaultFill.BackgroundColor.Key,
            PatternColor = defaultFill.PatternColor.Key
        },
    };

    #endregion static members

    #region Properties

    private readonly XLStyle _style;

    private XLFillValue _value;

    internal XLFillKey Key
    {
        get => _value.Key;
        private set => _value = XLFillValue.FromKey(ref value);
    }

    #endregion Properties

    #region Constructors

    /// <summary>
    /// Create an instance of XLFill initializing it with the specified value.
    /// </summary>
    /// <param name="style">Style to attach the new instance to.</param>
    /// <param name="value">Style value to use.</param>
    public XLFill(XLStyle? style, XLFillValue value)
    {
        _style = style ?? XLStyle.CreateEmptyStyle();
        _value = value;
    }

    private XLFill(XLStyle? style, XLFillKey key) : this(style, XLFillValue.FromKey(ref key))
    {
    }

    public XLFill(XLStyle? style = null, IXLFill? d = null) : this(style, GenerateKey(d))
    {
    }

    #endregion Constructors

    internal void SyncValue(XLFillValue value) { _value = value; }

    private void SetKey(XLFillKey newKey)
    {
        Key = newKey;
        _style.ModifyFill(Key);
    }

    private void Modify(Func<XLFillKey, XLFillKey> modification)
    {
        Key = modification(Key);
        _style.Modify(styleKey => styleKey with { Fill = modification(styleKey.Fill) });
    }

    #region IXLFill Members

    public XLColor BackgroundColor
    {
        get
        {
            var backgroundColorKey = Key.BackgroundColor;
            return XLColor.FromKey(ref backgroundColorKey);
        }
        set
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value), "Color cannot be null");

            if (PatternType is XLFillPatternValues.None or XLFillPatternValues.Solid
                && XLColor.IsNullOrTransparent(BackgroundColor))
            {
                var patternType = value.HasValue ? XLFillPatternValues.Solid : XLFillPatternValues.None;
                if (_style.IsCellContainer)
                    SetKey(Key with { BackgroundColor = value.Key, PatternType = patternType });
                else
                    Modify(k => k with { BackgroundColor = value.Key, PatternType = patternType });
            }
            else
            {
                if (_style.IsCellContainer)
                    SetKey(Key with { BackgroundColor = value.Key });
                else
                    Modify(k => k with { BackgroundColor = value.Key });
            }
        }
    }

    public XLColor PatternColor
    {
        get
        {
            var patternColorKey = Key.PatternColor;
            return XLColor.FromKey(ref patternColorKey);
        }
        set
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value), "Color cannot be null");

            if (Key.PatternColor == value.Key) return;
            if (_style.IsCellContainer)
                SetKey(Key with { PatternColor = value.Key });
            else
                Modify(k => k with { PatternColor = value.Key });
        }
    }

    public XLFillPatternValues PatternType
    {
        get => Key.PatternType;
        set
        {
            if (PatternType == XLFillPatternValues.None &&
                value != XLFillPatternValues.None)
            {
                // If fill was empty and the pattern changes to non-empty, we have to specify a background color too.
                // Otherwise, the fill will be considered empty, and the pattern won't update (the cached empty fill will be used).
                if (_style.IsCellContainer)
                    SetKey(Key with { BackgroundColor = XLColor.FromTheme(XLThemeColor.Text1).Key, PatternType = value });
                else
                    Modify(k => k with { BackgroundColor = XLColor.FromTheme(XLThemeColor.Text1).Key, PatternType = value });
            }
            else
            {
                if (Key.PatternType == value) return;
                if (_style.IsCellContainer)
                    SetKey(Key with { PatternType = value });
                else
                    Modify(k => k with { PatternType = value });
            }
        }
    }

    public IXLStyle SetBackgroundColor(XLColor value)
    {
        BackgroundColor = value;
        return _style;
    }

    public IXLStyle SetPatternColor(XLColor value)
    {
        PatternColor = value;
        return _style;
    }

    public IXLStyle SetPatternType(XLFillPatternValues value)
    {
        PatternType = value;
        return _style;
    }

    #endregion IXLFill Members

    #region Overridden

    public override bool Equals(object? obj)
    {
        return Equals(obj as XLFill);
    }

    public bool Equals(IXLFill? other)
    {
        var otherF = other as XLFill;
        if (otherF == null)
            return false;

        return Key == otherF.Key;
    }

    public override string ToString() => PatternType switch
    {
        XLFillPatternValues.None => "None",
        XLFillPatternValues.Solid => string.Concat("Solid ", BackgroundColor.ToString()),
        _ => string.Concat(PatternType.ToString(), " pattern: ", PatternColor.ToString(), " on ", BackgroundColor.ToString()),
    };

    public override int GetHashCode()
    {
        var hashCode = -1938644919;
        hashCode = hashCode * -1521134295 + Key.GetHashCode();
        return hashCode;
    }

    #endregion Overridden
}
