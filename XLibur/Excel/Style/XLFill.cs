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

    /// <summary>
    /// Apply a new component key to the cell this facade is attached to.
    /// </summary>
    /// <remarks>
    /// The new key is deliberately <em>not</em> interned before being applied. Assigning it to
    /// <c>Key</c> first would run a repository lookup -- hashing the key and probing a dictionary --
    /// whose result is then thrown away: on a transition-cache hit <c>ModifyFill</c> never needs the
    /// component value at all, and on a miss it interns the component anyway inside
    /// <c>XLStyleValue.FromKey</c>. Taking the interned value back off the resulting style instead
    /// leaves the facade just as correct for later reads, at no lookup. Measured over 20,000 cells
    /// setting one property each, this was the single largest cost on the per-cell styling path.
    /// </remarks>
    private void SetKey(XLFillKey newKey)
    {
        _style.ModifyFill(newKey);
        _value = _style.Value.Fill;
    }

    private void Modify(Func<XLFillKey, XLFillKey> modification)
    {
        Key = modification(Key);
        _style.Modify(styleKey => styleKey with { Fill = modification(styleKey.Fill) });
    }

    private void ApplyKeyUpdate(Func<XLFillKey, XLFillKey> update)
    {
        if (_style.IsCellContainer)
            SetKey(update(Key));
        else
            Modify(update);
    }

    private static XLFillPatternValues PatternTypeFromBackgroundColor(XLColor color)
        => color.HasValue ? XLFillPatternValues.Solid : XLFillPatternValues.None;

    private bool ShouldAdjustPatternTypeForBackgroundColor()
        => PatternType is XLFillPatternValues.None or XLFillPatternValues.Solid
           && XLColor.IsNullOrTransparent(BackgroundColor);

    private static XLColorKey DefaultPatternBackgroundColorKey()
        => XLColor.FromTheme(XLThemeColor.Text1).Key;

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

            if (ShouldAdjustPatternTypeForBackgroundColor())
            {
                var patternType = PatternTypeFromBackgroundColor(value);
                ApplyKeyUpdate(k => k with { BackgroundColor = value.Key, PatternType = patternType });
            }
            else
            {
                ApplyKeyUpdate(k => k with { BackgroundColor = value.Key });
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
            ApplyKeyUpdate(k => k with { PatternColor = value.Key });
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
                var defaultBackgroundColor = DefaultPatternBackgroundColorKey();
                ApplyKeyUpdate(k => k with { BackgroundColor = defaultBackgroundColor, PatternType = value });
            }
            else
            {
                if (Key.PatternType == value) return;
                ApplyKeyUpdate(k => k with { PatternType = value });
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
        if (other is not XLFill otherF)
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
