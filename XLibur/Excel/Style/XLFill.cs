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

    /// <inheritdoc cref="XLFont.Key"/>
    internal XLFillKey Key
    {
        get
        {
            var pending = _style.Pending;
            return pending is null ? _value.Key : pending.Fill;
        }
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

    /// <summary>
    /// Applies a key delta on behalf of a caller that has already read <see cref="Key"/>, so the
    /// cell path does not read it a second time. The delta is still needed for the non-cell path,
    /// which must apply it to each cell's own key rather than to this facade's.
    /// </summary>
    private void ApplyKeyUpdate(in XLFillKey key, Func<XLFillKey, XLFillKey> update)
    {
        if (_style.IsCellContainer)
            SetKey(update(key));
        else
            Modify(update);
    }

    private static XLFillPatternValues PatternTypeFromBackgroundColor(XLColor color)
        => color.HasValue ? XLFillPatternValues.Solid : XLFillPatternValues.None;

    /// <remarks>
    /// Takes the key the caller has already read rather than going back through
    /// <see cref="PatternType"/> and <see cref="BackgroundColor"/>, each of which reads it again.
    /// </remarks>
    private static bool ShouldAdjustPatternTypeForBackgroundColor(in XLFillKey key)
    {
        if (key.PatternType is not (XLFillPatternValues.None or XLFillPatternValues.Solid))
            return false;

        var backgroundColorKey = key.BackgroundColor;
        return XLColor.IsNullOrTransparent(XLColor.FromKey(ref backgroundColorKey));
    }

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

            var key = Key;
            if (ShouldAdjustPatternTypeForBackgroundColor(in key))
            {
                var patternType = PatternTypeFromBackgroundColor(value);
                ApplyKeyUpdate(in key, k => k with { BackgroundColor = value.Key, PatternType = patternType });
            }
            else
            {
                ApplyKeyUpdate(in key, k => k with { BackgroundColor = value.Key });
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

            var key = Key;
            if (key.PatternColor == value.Key) return;
            ApplyKeyUpdate(in key, k => k with { PatternColor = value.Key });
        }
    }

    public XLFillPatternValues PatternType
    {
        get => Key.PatternType;
        set
        {
            var key = Key;
            if (key.PatternType == XLFillPatternValues.None &&
                value != XLFillPatternValues.None)
            {
                // If fill was empty and the pattern changes to non-empty, we have to specify a background color too.
                // Otherwise, the fill will be considered empty, and the pattern won't update (the cached empty fill will be used).
                var defaultBackgroundColor = DefaultPatternBackgroundColorKey();
                ApplyKeyUpdate(in key, k => k with { BackgroundColor = defaultBackgroundColor, PatternType = value });
            }
            else
            {
                if (key.PatternType == value) return;
                ApplyKeyUpdate(in key, k => k with { PatternType = value });
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
