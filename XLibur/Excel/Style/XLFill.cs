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

    /// <summary>
    /// Non-cell path (ranges, worksheets, conditional formats): apply the delta to each cell's own
    /// fill key.
    /// </summary>
    /// <remarks>
    /// A setter skips a value equal to <see cref="Key"/> only where
    /// <see cref="XLStyle.SkipsUnchangedValues"/> allows it: on a cell or a worksheet. Every other
    /// decision - whether a background colour brings the solid pattern with it, whether a new
    /// pattern needs a background colour - is taken from the key only on the cell path, where it is
    /// the cell's own style. On any other container the key is only that container's record of its
    /// style, which its cells need not share: a range is held weakly by its worksheet, and one
    /// rebuilt after a collection starts from its parent's style rather than its cells' (#505). So
    /// the non-cell path hands the decision to <paramref name="modification"/>, which runs once per
    /// distinct cell style.
    /// </remarks>
    private void Modify(Func<XLFillKey, XLFillKey> modification)
    {
        Key = modification(Key);
        _style.Modify(styleKey => styleKey with { Fill = modification(styleKey.Fill) });
    }

    /// <summary>
    /// <paramref name="key"/> with <paramref name="value"/> as its background colour, and the
    /// pattern that colour implies if the fill had no pattern of its own.
    /// </summary>
    private static XLFillKey WithBackgroundColor(in XLFillKey key, XLColor value)
        => ShouldAdjustPatternTypeForBackgroundColor(in key)
            ? key with { BackgroundColor = value.Key, PatternType = PatternTypeFromBackgroundColor(value) }
            : key with { BackgroundColor = value.Key };

    /// <summary>
    /// <paramref name="key"/> with <paramref name="value"/> as its pattern.
    /// </summary>
    /// <remarks>
    /// A fill that was empty and is given a pattern needs a background colour too. Otherwise the
    /// fill is still considered empty, and the pattern does not update (the cached empty fill is
    /// used).
    /// </remarks>
    private static XLFillKey WithPatternType(in XLFillKey key, XLFillPatternValues value)
        => key.PatternType == XLFillPatternValues.None && value != XLFillPatternValues.None
            ? key with { BackgroundColor = DefaultPatternBackgroundColorKey(), PatternType = value }
            : key with { PatternType = value };

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

            if (_style.IsCellContainer)
                SetKey(WithBackgroundColor(Key, value));
            else
                Modify(k => WithBackgroundColor(k, value));
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
            if (key.PatternColor == value.Key && _style.SkipsUnchangedValues) return;
            if (_style.IsCellContainer)
                SetKey(key with { PatternColor = value.Key });
            else
                Modify(k => k with { PatternColor = value.Key });
        }
    }

    public XLFillPatternValues PatternType
    {
        get => Key.PatternType;
        set
        {
            var key = Key;
            if (key.PatternType == value && _style.SkipsUnchangedValues) return;
            if (_style.IsCellContainer)
                SetKey(WithPatternType(key, value));
            else
                Modify(k => WithPatternType(k, value));
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
