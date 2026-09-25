using System;

namespace XLibur.Excel;

internal sealed class XLNumberFormat : IXLNumberFormat
{
    #region Static members

    internal static XLNumberFormatKey GenerateKey(IXLNumberFormat? defaultNumberFormat) => defaultNumberFormat switch
    {
        null => XLNumberFormatValue.Default.Key,
        XLNumberFormat format => format.Key,
        _ => new XLNumberFormatKey
        {
            NumberFormatId = defaultNumberFormat.NumberFormatId,
            Format = defaultNumberFormat.Format
        },
    };

    #endregion Static members

    #region Properties

    private readonly XLStyle _style;

    private XLNumberFormatValue _value;

    /// <inheritdoc cref="XLFont.Key"/>
    internal XLNumberFormatKey Key
    {
        get
        {
            var pending = _style.Pending;
            return pending is null ? _value.Key : pending.NumberFormat;
        }
        private set => _value = XLNumberFormatValue.FromKey(ref value);
    }

    #endregion Properties

    #region Constructors

    /// <summary>
    /// Create an instance of XLNumberFormat initializing it with the specified value.
    /// </summary>
    /// <param name="style">Style to attach the new instance to.</param>
    /// <param name="value">Style value to use.</param>
    public XLNumberFormat(XLStyle? style, XLNumberFormatValue value)
    {
        _style = style ?? XLStyle.CreateEmptyStyle();
        _value = value;
    }

    public XLNumberFormat(XLStyle? style, XLNumberFormatKey key) : this(style, XLNumberFormatValue.FromKey(ref key))
    {
    }

    public XLNumberFormat(XLStyle? style = null, IXLNumberFormat? d = null) : this(style, GenerateKey(d))
    {
    }

    #endregion Constructors

    internal void SyncValue(XLNumberFormatValue value) { _value = value; }

    #region IXLNumberFormat Members

    public int NumberFormatId
    {
        get => Key.NumberFormatId;
        set => Apply(false, value, static (_, v) => new XLNumberFormatKey
        {
            Format = XLNumberFormatValue.Default.Format,
            NumberFormatId = v,
        });
    }

    public string Format
    {
        get => Key.Format;
        set => Apply(false, value, static (_, v) => new XLNumberFormatKey
        {
            Format = v,
            NumberFormatId = string.IsNullOrWhiteSpace(v)
                ? XLNumberFormatValue.Default.NumberFormatId
                : XLNumberFormatKey.CustomFormatNumberId
        });
    }

    public IXLStyle SetNumberFormatId(int value)
    {
        NumberFormatId = value;
        return _style;
    }

    public IXLStyle SetFormat(string value)
    {
        Format = value;
        return _style;
    }

    #endregion IXLNumberFormat Members

    /// <summary>
    /// Apply a new component key to the cell this facade is attached to.
    /// </summary>
    /// <remarks>
    /// The new key is deliberately <em>not</em> interned before being applied. Assigning it to
    /// <c>Key</c> first would run a repository lookup -- hashing the key and probing a dictionary --
    /// whose result is then thrown away: on a transition-cache hit <c>ModifyNumberFormat</c> never needs the
    /// component value at all, and on a miss it interns the component anyway inside
    /// <c>XLStyleValue.FromKey</c>. Taking the interned value back off the resulting style instead
    /// leaves the facade just as correct for later reads, at no lookup. Measured over 20,000 cells
    /// setting one property each, this was the single largest cost on the per-cell styling path.
    /// </remarks>
    private void SetKey(XLNumberFormatKey newKey)
    {
        _style.ModifyNumberFormat(newKey);
        _value = _style.Value.NumberFormat;
    }

    /// <remarks>
    /// <c>Key</c> is assigned only where the facade does not hold the very value its style does - see
    /// <see cref="XLFont"/>'s <c>Modify</c>.
    /// </remarks>
    private void Modify<T>(T value, Func<XLNumberFormatKey, T, XLNumberFormatKey> with)
    {
        if (!ReferenceEquals(_value, _style.Value.NumberFormat))
        {
            Key = with(Key, value);
            _style.Modify(styleKey => styleKey with { NumberFormat = with(styleKey.NumberFormat, value) });
            return;
        }

        _style.Modify(styleKey => styleKey with { NumberFormat = with(styleKey.NumberFormat, value) });
        _value = _style.Value.NumberFormat;
    }

    /// <inheritdoc cref="XLFont.Apply{T}"/>
    private void Apply<T>(bool unchanged, T value, Func<XLNumberFormatKey, T, XLNumberFormatKey> with)
    {
        if (unchanged && _style.SkipsUnchangedValues) return;
        if (_style.IsCellContainer)
            SetKey(with(Key, value));
        else
            Modify(value, with);
    }

    #region Overridden

    public override string ToString()
    {
        return NumberFormatId + "-" + Format;
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as IXLNumberFormatBase);
    }

    public bool Equals(IXLNumberFormatBase? other)
    {
        if (other is not XLNumberFormat otherN)
            return false;

        return Key == otherN.Key;
    }

    public override int GetHashCode()
    {
        var hashCode = 416600561;
        hashCode = hashCode * -1521134295 + Key.GetHashCode();
        return hashCode;
    }

    #endregion Overridden
}
