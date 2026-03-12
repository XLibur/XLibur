using System;
using System.Text;

namespace XLibur.Excel;

internal sealed class XLStyle : IXLStyle
{
    #region Static members

    public static XLStyle Default => new(XLStyleValue.Default);

    internal static XLStyleKey GenerateKey(IXLStyle? initialStyle) => initialStyle switch
    {
        null => Default.Key,
        XLStyle style => style.Key,
        _ => new XLStyleKey
        {
            Alignment = XLAlignment.GenerateKey(initialStyle.Alignment),
            Border = XLBorder.GenerateKey(initialStyle.Border),
            Fill = XLFill.GenerateKey(initialStyle.Fill),
            Font = XLFont.GenerateKey(initialStyle.Font),
            IncludeQuotePrefix = initialStyle.IncludeQuotePrefix,
            NumberFormat = XLNumberFormat.GenerateKey(initialStyle.NumberFormat),
            Protection = XLProtection.GenerateKey(initialStyle.Protection)
        },
    };

    internal static XLStyle CreateEmptyStyle()
    {
        return new XLStyle(new XLStylizedEmpty(null));
    }

    #endregion Static members

    #region properties

    private readonly IXLStylized? _container;

    internal XLStyleValue Value { get; private set; }

    internal XLStyleKey Key
    {
        get { return Value.Key; }
        private set
        {
            Value = XLStyleValue.FromKey(ref value);
        }
    }

    #endregion properties

    #region constructors

    public XLStyle(IXLStylized container, IXLStyle? initialStyle = null, bool useDefaultModify = true) : this(container, GenerateKey(initialStyle))
    {
    }

    public XLStyle(IXLStylized container, XLStyleKey key) : this(container, XLStyleValue.FromKey(ref key))
    {
    }

    internal XLStyle(IXLStylized container, XLStyleValue value)
    {
        _container = container ?? new XLStylizedEmpty(Default);
        Value = value;
    }

    /// <summary>
    /// To initialize XLStyle.Default only
    /// </summary>
    private XLStyle(XLStyleValue value)
    {
        Value = value;
    }

    #endregion constructors

    internal void Modify(Func<XLStyleKey, XLStyleKey> modification)
    {
        Key = modification(Key);

        if (_container != null)
        {
            _container.ModifyStyle(modification);
        }
    }

    /// <summary>
    /// Fast-path style modification for XLCell containers. Only called when <see cref="IsCellContainer"/> is true.
    /// Bypasses closure allocation by directly computing the new style key.
    /// </summary>
    internal void ModifyFont(XLFontKey newFontKey)
    {
        var styleKey = Key with { Font = newFontKey };
        Value = XLStyleValue.FromKey(ref styleKey);
        ((XLCell)_container!).SetStyleValue(Value);
    }

    /// <inheritdoc cref="ModifyFont"/>
    internal void ModifyBorder(XLBorderKey newBorderKey)
    {
        var styleKey = Key with { Border = newBorderKey };
        Value = XLStyleValue.FromKey(ref styleKey);
        ((XLCell)_container!).SetStyleValue(Value);
    }

    /// <inheritdoc cref="ModifyFont"/>
    internal void ModifyFill(XLFillKey newFillKey)
    {
        var styleKey = Key with { Fill = newFillKey };
        Value = XLStyleValue.FromKey(ref styleKey);
        ((XLCell)_container!).SetStyleValue(Value);
    }

    /// <inheritdoc cref="ModifyFont"/>
    internal void ModifyAlignment(XLAlignmentKey newAlignmentKey)
    {
        var styleKey = Key with { Alignment = newAlignmentKey };
        Value = XLStyleValue.FromKey(ref styleKey);
        ((XLCell)_container!).SetStyleValue(Value);
    }

    /// <inheritdoc cref="ModifyFont"/>
    internal void ModifyNumberFormat(XLNumberFormatKey newNumberFormatKey)
    {
        var styleKey = Key with { NumberFormat = newNumberFormatKey };
        Value = XLStyleValue.FromKey(ref styleKey);
        ((XLCell)_container!).SetStyleValue(Value);
    }

    /// <inheritdoc cref="ModifyFont"/>
    internal void ModifyProtection(XLProtectionKey newProtectionKey)
    {
        var styleKey = Key with { Protection = newProtectionKey };
        Value = XLStyleValue.FromKey(ref styleKey);
        ((XLCell)_container!).SetStyleValue(Value);
    }

    internal void SyncValue(XLStyleValue value) { Value = value; }

    /// <summary>
    /// True when the container is an XLCell, allowing fast-path style modifications without closures.
    /// </summary>
    internal bool IsCellContainer => _container is XLCell;

    #region Cached sub-wrappers

    private XLFont? _cachedFont;
    private XLAlignment? _cachedAlignment;
    private XLBorder? _cachedBorder;
    private XLFill? _cachedFill;
    private XLNumberFormat? _cachedNumberFormat;
    private XLProtection? _cachedProtection;

    #endregion Cached sub-wrappers

    #region IXLStyle members

    public IXLFont Font
    {
        get
        {
            if (_cachedFont == null)
                _cachedFont = new XLFont(this, Value.Font);
            else
                _cachedFont.SyncValue(Value.Font);
            return _cachedFont;
        }
        set
        {
            Modify(k => k with { Font = XLFont.GenerateKey(value) });
        }
    }

    public IXLAlignment Alignment
    {
        get
        {
            if (_cachedAlignment == null)
                _cachedAlignment = new XLAlignment(this, Value.Alignment);
            else
                _cachedAlignment.SyncValue(Value.Alignment);
            return _cachedAlignment;
        }
        set
        {
            Modify(k => k with { Alignment = XLAlignment.GenerateKey(value) });
        }
    }

    public IXLBorder Border
    {
        get
        {
            if (_cachedBorder == null)
                _cachedBorder = new XLBorder(_container!, this, Value.Border);
            else
                _cachedBorder.SyncValue(Value.Border);
            return _cachedBorder;
        }
        set
        {
            Modify(k => k with { Border = XLBorder.GenerateKey(value) });
        }
    }

    public IXLFill Fill
    {
        get
        {
            if (_cachedFill == null)
                _cachedFill = new XLFill(this, Value.Fill);
            else
                _cachedFill.SyncValue(Value.Fill);
            return _cachedFill;
        }
        set
        {
            Modify(k => k with { Fill = XLFill.GenerateKey(value) });
        }
    }

    public bool IncludeQuotePrefix
    {
        get { return Value.IncludeQuotePrefix; }
        set
        {
            Modify(k => k with { IncludeQuotePrefix = value });
        }
    }

    public IXLStyle SetIncludeQuotePrefix(bool includeQuotePrefix = true)
    {
        IncludeQuotePrefix = includeQuotePrefix;
        return this;
    }

    public IXLNumberFormat NumberFormat
    {
        get
        {
            if (_cachedNumberFormat == null)
                _cachedNumberFormat = new XLNumberFormat(this, Value.NumberFormat);
            else
                _cachedNumberFormat.SyncValue(Value.NumberFormat);
            return _cachedNumberFormat;
        }
        set
        {
            Modify(k => k with { NumberFormat = XLNumberFormat.GenerateKey(value) });
        }
    }

    public IXLProtection Protection
    {
        get
        {
            if (_cachedProtection == null)
                _cachedProtection = new XLProtection(this, Value.Protection);
            else
                _cachedProtection.SyncValue(Value.Protection);
            return _cachedProtection;
        }
        set
        {
            Modify(k => k with { Protection = XLProtection.GenerateKey(value) });
        }
    }

    public IXLNumberFormat DateFormat
    {
        get { return NumberFormat; }
    }

    #endregion IXLStyle members

    #region Overridden

    public override string ToString()
    {
        var sb = new StringBuilder();
        sb.Append("Font:");
        sb.Append(Font);
        sb.Append(" Fill:");
        sb.Append(Fill);
        sb.Append(" Border:");
        sb.Append(Border);
        sb.Append(" NumberFormat: ");
        sb.Append(NumberFormat);
        sb.Append(" Alignment: ");
        sb.Append(Alignment);
        sb.Append(" Protection: ");
        sb.Append(Protection);
        return sb.ToString();
    }

    public bool Equals(IXLStyle? other)
    {
        var otherS = other as XLStyle;

        if (otherS == null)
            return false;

        return Key == otherS.Key;
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as XLStyle);
    }

    public override int GetHashCode()
    {
        var hashCode = 416600561;
        hashCode = hashCode * -1521134295 + Key.GetHashCode();
        return hashCode;
    }

    #endregion Overridden
}
