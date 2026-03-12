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

    internal XLNumberFormatKey Key
    {
        get => _value.Key;
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
        set
        {
            if (_style.IsCellContainer)
            {
                SetKey(new XLNumberFormatKey
                {
                    Format = XLNumberFormatValue.Default.Format,
                    NumberFormatId = value,
                });
            }
            else
            {
                Modify(_ => new XLNumberFormatKey
                {
                    Format = XLNumberFormatValue.Default.Format,
                    NumberFormatId = value,
                });
            }
        }
    }

    public string Format
    {
        get => Key.Format;
        set
        {
            if (_style.IsCellContainer)
            {
                SetKey(new XLNumberFormatKey
                {
                    Format = value,
                    NumberFormatId = string.IsNullOrWhiteSpace(value)
                        ? XLNumberFormatValue.Default.NumberFormatId
                        : XLNumberFormatKey.CustomFormatNumberId
                });
            }
            else
            {
                Modify(_ => new XLNumberFormatKey
                {
                    Format = value,
                    NumberFormatId = string.IsNullOrWhiteSpace(value)
                        ? XLNumberFormatValue.Default.NumberFormatId
                        : XLNumberFormatKey.CustomFormatNumberId
                });
            }
        }
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

    private void SetKey(XLNumberFormatKey newKey)
    {
        Key = newKey;
        _style.ModifyNumberFormat(Key);
    }

    private void Modify(Func<XLNumberFormatKey, XLNumberFormatKey> modification)
    {
        Key = modification(Key);
        _style.Modify(styleKey => styleKey with { NumberFormat = modification(styleKey.NumberFormat) });
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
        var otherN = other as XLNumberFormat;
        if (otherN == null)
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
