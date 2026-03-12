namespace XLibur.Excel;

/// <summary>
/// Lightweight <see cref="IXLNumberFormat"/> that accumulates property changes into an <see cref="XLNumberFormatKey"/>
/// without triggering repository lookups or style-slice writes. Used by <see cref="XLDeferredStyle"/>
/// to support batch style mutations.
/// </summary>
internal sealed class XLDeferredNumberFormat : IXLNumberFormat
{
    private readonly XLDeferredStyle _style;
    internal XLNumberFormatKey Key;

    internal XLDeferredNumberFormat(XLDeferredStyle style, XLNumberFormatKey key)
    {
        _style = style;
        Key = key;
    }

    public int NumberFormatId
    {
        get => Key.NumberFormatId;
        set => Key = new XLNumberFormatKey
        {
            Format = XLNumberFormatValue.Default.Format,
            NumberFormatId = value,
        };
    }

    public string Format
    {
        get => Key.Format;
        set => Key = new XLNumberFormatKey
        {
            Format = value,
            NumberFormatId = string.IsNullOrWhiteSpace(value)
                ? XLNumberFormatValue.Default.NumberFormatId
                : XLNumberFormatKey.CustomFormatNumberId
        };
    }

    public IXLStyle SetNumberFormatId(int value) { NumberFormatId = value; return _style; }
    public IXLStyle SetFormat(string value) { Format = value; return _style; }

    public bool Equals(IXLNumberFormatBase? other) => other is XLDeferredNumberFormat dnf ? Key == dnf.Key : false;
    public override bool Equals(object? obj) => Equals(obj as IXLNumberFormatBase);
    public override int GetHashCode() => Key.GetHashCode();
}
