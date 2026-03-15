using System;

namespace XLibur.Excel;

/// <summary>
/// A lightweight <see cref="IXLStyle"/> implementation that accumulates property changes into an
/// <see cref="XLStyleKey"/> without any repository lookups, transition-cache queries, or style-slice
/// writes. After all modifications are applied, the final <see cref="Key"/> is resolved once and
/// written to the cell.
/// </summary>
internal sealed class XLDeferredStyle : IXLStyle
{
    private readonly XLDeferredFont _font;
    private readonly XLDeferredFill _fill;
    private readonly XLDeferredBorder _border;
    private readonly XLDeferredAlignment _alignment;
    private readonly XLDeferredNumberFormat _numberFormat;
    private readonly XLDeferredProtection _protection;
    private bool _includeQuotePrefix;

    internal XLDeferredStyle(XLStyleKey key)
    {
        _font = new XLDeferredFont(this, key.Font);
        _fill = new XLDeferredFill(this, key.Fill);
        _border = new XLDeferredBorder(this, key.Border);
        _alignment = new XLDeferredAlignment(this, key.Alignment);
        _numberFormat = new XLDeferredNumberFormat(this, key.NumberFormat);
        _protection = new XLDeferredProtection(this, key.Protection);
        _includeQuotePrefix = key.IncludeQuotePrefix;
    }

    /// <summary>
    /// Returns the accumulated style key after all modifications.
    /// </summary>
    internal XLStyleKey Key => new()
    {
        Font = _font.Key,
        Fill = _fill.Key,
        Border = _border.Key,
        Alignment = _alignment.Key,
        NumberFormat = _numberFormat.Key,
        Protection = _protection.Key,
        IncludeQuotePrefix = _includeQuotePrefix,
    };

    public IXLFont Font
    {
        get => _font;
        set => _font.Key = XLFont.GenerateKey(value);
    }

    public IXLAlignment Alignment
    {
        get => _alignment;
        set => _alignment.Key = XLAlignment.GenerateKey(value);
    }

    public IXLBorder Border
    {
        get => _border;
        set => _border.Key = XLBorder.GenerateKey(value);
    }

    public IXLFill Fill
    {
        get => _fill;
        set => _fill.Key = XLFill.GenerateKey(value);
    }

    public bool IncludeQuotePrefix
    {
        get => _includeQuotePrefix;
        set => _includeQuotePrefix = value;
    }

    public IXLNumberFormat NumberFormat
    {
        get => _numberFormat;
        set => _numberFormat.Key = XLNumberFormat.GenerateKey(value);
    }

    public IXLProtection Protection
    {
        get => _protection;
        set => _protection.Key = XLProtection.GenerateKey(value);
    }

    public IXLNumberFormat DateFormat => _numberFormat;

    public IXLStyle SetIncludeQuotePrefix(bool includeQuotePrefix = true)
    {
        _includeQuotePrefix = includeQuotePrefix;
        return this;
    }

    public IXLStyle Batch(Action<IXLStyle> modifications)
    {
        // Already deferred — just apply directly.
        modifications(this);
        return this;
    }

    public bool Equals(IXLStyle? other) => other is XLDeferredStyle ds && Key.Equals(ds.Key);
}
