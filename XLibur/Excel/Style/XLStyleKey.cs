namespace XLibur.Excel;

/// <summary>
/// The composite key of a style. Instances are only ever produced through object initialisers
/// (every member is <c>required</c>) or <c>with</c> expressions, so each component's hash code is
/// computed once in the corresponding <c>init</c> accessor and reused afterwards.
/// </summary>
/// <remarks>
/// The style repository hashes and compares this key on every style mutation, and the components
/// are large structs whose hashes are expensive (the border alone carries five colours, and the
/// number format hashes a string). Re-deriving all of that on every dictionary probe dominated
/// the cost of styling a sheet, so the per-component hashes are memoised in the key itself and
/// <see cref="GetHashCode"/> just folds seven ints together. The memoised hashes also give
/// <see cref="Equals(XLStyleKey)"/> a cheap way to reject non-matching keys before touching the
/// components.
/// </remarks>
internal readonly record struct XLStyleKey
{
    private const string DefaultLabel = "Default";

    private readonly XLAlignmentKey _alignment;
    private readonly XLBorderKey _border;
    private readonly XLFillKey _fill;
    private readonly XLFontKey _font;
    private readonly XLNumberFormatKey _numberFormat;
    private readonly XLProtectionKey _protection;

    private readonly int _alignmentHash;
    private readonly int _borderHash;
    private readonly int _fillHash;
    private readonly int _fontHash;
    private readonly int _numberFormatHash;
    private readonly int _protectionHash;

    public required XLAlignmentKey Alignment
    {
        get => _alignment;
        init
        {
            _alignment = value;
            _alignmentHash = value.GetHashCode();
        }
    }

    public required XLBorderKey Border
    {
        get => _border;
        init
        {
            _border = value;
            _borderHash = value.GetHashCode();
        }
    }

    public required XLFillKey Fill
    {
        get => _fill;
        init
        {
            _fill = value;
            _fillHash = value.GetHashCode();
        }
    }

    public required XLFontKey Font
    {
        get => _font;
        init
        {
            _font = value;
            _fontHash = value.GetHashCode();
        }
    }

    public required bool IncludeQuotePrefix { get; init; }

    public required XLNumberFormatKey NumberFormat
    {
        get => _numberFormat;
        init
        {
            _numberFormat = value;
            _numberFormatHash = value.GetHashCode();
        }
    }

    public required XLProtectionKey Protection
    {
        get => _protection;
        init
        {
            _protection = value;
            _protectionHash = value.GetHashCode();
        }
    }

    public override string ToString()
    {
        if (this == XLStyle.Default.Key)
            return DefaultLabel;

        var defaultKey = XLStyle.Default.Key;
        var alignment = Alignment == defaultKey.Alignment ? DefaultLabel : Alignment.ToString();
        var border = Border == defaultKey.Border ? DefaultLabel : Border.ToString();
        var fill = Fill == defaultKey.Fill ? DefaultLabel : Fill.ToString();
        var font = Font == defaultKey.Font ? DefaultLabel : Font.ToString();
        var includeQuotePrefix = IncludeQuotePrefix == defaultKey.IncludeQuotePrefix ? DefaultLabel : IncludeQuotePrefix.ToString();
        var numberFormat = NumberFormat == defaultKey.NumberFormat ? DefaultLabel : NumberFormat.ToString();
        var protection = Protection == defaultKey.Protection ? DefaultLabel : Protection.ToString();

        return string.Format("Alignment: {0} Border: {1} Fill: {2} Font: {3} IncludeQuotePrefix: {4} NumberFormat: {5} Protection: {6}",
            alignment, border, fill, font, includeQuotePrefix, numberFormat, protection);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hash = _alignmentHash;
            hash = (hash * 397) ^ _borderHash;
            hash = (hash * 397) ^ _fillHash;
            hash = (hash * 397) ^ _fontHash;
            hash = (hash * 397) ^ (IncludeQuotePrefix ? 1 : 0);
            hash = (hash * 397) ^ _numberFormatHash;
            hash = (hash * 397) ^ _protectionHash;
            return hash;
        }
    }

    public bool Equals(XLStyleKey other)
    {
        // Reject on the memoised component hashes first: a mismatch there is proof of inequality
        // (each component's hash is consistent with its Equals) and costs a single int compare.
        if (_fontHash != other._fontHash
            || _fillHash != other._fillHash
            || _borderHash != other._borderHash
            || _numberFormatHash != other._numberFormatHash
            || _alignmentHash != other._alignmentHash
            || _protectionHash != other._protectionHash
            || IncludeQuotePrefix != other.IncludeQuotePrefix)
        {
            return false;
        }

        // Order by discrimination power: font/fill/border vary most, protection/alignment least.
        return _font.Equals(other._font)
               && _fill.Equals(other._fill)
               && _border.Equals(other._border)
               && _numberFormat.Equals(other._numberFormat)
               && _alignment.Equals(other._alignment)
               && _protection.Equals(other._protection);
    }

    public void Deconstruct(
        out XLAlignmentKey alignment,
        out XLBorderKey border,
        out XLFillKey fill,
        out XLFontKey font,
        out bool includeQuotePrefix,
        out XLNumberFormatKey numberFormat,
        out XLProtectionKey protection)
    {
        alignment = _alignment;
        border = _border;
        fill = _fill;
        font = _font;
        includeQuotePrefix = IncludeQuotePrefix;
        numberFormat = _numberFormat;
        protection = _protection;
    }
}
