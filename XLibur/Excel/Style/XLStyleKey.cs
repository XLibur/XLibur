namespace XLibur.Excel;

internal readonly record struct XLStyleKey
{
    private const string DefaultLabel = "Default";

    public required XLAlignmentKey Alignment { get; init; }

    public required XLBorderKey Border { get; init; }

    public required XLFillKey Fill { get; init; }

    public required XLFontKey Font { get; init; }

    public required bool IncludeQuotePrefix { get; init; }

    public required XLNumberFormatKey NumberFormat { get; init; }

    public required XLProtectionKey Protection { get; init; }

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
            var hash = Alignment.GetHashCode();
            hash = (hash * 397) ^ Border.GetHashCode();
            hash = (hash * 397) ^ Fill.GetHashCode();
            hash = (hash * 397) ^ Font.GetHashCode();
            hash = (hash * 397) ^ IncludeQuotePrefix.GetHashCode();
            hash = (hash * 397) ^ NumberFormat.GetHashCode();
            hash = (hash * 397) ^ Protection.GetHashCode();
            return hash;
        }
    }

    public bool Equals(XLStyleKey other)
    {
        // Order by discrimination power: font/fill/border vary most, protection/alignment least.
        return Font.Equals(other.Font)
               && Fill.Equals(other.Fill)
               && Border.Equals(other.Border)
               && NumberFormat.Equals(other.NumberFormat)
               && Alignment.Equals(other.Alignment)
               && IncludeQuotePrefix == other.IncludeQuotePrefix
               && Protection.Equals(other.Protection);
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
        alignment = Alignment;
        border = Border;
        fill = Fill;
        font = Font;
        includeQuotePrefix = IncludeQuotePrefix;
        numberFormat = NumberFormat;
        protection = Protection;
    }
}
