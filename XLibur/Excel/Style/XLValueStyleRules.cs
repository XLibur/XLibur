using System;

namespace XLibur.Excel;

/// <summary>
/// The adjustments a cell's style needs so a value survives a round-trip through the file.
/// </summary>
/// <remarks>
/// A worksheet cell stores a date, a duration and a number identically - as a serial number -
/// so the number format is the only thing that tells a reader which it was. The same goes for a
/// text value that starts with an apostrophe or spans several lines. Both writers depend on
/// these rules: <see cref="XLWorksheet.GetStyleForValue"/> applies them when a value is assigned
/// to a cell, and the streaming writer applies them as it serialises one.
/// </remarks>
internal static class XLValueStyleRules
{
    /// <summary>Built-in format <c>d/m/yyyy</c>, used for a date with no time part.</summary>
    internal const int DateOnlyNumberFormatId = 14;

    /// <summary>Built-in format <c>d/m/yyyy H:mm</c>, used for a date that carries a time.</summary>
    internal const int DateTimeNumberFormatId = 22;

    /// <summary>Built-in format <c>[h]:mm:ss</c>, used for an elapsed duration.</summary>
    internal const int DurationNumberFormatId = 46;

    /// <summary>
    /// Whether the style leaves the number format at General, in which case a date or duration
    /// written under it would read back as a plain number.
    /// </summary>
    internal static bool HasGeneralNumberFormat(XLStyleValue styleValue) =>
        styleValue.NumberFormat.Format.Length == 0 && styleValue.NumberFormat.NumberFormatId == 0;

    internal static XLStyleValue WithDateTimeFormat(XLStyleValue styleValue, bool onlyDatePart)
    {
        var numberFormat = styleValue.NumberFormat.WithNumberFormatId(
            onlyDatePart ? DateOnlyNumberFormatId : DateTimeNumberFormatId);
        return styleValue.WithNumberFormat(numberFormat);
    }

    internal static XLStyleValue WithDurationFormat(XLStyleValue styleValue)
    {
        var numberFormat = styleValue.NumberFormat.WithNumberFormatId(DurationNumberFormatId);
        return styleValue.WithNumberFormat(numberFormat);
    }

    /// <summary>
    /// The style a text value needs, or <c>null</c> when the passed one already suits it. A
    /// leading apostrophe is Excel's marker for "treat this as text", carried by the quote
    /// prefix flag rather than by the stored value; a multi-line value needs wrapping to be
    /// legible.
    /// </summary>
    internal static XLStyleValue? AdjustForText(XLStyleValue styleValue, string text)
    {
        XLStyleValue? adjusted = null;

        if (text.Length > 0 && text[0] == '\'')
            adjusted = styleValue.WithIncludeQuotePrefix(true);

        if (text.AsSpan().Contains(Environment.NewLine.AsSpan(), StringComparison.Ordinal))
        {
            adjusted ??= styleValue;
            if (!adjusted.Alignment.WrapText)
                adjusted = adjusted.WithAlignment(static alignment => alignment.WithWrapText(true));
        }

        return adjusted;
    }
}
