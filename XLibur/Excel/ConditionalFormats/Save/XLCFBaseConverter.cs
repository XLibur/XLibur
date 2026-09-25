using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Utils;

namespace XLibur.Excel;

internal static class XLCFBaseConverter
{
    public static ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority)
    {
        return new ConditionalFormattingRule
        {
            Type = cf.ConditionalFormatType.ToOpenXml(),
            Priority = priority,
            StopIfTrue = OpenXmlHelper.GetBooleanValue(cf.StopIfTrue, false)
        };
    }

    /// <summary>
    /// The rule as <see cref="Convert(IXLConditionalFormat, int)"/> gives it, pointing at the
    /// differential format of the rule's style unless that style is the default one.
    /// </summary>
    public static ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority,
        XLWorkbook.SaveContext context)
    {
        var rule = Convert(cf, priority);
        var cfStyle = ((XLStyle)cf.Style).Value;
        if (!cfStyle.Equals(XLWorkbook.DefaultStyleValue))
            rule.FormatId = (uint)context.DifferentialFormats[cfStyle];

        return rule;
    }

    /// <summary>
    /// The cell a formula the converter builds is written relative to, as a relative A1 address.
    /// </summary>
    public static string AnchorAddress(IXLConditionalFormat cf)
        => cf.Range.RangeAddress.FirstAddress.ToStringRelative(false);
}
