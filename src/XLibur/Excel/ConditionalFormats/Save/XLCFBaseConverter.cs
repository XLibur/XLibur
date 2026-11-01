using System;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ConditionalFormats;
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
    /// The cell a formula the converter builds is written relative to, as a relative A1 address: the
    /// rule's anchor (<see cref="XLConditionalFormat.AnchorOf"/>), which is also where Excel writes it
    /// from. For <c>B1:B5 A3:A5</c> that is <c>A1</c>, not the first area's <c>B1</c>.
    /// </summary>
    public static string AnchorAddress(IXLConditionalFormat cf)
    {
        var areas = ((XLConditionalFormat)cf).Areas;
        if (areas.Count == 0)
            throw new InvalidOperationException("XLConditionalFormat requires at least one Range.");

        return XLConditionalFormat.AnchorOf(areas).ToString();
    }
}
