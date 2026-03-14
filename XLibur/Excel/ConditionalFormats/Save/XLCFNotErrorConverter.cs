using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel;

internal class XLCFNotErrorConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
    {
        var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
        var cfStyle = ((XLStyle)cf.Style).Value;
        if (!cfStyle.Equals(XLWorkbook.DefaultStyleValue))
            conditionalFormattingRule.FormatId = (uint)context.DifferentialFormats[cfStyle];

        var formula = new Formula { Text = "NOT(ISERROR(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + "))" };

        conditionalFormattingRule.Append(formula);

        return conditionalFormattingRule;
    }
}
