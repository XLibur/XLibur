using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel;

internal class XLCFTopConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
    {
        uint val = uint.Parse(cf.Values[1].Value);
        var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
        var cfStyle = ((XLStyle)cf.Style).Value;
        if (!cfStyle.Equals(XLWorkbook.DefaultStyleValue))
            conditionalFormattingRule.FormatId = (uint)context.DifferentialFormats[cfStyle];

        conditionalFormattingRule.Percent = cf.Percent;
        conditionalFormattingRule.Rank = val;
        conditionalFormattingRule.Bottom = cf.Bottom;
        return conditionalFormattingRule;
    }
}
