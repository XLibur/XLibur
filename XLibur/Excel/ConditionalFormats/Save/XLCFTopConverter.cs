using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel;

internal sealed class XLCFTopConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
    {
        var val = uint.Parse(cf.Values[1].Value);
        var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority, context);

        conditionalFormattingRule.Percent = cf.Percent;
        conditionalFormattingRule.Rank = val;
        conditionalFormattingRule.Bottom = cf.Bottom;
        return conditionalFormattingRule;
    }
}
