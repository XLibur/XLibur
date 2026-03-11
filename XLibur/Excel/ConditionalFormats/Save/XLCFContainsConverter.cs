using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel;

internal sealed class XLCFContainsConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
    {
        string val = cf.Values[1].Value;
        var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
        var cfStyle = ((XLStyle)cf.Style).Value;
        if (!cfStyle.Equals(XLWorkbook.DefaultStyleValue))
            conditionalFormattingRule.FormatId = (uint)context.DifferentialFormats[cfStyle];

        conditionalFormattingRule.Operator = ConditionalFormattingOperatorValues.ContainsText;
        conditionalFormattingRule.Text = val;

        var formula = new Formula { Text = "NOT(ISERROR(SEARCH(\"" + val + "\"," + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + ")))" };

        conditionalFormattingRule.Append(formula);

        return conditionalFormattingRule;
    }
}
