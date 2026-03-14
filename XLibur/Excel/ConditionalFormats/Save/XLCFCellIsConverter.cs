using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel;

internal class XLCFCellIsConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
    {
        string val = GetQuoted(cf.Values[1]);

        var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);
        var cfStyle = ((XLStyle)cf.Style).Value;
        if (!cfStyle.Equals(XLWorkbook.DefaultStyleValue))
            conditionalFormattingRule.FormatId = (uint)context.DifferentialFormats[cfStyle];

        conditionalFormattingRule.Operator = cf.Operator.ToOpenXml();

        var formula = new Formula(val);
        conditionalFormattingRule.Append(formula);

        if (cf.Operator == XLCFOperator.Between || cf.Operator == XLCFOperator.NotBetween)
        {
            var formula2 = new Formula { Text = GetQuoted(cf.Values[2]) };
            conditionalFormattingRule.Append(formula2);
        }

        return conditionalFormattingRule;
    }

    private string GetQuoted(XLFormula formula)
    {
        string value = formula.Value;

        if (formula.IsFormula ||
            value.StartsWith("\"") && value.EndsWith("\"") ||
            double.TryParse(value, XLHelper.NumberStyle, XLHelper.ParseCulture, out _))
        {
            return value;
        }

        return $"\"{value.Replace("\"", "\"\"")}\"";
    }
}
