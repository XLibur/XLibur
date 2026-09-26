using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel;

/// <summary>
/// A rule whose one formula Excel writes from a fixed template: the blank, error and text rules.
/// The template's <c>{0}</c> is the rule's anchor cell (<see cref="XLCFBaseConverter.AnchorAddress"/>);
/// for a text rule, <c>{1}</c> is the text it looks for, escaped as a formula string literal, and
/// <c>{2}</c> the raw text's length.
/// </summary>
/// <remarks>As <see cref="XLCFDatesOccurringConverter"/> does for its time periods.</remarks>
internal sealed class XLCFTemplateConverter : IXLCFConverter
{
    public static readonly XLCFTemplateConverter IsBlank = new("LEN(TRIM({0}))=0");
    public static readonly XLCFTemplateConverter NotBlank = new("LEN(TRIM({0}))>0");
    public static readonly XLCFTemplateConverter IsError = new("ISERROR({0})");
    public static readonly XLCFTemplateConverter NotError = new("NOT(ISERROR({0}))");

    public static readonly XLCFTemplateConverter Contains =
        new("NOT(ISERROR(SEARCH(\"{1}\",{0})))", ConditionalFormattingOperatorValues.ContainsText);

    public static readonly XLCFTemplateConverter NotContains =
        new("ISERROR(SEARCH(\"{1}\",{0}))", ConditionalFormattingOperatorValues.NotContains);

    public static readonly XLCFTemplateConverter StartsWith =
        new("LEFT({0},{2})=\"{1}\"", ConditionalFormattingOperatorValues.BeginsWith);

    public static readonly XLCFTemplateConverter EndsWith =
        new("RIGHT({0},{2})=\"{1}\"", ConditionalFormattingOperatorValues.EndsWith);

    private readonly string _template;

    /// <summary>The operator of a text rule, or <c>null</c> for a rule that takes no text.</summary>
    private readonly ConditionalFormattingOperatorValues? _textOperator;

    private XLCFTemplateConverter(string template, ConditionalFormattingOperatorValues? textOperator = null)
    {
        _template = template;
        _textOperator = textOperator;
    }

    public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
    {
        var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority, context);

        string? text = null;
        if (_textOperator is { } textOperator)
        {
            text = cf.Values[1].Value;
            conditionalFormattingRule.Operator = textOperator;
            conditionalFormattingRule.Text = text;
        }

        // {1} sits inside a formula string literal, so a quote in the text is doubled there (as
        // Excel writes it); the text attribute keeps the raw text, and {2} is the raw text's length.
        var formulaText = string.Format(CultureInfo.InvariantCulture, _template,
            XLCFBaseConverter.AnchorAddress(cf), text?.Replace("\"", "\"\""), text?.Length);
        conditionalFormattingRule.Append(new Formula { Text = formulaText });

        return conditionalFormattingRule;
    }
}
