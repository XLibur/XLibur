using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.IO;

namespace XLibur.Excel;

internal sealed class XLCFIconSetConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
    {
        var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);

        // A schema default is left out, as Excel leaves it out (#709).
        var iconSet = new IconSet
        {
            ShowValue = SchemaDefault.Bool(null, !cf.ShowIconOnly, true),
            Reverse = SchemaDefault.Bool(null, cf.ReverseIconOrder, false),
            IconSetValue = cf.IconSetStyle.ToOpenXml()
        };
        var count = cf.Values.Count;
        for (var i = 1; i <= count; i++)
        {
            var conditionalFormatValueObject = new ConditionalFormatValueObject
            {
                Type = cf.ContentTypes[i].ToOpenXml(),
                Val = cf.Values[i].Value,
                GreaterThanOrEqual = SchemaDefault.Bool(null,
                    cf.IconSetOperators[i] == XLCFIconSetOperator.EqualOrGreaterThan, true)
            };
            iconSet.Append(conditionalFormatValueObject);

        }
        conditionalFormattingRule.Append(iconSet);
        return conditionalFormattingRule;
    }
}
