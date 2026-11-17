using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Extensions;
using XLibur.Utils;

namespace XLibur.Excel;

internal sealed class XLCFColorScaleConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
    {
        var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);

        var colorScale = new ColorScale();
        for (int i = 1; i <= cf.ContentTypes.Count; i++)
        {
            var type = cf.ContentTypes[i].ToOpenXml();
            var val = cf.Values.TryGetValue(i, out var formula) ? formula?.Value : null;

            var conditionalFormatValueObject = new ConditionalFormatValueObject { Type = type };
            if (val != null)
                conditionalFormatValueObject.Val = val;

            colorScale.Append(conditionalFormatValueObject);
        }

        for (int i = 1; i <= cf.Colors.Count; i++)
            colorScale.Append(new Color().FromXLiburColor<Color>(cf.Colors[i]));

        conditionalFormattingRule.Append(colorScale);

        return conditionalFormattingRule;
    }
}
