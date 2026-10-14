using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ConditionalFormats;
using XLibur.Utils;

namespace XLibur.Excel;

/// <summary>
/// Writes an above- or below-average rule (<c>cfRule type="aboveAverage"</c>) with the attributes the load
/// kept: <c>aboveAverage</c> and <c>equalAverage</c> only when they differ from their schema defaults, and
/// <c>stdDev</c> only when the rule has one (#743). The rule holds no formula.
/// </summary>
internal sealed class XLCFAboveAverageConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
    {
        var xlCf = (XLConditionalFormat)cf;
        var rule = XLCFBaseConverter.Convert(cf, priority, context);

        rule.AboveAverage = OpenXmlHelper.GetBooleanValue(xlCf.AboveAverage, true);
        rule.EqualAverage = OpenXmlHelper.GetBooleanValue(xlCf.EqualAverage, false);
        if (xlCf.StdDev is { } stdDev)
            rule.StdDev = stdDev;

        return rule;
    }
}
