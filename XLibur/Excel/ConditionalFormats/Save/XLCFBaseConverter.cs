using XLibur.Utils;
using DocumentFormat.OpenXml.Spreadsheet;

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
}
