using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel;

internal sealed class XLCFUniqueConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        => XLCFBaseConverter.Convert(cf, priority, context);
}
