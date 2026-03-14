using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel;

internal interface IXLCFConverter
{
    ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context);
}
