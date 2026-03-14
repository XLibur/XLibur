#nullable disable

using DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Excel;

internal interface IXLCFConverterExtension
{
    ConditionalFormattingRule Convert(IXLConditionalFormat cf, XLWorkbook.SaveContext context);
}
