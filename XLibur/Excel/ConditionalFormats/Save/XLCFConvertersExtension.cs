using DocumentFormat.OpenXml.Office2010.Excel;
using System.Collections.Generic;

namespace XLibur.Excel;

internal sealed class XLCFConvertersExtension
{
    private static readonly Dictionary<XLConditionalFormatType, IXLCFConverterExtension> Converters;

    static XLCFConvertersExtension()
    {
        Converters = new Dictionary<XLConditionalFormatType, IXLCFConverterExtension>()
        {
            { XLConditionalFormatType.DataBar, new XLCFDataBarConverterExtension() }
        };
    }

    public static ConditionalFormattingRule Convert(IXLConditionalFormat conditionalFormat, XLWorkbook.SaveContext context)
    {
        return Converters[conditionalFormat.ConditionalFormatType].Convert(conditionalFormat, context);
    }
}
