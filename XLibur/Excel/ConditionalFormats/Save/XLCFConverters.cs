using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel;

internal static class XLCFConverters
{
    private static readonly Dictionary<XLConditionalFormatType, IXLCFConverter> Converters = new()
    {
        {XLConditionalFormatType.ColorScale, new XLCFColorScaleConverter()},
        {XLConditionalFormatType.StartsWith, XLCFTemplateConverter.StartsWith},
        {XLConditionalFormatType.EndsWith, XLCFTemplateConverter.EndsWith},
        {XLConditionalFormatType.IsBlank, XLCFTemplateConverter.IsBlank},
        {XLConditionalFormatType.NotBlank, XLCFTemplateConverter.NotBlank},
        {XLConditionalFormatType.IsError, XLCFTemplateConverter.IsError},
        {XLConditionalFormatType.NotError, XLCFTemplateConverter.NotError},
        {XLConditionalFormatType.ContainsText, XLCFTemplateConverter.Contains},
        {XLConditionalFormatType.NotContainsText, XLCFTemplateConverter.NotContains},
        {XLConditionalFormatType.CellIs, new XLCFCellIsConverter()},
        {XLConditionalFormatType.IsUnique, new XLCFUniqueConverter()},
        {XLConditionalFormatType.IsDuplicate, new XLCFUniqueConverter()},
        {XLConditionalFormatType.Expression, new XLCFCellIsConverter()},
        {XLConditionalFormatType.Top10, new XLCFTopConverter()},
        {XLConditionalFormatType.DataBar, new XLCFDataBarConverter()},
        {XLConditionalFormatType.IconSet, new XLCFIconSetConverter()},
        {XLConditionalFormatType.TimePeriod, new XLCFDatesOccurringConverter()}
    };

    public static ConditionalFormattingRule Convert(IXLConditionalFormat conditionalFormat, int priority, XLWorkbook.SaveContext context)
    {
        if (!Converters.TryGetValue(conditionalFormat.ConditionalFormatType, out var converter))
            throw new NotImplementedException($"Conditional formatting rule '{conditionalFormat.ConditionalFormatType}' hasn't been implemented");

        return converter.Convert(conditionalFormat, priority, context);
    }
}
