using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ConditionalFormats;
using XLibur.Excel.IO;
using XLibur.Extensions;
using XLibur.Utils;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using ConditionalFormattingRule = DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule;
using DataBar = DocumentFormat.OpenXml.Spreadsheet.DataBar;

namespace XLibur.Excel;

internal sealed class XLCFDataBarConverter : IXLCFConverter
{
    public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
    {
        var conditionalFormattingRule = XLCFBaseConverter.Convert(cf, priority);

        var dataBar = new DataBar { ShowValue = SchemaDefault.Bool(null, !cf.ShowBarOnly, true) };

        var conditionalFormatValueObject1 = GetConditionalFormatValueObjectByIndex(cf, 1, ConditionalFormatValueObjectValues.Min);
        var conditionalFormatValueObject2 = GetConditionalFormatValueObjectByIndex(cf, 2, ConditionalFormatValueObjectValues.Max);

        dataBar.Append(conditionalFormatValueObject1);
        dataBar.Append(conditionalFormatValueObject2);
        dataBar.Append(new Color().FromXLiburColor<Color>(cf.Colors[1]));

        conditionalFormattingRule.Append(dataBar);

        var conditionalFormattingRuleExtensionList = new ConditionalFormattingRuleExtensionList();
        conditionalFormattingRuleExtensionList.Append(BuildRuleExtension(cf));
        conditionalFormattingRule.Append(conditionalFormattingRuleExtensionList);

        return conditionalFormattingRule;
    }

    private static ConditionalFormattingRuleExtension BuildRuleExtension(IXLConditionalFormat cf)
    {
        var conditionalFormattingRuleExtension = new ConditionalFormattingRuleExtension { Uri = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" };
        conditionalFormattingRuleExtension.AddNamespaceDeclaration("x14", OpenXmlConst.X14Main2009SsNs);
        var id = new Id
        {
            Text = ((XLConditionalFormat)cf).Id.WrapInBraces()
        };
        conditionalFormattingRuleExtension.Append(id);

        return conditionalFormattingRuleExtension;
    }

    private static ConditionalFormatValueObject GetConditionalFormatValueObjectByIndex(IXLConditionalFormat cf, int index, ConditionalFormatValueObjectValues defaultType)
    {
        var conditionalFormatValueObject = new ConditionalFormatValueObject();

        if (cf.ContentTypes.TryGetValue(index, out var contentType))
        {
            conditionalFormatValueObject.Type = contentType.ToOpenXml();
        }
        else
        {
            conditionalFormatValueObject.Type = defaultType;
        }

        if (cf.Values.TryGetValue(index, out var value1) && value1?.Value != null)
        {
            conditionalFormatValueObject.Val = value1.Value;
        }

        return conditionalFormatValueObject;
    }
}
