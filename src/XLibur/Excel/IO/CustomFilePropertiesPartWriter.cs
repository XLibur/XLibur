using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;

namespace XLibur.Excel.IO;

internal static class CustomFilePropertiesPartWriter
{
    internal static void GenerateContent(CustomFilePropertiesPart customFilePropertiesPart, XLWorkbook workbook)
    {
        var properties = new Properties();
        properties.AddNamespaceDeclaration("vt",
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
        var propertyId = 1;
        foreach (var p in workbook.CustomProperties)
        {
            propertyId++;
            var customDocumentProperty = new CustomDocumentProperty
            {
                FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                PropertyId = propertyId,
                Name = p.Name
            };
            if (p is XLCustomProperty { LinkTarget: { } linkTarget })
                customDocumentProperty.LinkTarget = linkTarget;
            if (p.Type == XLCustomPropertyType.Text)
            {
                var vTlpwstr1 = new VTLPWSTR { Text = p.GetValue<string>() };
                customDocumentProperty.AppendChild(vTlpwstr1);
            }
            else if (p.Type == XLCustomPropertyType.Date)
            {
                var vTFileTime1 = new VTFileTime
                {
                    Text =
                        p.GetValue<DateTime>().ToUniversalTime().ToString(
                            "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'")
                };
                customDocumentProperty.AppendChild(vTFileTime1);
            }
            else if (p.Type == XLCustomPropertyType.Number)
            {
                customDocumentProperty.AppendChild(CreateNumber(p));
            }
            else
            {
                var vTBool1 = new VTBool { Text = p.GetValue<bool>().ToString().ToLower() };
                customDocumentProperty.AppendChild(vTBool1);
            }
            properties.AppendChild(customDocumentProperty);
        }

        foreach (var unmodeled in workbook.CustomPropertiesInternal.UnmodeledPropertiesToSave)
        {
            propertyId++;
            var customDocumentProperty = (CustomDocumentProperty)unmodeled.CloneNode(true);
            customDocumentProperty.PropertyId = propertyId;
            properties.AppendChild(customDocumentProperty);
        }

        customFilePropertiesPart.Properties = properties;
    }

    private static OpenXmlLeafTextElement CreateNumber(IXLCustomProperty property)
    {
        return property.Value switch
        {
            sbyte v => new VTByte { Text = v.ToInvariantString() },
            short v => new VTShort { Text = v.ToInvariantString() },
            int v => new VTInt32 { Text = v.ToInvariantString() },
            long v => new VTInt64 { Text = v.ToInvariantString() },
            byte v => new VTUnsignedByte { Text = v.ToInvariantString() },
            ushort v => new VTUnsignedShort { Text = v.ToInvariantString() },
            uint v => new VTUnsignedInt32 { Text = v.ToInvariantString() },
            ulong v => new VTUnsignedInt64 { Text = v.ToInvariantString() },
            _ => new VTDouble { Text = property.GetValue<double>().ToInvariantString() }
        };
    }
}
