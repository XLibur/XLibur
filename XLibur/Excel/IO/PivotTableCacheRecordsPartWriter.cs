using System;
using System.IO;
using System.Xml;
using XLibur.Extensions;
using DocumentFormat.OpenXml.Packaging;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.IO;

internal static class PivotTableCacheRecordsPartWriter
{
    internal static void WriteContent(PivotTableCacheRecordsPart recordsPart, XLPivotCache pivotCache)
    {
        var settings = new XmlWriterSettings
        {
            Encoding = XLHelper.NoBomUTF8
        };

        using var partStream = recordsPart.GetStream(FileMode.Create);
        using var xml = XmlWriter.Create(partStream, settings);

        xml.WriteStartDocument();
        xml.WriteStartElement("pivotCacheRecords", Main2006SsNs);
        xml.WriteAttributeString("xmlns", "r", null, RelationshipsNs);
        xml.WriteAttributeString("xmlns", "mc", null, MarkupCompatibilityNs);

        // Mark revision as ignorable extension
        xml.WriteAttributeString("mc", "Ignorable", null, "xr");
        xml.WriteAttributeString("xmlns", "xr", null, RevisionNs);

        var recordCount = pivotCache.RecordCount;
        var fieldCount = pivotCache.FieldCount;
        for (var recordIdx = 0; recordIdx < recordCount; ++recordIdx)
        {
            xml.WriteStartElement("r");
            for (var fieldIdx = 0; fieldIdx < fieldCount; ++fieldIdx)
            {
                // Non-database fields (calculated and grouping) don't have records.
                if (pivotCache.IsNonDatabaseField(fieldIdx))
                    continue;
                var fieldValues = pivotCache.GetFieldValues(fieldIdx);
                var value = fieldValues.GetValue(recordIdx);
                switch (value.Type)
                {
                    case XLPivotCacheValueType.Missing:
                        xml.WriteEmptyElement("m");
                        break;
                    case XLPivotCacheValueType.Number:
                        xml.WriteStartElement("n");
                        xml.WriteAttribute("v", value.GetNumber());
                        xml.WriteEndElement();
                        break;
                    case XLPivotCacheValueType.Boolean:
                        xml.WriteStartElement("b");
                        xml.WriteAttribute("v", value.GetBoolean());
                        xml.WriteEndElement();
                        break;
                    case XLPivotCacheValueType.Error:
                        xml.WriteStartElement("b");
                        xml.WriteAttribute("v", value.GetError().ToDisplayString());
                        xml.WriteEndElement();
                        break;
                    case XLPivotCacheValueType.String:
                        xml.WriteStartElement("s");
                        xml.WriteAttribute("v", fieldValues.GetText(value));
                        xml.WriteEndElement();
                        break;
                    case XLPivotCacheValueType.DateTime:
                        xml.WriteStartElement("d");
                        xml.WriteAttribute("v", value.GetDateTime());
                        xml.WriteEndElement();
                        break;
                    case XLPivotCacheValueType.Index:
                        xml.WriteStartElement("x");
                        xml.WriteAttribute("v", value.GetIndex());
                        xml.WriteEndElement();
                        break;
                    default:
                        throw new NotSupportedException();
                }
            }
            xml.WriteEndElement(); // r
        }

        xml.WriteEndElement(); // pivotCacheRecords
    }
}
