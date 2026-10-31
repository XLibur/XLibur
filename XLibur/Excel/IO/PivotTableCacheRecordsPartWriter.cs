using System;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Extensions;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.IO;

internal static class PivotTableCacheRecordsPartWriter
{
    internal static void WriteContent(PivotTableCacheRecordsPart recordsPart, XLPivotCache pivotCache)
    {
        using var xml = PartXmlWriter.Create(recordsPart);

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
                var text = value.Type == XLPivotCacheValueType.String ? fieldValues.GetText(value) : null;
                PivotCacheValueCodec.WriteRecordValue(xml, value, text);
            }
            xml.WriteEndElement(); // r
        }

        xml.WriteEndElement(); // pivotCacheRecords
    }
}
