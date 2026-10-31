using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads the records part of a pivot cache into the fields the definition declared. The counterpart of
/// <see cref="PivotTableCacheRecordsPartWriter"/>, so the two can be tested against each other.
/// </summary>
internal static class PivotTableCacheRecordsPartReader
{
    internal static void Read(PivotCacheRecords recordsPart, XLPivotCache pivotCache)
    {
        // Number of records can be rather large, preallocate capacity to avoid reallocation.
        var recordCount = recordsPart.Count?.Value is not null
            ? checked((int)recordsPart.Count.Value)
            : 0;
        pivotCache.AllocateRecordCapacity(recordCount);

        // Non-database fields (calculated and grouping) don't have records — only database fields do.
        var databaseFieldCount = pivotCache.DatabaseFieldCount;
        foreach (var record in recordsPart.Elements<PivotCacheRecord>())
        {
            var recordColumns = record.ChildElements.Count;
            if (recordColumns != databaseFieldCount)
                throw PartStructureException.IncorrectElementsCount();

            // Map record column index to field index, skipping non-database fields.
            var recordColIdx = 0;
            for (var fieldIdx = 0; fieldIdx < pivotCache.FieldCount; ++fieldIdx)
            {
                if (pivotCache.IsNonDatabaseField(fieldIdx))
                    continue;

                var fieldValues = pivotCache.GetFieldValues(fieldIdx);
                var recordItem = record.ElementAt(recordColIdx);
                recordColIdx++;

                // Don't add values to the shared items of a cache when record value is added, because we want 1:1
                // read/write. Read them from definition. Whatever is in shared items now should be written out,
                // unless there is a cache refresh. Basically trust the author of the workbook that it is valid.
                AddRecordItem(fieldValues, recordItem);
            }
        }
    }

    private static void AddRecordItem(XLPivotCacheValues fieldValues, OpenXmlElement recordItem)
    {
        // Only a record can hold an index (`x`) into the shared items; every other item is a value.
        if (recordItem is FieldItem indexItem)
            fieldValues.AddIndex(GetFieldIndex(indexItem, fieldValues.SharedCount));
        else
            PivotCacheValueCodec.ReadValue(fieldValues, recordItem);
    }

    private static uint GetFieldIndex(FieldItem indexItem, int sharedCount)
    {
        var index = indexItem.Val?.Value ?? throw PartStructureException.MissingAttribute();
        if (index >= sharedCount)
            throw PartStructureException.IncorrectAttributeValue();

        return index;
    }
}
