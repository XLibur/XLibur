using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// A pivot cache whose fields are grouped keeps that grouping through a save (#743).
/// </summary>
/// <remarks>
/// Grouping items of a field (by hand, or a date by months, quarters or years) adds a
/// <c>fieldGroup</c> to the cache field and, for each extra grouping level, a whole cache field
/// with <c>databaseField="0"</c> and no records. The save used to write that field back as an
/// ordinary database field and drop every <c>fieldGroup</c>, so the records it wrote had one
/// column fewer than the fields claiming one, and the saved file would not load again.
/// </remarks>
public class XLPivotCacheGroupFieldTests
{
    [Test]
    [Arguments(@"TryToLoad\LO\xlsx\column-style-autofilter.xlsx")]
    [Arguments(@"TryToLoad\LO\xlsx\pivot_table\shared-group-field.xlsx")]
    [Arguments(@"TryToLoad\LO\xlsx\pivot_table\shared-nested-dategroup.xlsx")]
    [Arguments(@"TryToLoad\LO\xlsx\tdf89139_pivot_table.xlsx")]
    [Arguments(@"TryToLoad\PivotTable_OverDateField_783.xlsx")]
    [Arguments(@"Other\ExcelAuthored\PivotGroupedByMonths.xlsx")]
    public async Task Saved_file_with_grouped_cache_fields_loads_again(string resource)
    {
        using var saved = LoadAndSave(resource);

        await Assert.That(() =>
        {
            using var reloaded = new XLWorkbook(saved);
        }).ThrowsNothing();
    }

    [Test]
    [Arguments(@"TryToLoad\LO\xlsx\column-style-autofilter.xlsx")]
    [Arguments(@"TryToLoad\LO\xlsx\pivot_table\shared-group-field.xlsx")]
    [Arguments(@"TryToLoad\LO\xlsx\pivot_table\shared-nested-dategroup.xlsx")]
    [Arguments(@"TryToLoad\LO\xlsx\tdf89139_pivot_table.xlsx")]
    [Arguments(@"TryToLoad\PivotTable_OverDateField_783.xlsx")]
    [Arguments(@"Other\ExcelAuthored\PivotGroupedByMonths.xlsx")]
    public async Task Save_keeps_grouping_fields_and_their_field_groups(string resource)
    {
        using var source = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resource));
        var sourceCaches = ReadCaches(source);

        using var saved = LoadAndSave(resource);
        var savedCaches = ReadCaches(saved);

        // Every cache a pivot table uses is written again, under the relationship it was read from.
        await Assert.That(savedCaches.Keys).IsEquivalentTo(sourceCaches.Keys);
        foreach (var (relId, sourceFields) in sourceCaches)
        {
            var savedCache = savedCaches[relId];
            await Assert.That(savedCache.Fields).IsEquivalentTo(sourceFields.Fields, TUnit.Assertions.Enums.CollectionOrdering.Matching);

            // A record has one item per database field, and only per database field.
            var databaseFieldCount = savedCache.Fields.Count(f => f.IsDatabaseField);
            await Assert.That(savedCache.RecordWidths.All(w => w == databaseFieldCount)).IsTrue();
            await Assert.That(savedCache.RecordWidths.Count).IsEqualTo(sourceFields.RecordWidths.Count);
        }
    }

    /// <summary>
    /// A refresh rebuilds the cache from the source columns, which have no grouping fields, and
    /// whose shared items the old groups no longer index. The groups go with it.
    /// </summary>
    [Test]
    public async Task Refresh_drops_grouping_fields_and_field_groups()
    {
        using var stream = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"TryToLoad\PivotTable_OverDateField_783.xlsx"));
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(stream))
        {
            var cache = (XLPivotCache)wb.PivotCaches.Single();
            cache.Refresh();

            await Assert.That(cache.FieldNames).IsEquivalentTo(["Date", "Number"]);
            await Assert.That(cache.DatabaseFieldCount).IsEqualTo(2);
            await Assert.That(cache.GetFieldGroupXml(0)).IsNull();
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        var savedCache = ReadCaches(saved).Values.Single();
        await Assert.That(savedCache.Fields).IsEquivalentTo(
            [new FieldShape("Date", true, null), new FieldShape("Number", true, null)],
            TUnit.Assertions.Enums.CollectionOrdering.Matching);
        await Assert.That(savedCache.RecordWidths.All(w => w == 2)).IsTrue();
    }

    private static MemoryStream LoadAndSave(string resource)
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resource));
        using var wb = new XLWorkbook(stream);
        var saved = new MemoryStream();
        wb.SaveAs(saved);
        saved.Position = 0;
        return saved;
    }

    private static Dictionary<string, CacheShape> ReadCaches(Stream stream)
    {
        var caches = new Dictionary<string, CacheShape>();
        using (var document = SpreadsheetDocument.Open(stream, false))
        {
            var workbookPart = document.WorkbookPart!;
            foreach (var pivotCache in workbookPart.Workbook!.PivotCaches?.Elements<PivotCache>() ?? [])
            {
                var part = (PivotTableCacheDefinitionPart)workbookPart.GetPartById(pivotCache.Id!.Value!);
                var fields = part.PivotCacheDefinition!.CacheFields!.Elements<CacheField>()
                    .Select(f => new FieldShape(
                        f.Name!.Value!,
                        f.DatabaseField?.Value ?? true,
                        Canonical(f.FieldGroup?.OuterXml)))
                    .ToList();
                var recordWidths = part.PivotTableCacheRecordsPart?.PivotCacheRecords?
                    .Elements<PivotCacheRecord>()
                    .Select(r => r.ChildElements.Count)
                    .ToList() ?? [];
                caches.Add(pivotCache.Id.Value!, new CacheShape(fields, recordWidths));
            }
        }

        stream.Position = 0;
        return caches;
    }

    /// <summary>
    /// The element without its namespace declarations, which say where the namespace was declared
    /// rather than what the element holds.
    /// </summary>
    private static string? Canonical(string? xml)
    {
        if (xml is null)
            return null;

        var element = XElement.Parse(xml);
        element.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
        return element.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
    }

    private sealed record FieldShape(string Name, bool IsDatabaseField, string? FieldGroupXml);

    private sealed record CacheShape(List<FieldShape> Fields, List<int> RecordWidths);
}
