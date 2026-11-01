using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TUnit.Assertions.Enums;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// A pivot cache value is spelled the same way by the shared items and the records, and read back
/// the same way (spec 41). The records part is checked against the definition that declares it.
/// </summary>
public class PivotCacheValueCodecTests
{
    private static readonly XLCellValue[] EveryValueType =
    [
        Blank.Value,
        true,
        1.5,
        "text",
        XLError.DivisionByZero,
        XLError.NoValueAvailable,
        XLError.CellReference,
        new DateTime(2024, 1, 2),
    ];

    private static MemoryStream SavePivotOver(XLCellValue[] values)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Field";
        ws.Cell("B1").Value = "Amount";
        for (var i = 0; i < values.Length; i++)
        {
            ws.Cell(i + 2, 1).Value = values[i];
            ws.Cell(i + 2, 2).Value = i + 1;
        }

        var pt = ws.PivotTables.Add("pt", ws.Cell("E2"), ws.Range(1, 1, values.Length + 1, 2));
        pt.RowLabels.Add("Field");
        pt.Values.Add("Amount");

        var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms;
    }

    [Test]
    public async Task EveryValueTypeSurvivesASaveAndReload()
    {
        using var saved = SavePivotOver(EveryValueType);

        using var reloaded = new XLWorkbook(saved);
        var cache = (XLPivotCache)reloaded.PivotCaches.Single();
        var fieldValues = cache.GetFieldValues(0);

        await Assert.That(fieldValues.GetCellValues().ToArray())
            .IsEquivalentTo(EveryValueType, CollectionOrdering.Matching);
        await Assert.That(fieldValues.SharedItems.GetCellValues().ToArray())
            .IsEquivalentTo(EveryValueType, CollectionOrdering.Matching);
    }

    [Test]
    public async Task EveryValueTypeSurvivesASecondSave()
    {
        using var first = SavePivotOver(EveryValueType);
        using var wb = new XLWorkbook(first);
        using var second = new MemoryStream();
        wb.SaveAs(second);

        using var reloaded = new XLWorkbook(second);
        var fieldValues = ((XLPivotCache)reloaded.PivotCaches.Single()).GetFieldValues(0);
        await Assert.That(fieldValues.GetCellValues().ToArray())
            .IsEquivalentTo(EveryValueType, CollectionOrdering.Matching);
    }

    [Test]
    public async Task ErrorValuesAreWrittenAsErrorElementsInSharedItemsAndRecords()
    {
        using var saved = SavePivotOver(EveryValueType);

        var records = saved.ReadPartUnder("pivotCache/pivotCacheRecords");
        var definition = saved.ReadPartUnder("pivotCache/pivotCacheDefinition");
        await Assert.That(records).Contains("<e v=\"#DIV/0!\"");
        await Assert.That(records).Contains("<e v=\"#N/A\"");
        await Assert.That(records).Contains("<e v=\"#REF!\"");
        await Assert.That(records).DoesNotContain("<b v=\"#");
        await Assert.That(definition).Contains("<x:e v=\"#DIV/0!\"");
    }

    /// <summary>
    /// After any save, the fields the definition declares as database fields are the columns every
    /// record carries.
    /// </summary>
    [Test]
    public async Task DefinitionAndRecordsAgreeAboutTheNumberOfColumns()
    {
        using var saved = SavePivotOver(EveryValueType);

        saved.Position = 0;
        using var document = SpreadsheetDocument.Open(saved, false);
        var definitionPart = document.WorkbookPart!.PivotTableCacheDefinitionParts.Single();
        var databaseFields = definitionPart.PivotCacheDefinition!.CacheFields!.Elements<CacheField>()
            .Count(f => f.DatabaseField?.Value != false);
        var records = definitionPart.PivotTableCacheRecordsPart!.PivotCacheRecords!
            .Elements<PivotCacheRecord>().ToList();

        await Assert.That(records.Count).IsEqualTo(EveryValueType.Length);
        foreach (var record in records)
            await Assert.That(record.ChildElements.Count).IsEqualTo(databaseFields);
    }

    [Test]
    public async Task ItemsToRetainSetBackToAutomaticIsWrittenAsAutomatic()
    {
        using var first = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Field";
            ws.Cell("A2").Value = "x";
            ws.Cell("B1").Value = "Amount";
            ws.Cell("B2").Value = 1;
            var pt = ws.PivotTables.Add("pt", ws.Cell("E2"), ws.Range("A1:B2"));
            pt.RowLabels.Add("Field");
            pt.Values.Add("Amount");
            pt.PivotCache.ItemsToRetainPerField = XLItemsToRetain.None;
            wb.SaveAs(first);
        }

        using var second = new MemoryStream();
        using (var wb = new XLWorkbook(first))
        {
            var cache = wb.PivotCaches.Single();
            await Assert.That(cache.ItemsToRetainPerField).IsEqualTo(XLItemsToRetain.None);
            cache.ItemsToRetainPerField = XLItemsToRetain.Automatic;
            wb.SaveAs(second);
        }

        using var reloaded = new XLWorkbook(second);
        await Assert.That(reloaded.PivotCaches.Single().ItemsToRetainPerField).IsEqualTo(XLItemsToRetain.Automatic);
    }
}
