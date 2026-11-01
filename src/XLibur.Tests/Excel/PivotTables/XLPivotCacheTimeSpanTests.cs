using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TUnit.Assertions.Enums;
using XLibur.Excel;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// A pivot cache stores a <see cref="TimeSpan"/> value as a <c>d</c> item. Excel writes the
/// serial as a plain OLE Automation date, without the 1900 leap-year shift that applies to
/// cell dates: 14:30 (serial 0.604) is <c>1899-12-30T14:30:00</c>, 25:00 (serial 1.042) is
/// <c>1899-12-31T01:00:00</c> and serial 59.5 is <c>1900-02-27T12:00:00</c>. The shared items
/// and the records must agree with each other and with Excel (#628).
/// </summary>
public class XLPivotCacheTimeSpanTests
{
    private const string ExcelFixture = @"Other\PivotTableReferenceFiles\TimeSpanCacheValues\input.xlsx";

    [Test]
    public async Task TimeSpan_shared_items_and_records_are_written_as_Excel_writes_them()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Time";
        ws.Cell("B1").Value = "Amount";
        TimeSpan[] times =
        [
            new(14, 30, 0),
            TimeSpan.Zero,
            new(23, 59, 59),
            new(25, 0, 0),
        ];
        for (var i = 0; i < times.Length; i++)
        {
            ws.Cell(i + 2, 1).Value = times[i];
            ws.Cell(i + 2, 2).Value = (i + 1) * 10;
        }

        var pt = ws.PivotTables.Add("pt", ws.Cell("E2"), ws.Range("A1:B5"));
        pt.RowLabels.Add("Time");
        pt.Values.Add("Amount");

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        var saved = ReadCacheField(ms, "Time");
        DateTime[] expected =
        [
            new(1899, 12, 30, 14, 30, 0),
            new(1899, 12, 30, 0, 0, 0),
            new(1899, 12, 30, 23, 59, 59),
            new(1899, 12, 31, 1, 0, 0),
        ];

        await Assert.That(saved.SharedItems).IsEquivalentTo(expected, CollectionOrdering.Matching);
        await Assert.That(saved.Records).IsEquivalentTo(expected, CollectionOrdering.Matching);
        await Assert.That(saved.MinDate).IsEqualTo(new DateTime(1899, 12, 30, 0, 0, 0));
        await Assert.That(saved.MaxDate).IsEqualTo(new DateTime(1899, 12, 31, 1, 0, 0));
        await Assert.That(saved.ContainsDate).IsTrue();
        await Assert.That(saved.ContainsNonDate).IsFalse();
    }

    /// <summary>
    /// The fixture was saved by Excel from a pivot table over two time-formatted columns. Refreshing
    /// the cache in XLibur rebuilds both columns from the cells, and the result must match the items
    /// Excel wrote, including the serials below 61 where the leap-year shift would otherwise apply.
    /// </summary>
    [Test]
    [Arguments("Time")]
    [Arguments("Duration")]
    public async Task Refreshed_TimeSpan_cache_field_matches_Excel(string fieldName)
    {
        using var fixture = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(ExcelFixture));
        var excel = ReadCacheField(fixture, fieldName);

        fixture.Position = 0;
        using var wb = new XLWorkbook(fixture);
        var sourceColumn = fieldName == "Time" ? "A" : "B";
        await Assert.That(wb.Worksheet("Data").Cell(sourceColumn + "2").DataType).IsEqualTo(XLDataType.TimeSpan);

        wb.PivotCaches.Single().Refresh();
        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        var saved = ReadCacheField(ms, fieldName);

        await Assert.That(saved.SharedItems).IsEquivalentTo(excel.SharedItems, CollectionOrdering.Matching);
        await Assert.That(saved.MinDate).IsEqualTo(excel.MinDate);
        await Assert.That(saved.MaxDate).IsEqualTo(excel.MaxDate);
        await Assert.That(saved.ContainsDate).IsEqualTo(excel.ContainsDate);
        await Assert.That(saved.ContainsNonDate).IsEqualTo(excel.ContainsNonDate);

        // Excel's records index the shared items; XLibur writes the date itself. Either way each
        // record must resolve to the same date.
        await Assert.That(saved.Records).IsEquivalentTo(excel.Records, CollectionOrdering.Matching);
    }

    /// <summary>
    /// A duration between 60 and 61 days lands on 1900-02-28, the date whose serial the cell-date
    /// conversion refuses (serial 60 is Excel's fictional 1900-02-29). The cache date must be written
    /// as is, not sent back through that conversion.
    /// </summary>
    [Test]
    public async Task TimeSpan_on_1900_02_28_is_saved()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Time";
        ws.Cell("A2").Value = TimeSpan.FromDays(60.5);
        ws.Cell("B1").Value = "Amount";
        ws.Cell("B2").Value = 1;

        var pt = ws.PivotTables.Add("pt", ws.Cell("E2"), ws.Range("A1:B2"));
        pt.RowLabels.Add("Time");
        pt.Values.Add("Amount");

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        var saved = ReadCacheField(ms, "Time");
        DateTime[] expected = [new(1900, 2, 28, 12, 0, 0)];
        await Assert.That(saved.SharedItems).IsEquivalentTo(expected, CollectionOrdering.Matching);
        await Assert.That(saved.Records).IsEquivalentTo(expected, CollectionOrdering.Matching);
    }

    private static SavedCacheField ReadCacheField(Stream stream, string fieldName)
    {
        stream.Position = 0;
        using var document = SpreadsheetDocument.Open(stream, false);
        var definitionPart = document.WorkbookPart!.PivotTableCacheDefinitionParts.Single();
        var cacheFields = definitionPart.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().ToList();
        var fieldIndex = cacheFields.FindIndex(f => f.Name == fieldName);
        var sharedItems = cacheFields[fieldIndex].SharedItems!;
        var sharedDates = sharedItems.Elements<DateTimeItem>().Select(d => d.Val!.Value).ToList();

        var records = new List<DateTime>();
        foreach (var record in definitionPart.PivotTableCacheRecordsPart!.PivotCacheRecords!.Elements<PivotCacheRecord>())
        {
            var item = record.ChildElements[fieldIndex];
            records.Add(item switch
            {
                DateTimeItem d => d.Val!.Value,
                FieldItem x => sharedDates[checked((int)x.Val!.Value)],
                _ => throw new InvalidOperationException($"Unexpected record item {item.LocalName}."),
            });
        }

        return new SavedCacheField(
            sharedDates,
            records,
            sharedItems.MinDate?.Value,
            sharedItems.MaxDate?.Value,
            sharedItems.ContainsDate?.Value ?? false,
            sharedItems.ContainsNonDate?.Value ?? true);
    }

    private sealed record SavedCacheField(
        List<DateTime> SharedItems,
        List<DateTime> Records,
        DateTime? MinDate,
        DateTime? MaxDate,
        bool ContainsDate,
        bool ContainsNonDate);
}
