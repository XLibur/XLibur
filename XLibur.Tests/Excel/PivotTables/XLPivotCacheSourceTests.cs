using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Coordinates;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// The public source surface of a pivot cache — what it reads from, and re-pointing it.
/// </summary>
/// <remarks>
/// The three-way distinction matters to any caller that rewrites a workbook after generating into
/// it: a source that cannot be read at all is not the same as one that can be read but no longer
/// resolves, and treating them alike turns an untouched pivot into a reported error.
/// </remarks>
public class XLPivotCacheSourceTests
{
    private static readonly object[][] Data =
    [
        ["Name", "Count"],
        ["Cake", 1.0],
        ["Pie", 2.0],
        ["Tart", 3.0],
    ];

    private static XLWorkbook WorkbookWithData(out IXLWorksheet sheet)
    {
        var wb = new XLWorkbook();
        sheet = wb.AddWorksheet("Data");
        sheet.FirstCell().InsertData(Data);
        return wb;
    }

    [Test]
    public async Task A_range_source_reports_its_range()
    {
        using var wb = WorkbookWithData(out var sheet);
        var range = sheet.Range("A1:B4");

        var cache = wb.PivotCaches.Add(range!);

        await Assert.That(cache.SourceKind).IsEqualTo(XLPivotSourceKind.Range);
        await Assert.That(cache.SourceRange).IsNotNull();
        await Assert.That(cache.SourceRange!.RangeAddress.ToString()).IsEqualTo("A1:B4");
        await Assert.That(cache.SourceWorksheet!.Name).IsEqualTo("Data");

        // A range source has no name of its own.
        await Assert.That(cache.SourceName).IsNull();
    }

    [Test]
    public async Task A_table_source_reports_its_name_and_resolved_sheet()
    {
        using var wb = WorkbookWithData(out var sheet);
        var table = sheet.Range("A1:B4").CreateTable("SourceTable");

        // Adding a cache for a range that exactly matches a table records the table by name.
        var cache = wb.PivotCaches.Add(table.AsRange());

        await Assert.That(cache.SourceKind).IsEqualTo(XLPivotSourceKind.Name);
        await Assert.That(cache.SourceName).IsEqualTo("SourceTable");
        await Assert.That(cache.SourceWorksheet).IsNotNull();
        await Assert.That(cache.SourceWorksheet!.Name).IsEqualTo("Data");

        // A named source has no rectangle of its own — the table owns one, and it can move.
        await Assert.That(cache.SourceRange).IsNull();
    }

    [Test]
    public async Task A_named_source_that_no_longer_resolves_keeps_its_name_and_has_no_sheet()
    {
        using var wb = WorkbookWithData(out var sheet);
        var table = sheet.Range("A1:B4").CreateTable("SourceTable");
        var cache = wb.PivotCaches.Add(table.AsRange());

        sheet.Tables.Remove("SourceTable");

        // This is the case a nullable-only surface could not tell from an unreadable source kind:
        // the name is still what the file recorded, there is just nothing behind it any more.
        await Assert.That(cache.SourceKind).IsEqualTo(XLPivotSourceKind.Name);
        await Assert.That(cache.SourceName).IsEqualTo("SourceTable");
        await Assert.That(cache.SourceWorksheet).IsNull();
        await Assert.That(cache.SourceRange).IsNull();
    }

    [Test]
    public async Task SetSourceRange_re_points_the_cache()
    {
        using var wb = WorkbookWithData(out var sheet);
        var cache = wb.PivotCaches.Add(sheet.Range("A1:B2"));

        cache.SetSourceRange(sheet.Range("A1:B4"));

        await Assert.That(cache.SourceKind).IsEqualTo(XLPivotSourceKind.Range);
        await Assert.That(cache.SourceRange!.RangeAddress.ToString()).IsEqualTo("A1:B4");
    }

    [Test]
    public async Task SetSourceRange_then_Refresh_reads_the_new_range()
    {
        using var wb = WorkbookWithData(out var sheet);
        var cache = wb.PivotCaches.Add(sheet.Range("A1:A4"));

        await Assert.That(cache.FieldNames.Count).IsEqualTo(1);

        cache.SetSourceRange(sheet.Range("A1:B4"));
        cache.Refresh();

        await Assert.That(cache.FieldNames).Contains("Name");
        await Assert.That(cache.FieldNames).Contains("Count");
    }

    [Test]
    public async Task SetSourceRange_returns_the_cache_for_chaining()
    {
        using var wb = WorkbookWithData(out var sheet);
        var cache = wb.PivotCaches.Add(sheet.Range("A1:B2"));

        await Assert.That(cache.SetSourceRange(sheet.Range("A1:B4"))).IsSameReferenceAs(cache);
    }

    [Test]
    public async Task A_range_source_whose_sheet_is_gone_has_no_range_and_no_sheet()
    {
        using var wb = WorkbookWithData(out var sheet);
        var cache = wb.PivotCaches.Add(sheet.Range("A1:B4"));

        wb.Worksheets.Delete("Data");

        // Still a range source — the file recorded a rectangle on a sheet — but nothing resolves.
        await Assert.That(cache.SourceKind).IsEqualTo(XLPivotSourceKind.Range);
        await Assert.That(cache.SourceRange).IsNull();
        await Assert.That(cache.SourceWorksheet).IsNull();
    }

    /// <remarks>
    /// Built through the internal source types rather than loaded from a file, because a workbook
    /// carrying a live OLAP connection or a consolidation is not something this suite can author.
    /// The mapping is what is under test, and it is worth testing directly: no caller in this
    /// repository exercises these kinds, so a wrong answer here fails nothing else.
    /// </remarks>
    [Test]
    [Arguments(XLPivotSourceKind.Scenario)]
    [Arguments(XLPivotSourceKind.Connection)]
    [Arguments(XLPivotSourceKind.Consolidation)]
    [Arguments(XLPivotSourceKind.ExternalWorkbook)]
    public async Task A_source_XLibur_cannot_read_reports_its_kind_and_resolves_to_nothing(
        XLPivotSourceKind expected)
    {
        using var wb = WorkbookWithData(out var sheet);
        IXLPivotSource source = expected switch
        {
            XLPivotSourceKind.Scenario => new XLPivotSourceScenario(),
            XLPivotSourceKind.Connection => new XLPivotSourceConnection(1),
            XLPivotSourceKind.Consolidation => new XLPivotSourceConsolidation(),
            _ => new XLPivotSourceExternalWorkbook("rId1", SheetArea.From(sheet.Range("A1:B4"))),
        };

        var cache = new XLPivotCache(source, wb);

        await Assert.That(cache.SourceKind).IsEqualTo(expected);

        // Null for a different reason than a name that stopped resolving: there is nothing here
        // XLibur could read even in principle. SourceKind is what tells the two apart.
        await Assert.That(cache.SourceRange).IsNull();
        await Assert.That(cache.SourceWorksheet).IsNull();
        await Assert.That(cache.SourceName).IsNull();
    }
}
