using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XLibur.Excel;

namespace XLibur.Tests.Excel.PivotTables;

/// <summary>
/// Deleting a pivot table has to take its part with it.
/// </summary>
/// <remarks>
/// Deleting a worksheet already took its pivot table parts along, but deleting a pivot table on its
/// own left the part in the package — still carrying a <c>cacheId</c> pointing into a
/// <c>pivotCaches</c> element the same save then rebuilt without it. The result validates as a
/// dangling reference and Excel offers to repair it. This is unrelated to slicers; it was found
/// while building the slicer cascade, whose whole purpose is not to leave orphans of that kind.
/// </remarks>
public class PivotTableDeletionTests
{
    private const string Fixture = @"Other\PivotTableReferenceFiles\ChartsheetAndPivotTable.xlsx";

    [Test]
    public async Task Deleting_a_pivot_table_removes_its_part()
    {
        using var saved = new MemoryStream();
        string sheetName;

        using (var wb = Load())
        {
            var worksheet = wb.Worksheets.First(ws => ws.PivotTables.Any());
            sheetName = worksheet.Name;
            var pivotTableName = worksheet.PivotTables.First().Name;

            worksheet.PivotTables.Delete(pivotTableName);
            await Assert.That(worksheet.PivotTables.Any()).IsFalse();

            wb.SaveAs(saved);
        }

        await Assert.That(PartsUnder(saved, "xl/pivotTables/")).IsEmpty();
        await Assert.That(sheetName).IsNotEmpty();
    }

    [Test]
    public async Task Deleting_a_pivot_table_leaves_no_dangling_cache_reference()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var worksheet = wb.Worksheets.First(ws => ws.PivotTables.Any());
            worksheet.PivotTables.Delete(worksheet.PivotTables.First().Name);
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2010)
            .Validate(doc)
            .Select(error => $"{error.Path?.XPath}: {error.Description}")
            .Where(error => error.Contains("pivot", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Scoped to the pivot references rather than asserting the whole package is clean: this
        // fixture's chartsheet carries a c:showDLblsOverMax the Office2010 validator rejects, which
        // predates all of this and has nothing to do with deleting a pivot table. The unfiltered
        // assertion lives in SlicerWriteTests against a fixture that validates clean to begin with.
        await Assert.That(string.Join(Environment.NewLine, errors)).IsEmpty();
    }

    [Test]
    public async Task A_workbook_that_lost_its_pivot_table_still_opens()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var worksheet = wb.Worksheets.First(ws => ws.PivotTables.Any());
            worksheet.PivotTables.Delete(worksheet.PivotTables.First().Name);
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);

        await Assert.That(reloaded.Worksheets.SelectMany(ws => ws.PivotTables)).IsEmpty();
    }

    [Test]
    public async Task Deleting_a_pivot_table_takes_its_timelines_with_it()
    {
        using var saved = new MemoryStream();

        using (var wb = TimelineWorkbook())
        {
            var pivotSheet = wb.Worksheet("Pivot");
            await Assert.That(pivotSheet.Timelines.Count).IsEqualTo(1);

            pivotSheet.PivotTables.Delete(pivotSheet.PivotTables.Single().Name);

            // The cache served only that pivot table, so the timeline has nothing left to filter.
            await Assert.That(pivotSheet.Timelines.Count).IsEqualTo(0);

            wb.SaveAs(saved);
        }

        var entries = EntryNames(saved);

        // The part, the cache part and the #N/A defined name all go, or the saved file has an orphan
        // Excel will offer to repair.
        await Assert.That(entries.Any(n => n.StartsWith("xl/timelines/", StringComparison.Ordinal))).IsFalse();
        await Assert.That(entries.Any(n => n.StartsWith("xl/timelineCaches/", StringComparison.Ordinal))).IsFalse();

        var workbookXml = ReadPart(saved, "xl/workbook.xml");
        await Assert.That(workbookXml).DoesNotContain("timelineCacheRef");
        await Assert.That(workbookXml).DoesNotContain("ВстроеннаяВременнаяШкала_Date");

        // And the drawing no longer asks Excel to draw a band the package does not define.
        await Assert.That(ReadPart(saved, "xl/drawings/drawing1.xml")).DoesNotContain("timeslicer");

        // The worksheet's own extLst reference has to go too, or a stale relationship id sits in an
        // extension list the XSD does not check. Pivot is sheet1.xml in this fixture — the sheetIds
        // are crossed, so this is not the same part sheetIds would suggest.
        var pivotSheetXml = ReadPart(saved, "xl/worksheets/sheet1.xml");
        await Assert.That(pivotSheetXml).DoesNotContain("timelineRef");

        // Stronger still: the whole <ext> should be pruned once its list empties, not just the ref
        // inside it.
        await Assert.That(pivotSheetXml).DoesNotContain("7E03D99C-DC04-49d9-9315-930204A7B6E9");
    }

    [Test]
    public async Task A_workbook_whose_pivot_table_was_deleted_is_schema_valid()
    {
        using var saved = new MemoryStream();

        using (var wb = TimelineWorkbook())
        {
            var pivotSheet = wb.Worksheet("Pivot");
            pivotSheet.PivotTables.Delete(pivotSheet.PivotTables.Single().Name);
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2013)
            .Validate(doc)
            .Select(error => $"{error.Path?.XPath}: {error.Description}")
            .ToList();

        await Assert.That(errors).IsEmpty();
    }

    #region Helpers

    private static XLWorkbook Load()
    {
        using var resource = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Fixture));
        var stream = new MemoryStream();
        resource.CopyTo(stream);
        stream.Position = 0;

        // The workbook reads its original stream again on save, so it cannot be disposed here.
        return new XLWorkbook(stream);
    }

    private static XLWorkbook TimelineWorkbook()
    {
        using var source = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"TryToLoad\Timelines_Missing_21232.xlsx"));
        var ms = new MemoryStream();
        source.CopyTo(ms);
        ms.Position = 0;
        return new XLWorkbook(ms);
    }

    private static string[] PartsUnder(MemoryStream package, string prefix)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        return archive.Entries
            .Where(e => e.FullName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .Select(e => e.FullName)
            .ToArray();
    }

    private static List<string> EntryNames(MemoryStream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        return archive.Entries.Select(e => e.FullName).ToList();
    }

    private static string ReadPart(MemoryStream package, string partName)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.GetEntry(partName);
        if (entry is null)
            return string.Empty;

        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    #endregion
}
