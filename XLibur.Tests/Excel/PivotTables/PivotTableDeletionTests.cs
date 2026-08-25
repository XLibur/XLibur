using System;
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

    private static string[] PartsUnder(MemoryStream package, string prefix)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        return archive.Entries
            .Where(e => e.FullName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .Select(e => e.FullName)
            .ToArray();
    }

    #endregion
}
