using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel;

/// <summary>
/// Pins down what survives a load/save round trip for content XLibur has no model for.
/// </summary>
/// <remarks>
/// XLibur saves by reopening the package it loaded and rewriting the parts it understands, so a part
/// it never touches is carried through untouched. These tests exist so that a future change which
/// starts deleting unknown parts fails here rather than silently dropping a user's chartsheets.
/// The findings are written up in docs/round-trip-fidelity.md.
/// </remarks>
public class RoundTripFidelityTests
{
    [Test]
    public async Task Chartsheets_survive_a_round_trip()
    {
        using var saved = LoadAndSave(@"Other\PivotTableReferenceFiles\ChartsheetAndPivotTable.xlsx");

        await Assert.That(PartExists(saved, "xl/chartsheets/sheet1.xml")).IsTrue();

        // The sheet entry has to stay in workbook.xml too, or the surviving part is an orphan.
        // WorkbookPartWriter reorders the sheets it models around the unsupported ones instead of
        // rewriting the list from scratch, which is what keeps this entry alive.
        var workbookXml = ReadPart(saved, "xl/workbook.xml");
        await Assert.That(workbookXml).Contains("name=\"Chart\"");
    }

    [Test]
    public async Task A_chartsheet_still_loads_after_a_round_trip()
    {
        using var saved = LoadAndSave(@"Other\PivotTableReferenceFiles\ChartsheetAndPivotTable.xlsx");

        // Reopening proves the relationships and content types survived, not just the part bytes.
        // The chartsheet is not an IXLWorksheet — it stays in the unsupported-sheet list — so only
        // the two real worksheets are counted here.
        using var wb = new XLWorkbook(saved);
        await Assert.That(wb.Worksheets.Count).IsEqualTo(2);
        await Assert.That(wb.Worksheets.Select(ws => ws.Name)).Contains("Data");
        await Assert.That(wb.Worksheets.Select(ws => ws.Name)).Contains("Pivot");
    }

    [Test]
    public async Task ActiveX_controls_survive_a_round_trip()
    {
        using var saved = LoadAndSave(@"TryToLoad\LO\xlsx\activex_checkbox.xlsx");

        await Assert.That(PartExists(saved, "xl/activeX/activeX1.xml")).IsTrue();
        await Assert.That(PartExists(saved, "xl/activeX/activeX1.bin")).IsTrue();
    }

    [Test]
    public async Task Form_control_references_survive_in_the_worksheet_xml()
    {
        using var saved = LoadAndSave(@"TryToLoad\LO\xlsx\activex_checkbox.xlsx");

        // Surviving parts are not enough: the worksheet is rewritten from the model on every save,
        // so if <controls> were dropped the activeX parts would be left as unreachable orphans.
        var sheetXml = ReadPart(saved, "xl/worksheets/sheet1.xml");
        await Assert.That(sheetXml).Contains("controls>");
        await Assert.That(sheetXml).Contains("CheckBox1343");
        await Assert.That(sheetXml).Contains("legacyDrawing");

        // The anchor inside controlPr is what positions the control, and it is nested in an
        // mc:AlternateContent that a naive rewrite would flatten or drop.
        await Assert.That(sheetXml).Contains("mc:AlternateContent");
        await Assert.That(sheetXml).Contains("438150");
    }

    [Test]
    public async Task Custom_xml_parts_survive_a_round_trip()
    {
        using var saved = LoadAndSave(@"TryToLoad\LO\xlsx\customxml.xlsx");

        await Assert.That(PartExists(saved, "customXml/item1.xml")).IsTrue();
        await Assert.That(PartExists(saved, "customXml/itemProps1.xml")).IsTrue();
    }

    [Test]
    public async Task Timelines_and_their_caches_survive_a_round_trip()
    {
        using var saved = LoadAndSave(@"TryToLoad\Timelines_Missing_21232.xlsx");

        await Assert.That(PartExists(saved, "xl/timelines/timeline1.xml")).IsTrue();
        await Assert.That(PartExists(saved, "xl/timelineCaches/timelineCache1.xml")).IsTrue();
    }

    [Test]
    public async Task Saving_to_a_new_file_also_preserves_unmodelled_parts()
    {
        // SaveAs to a fresh path copies the original package first, so the pass-through holds for
        // more than the save-in-place case.
        var target = Path.Combine(Path.GetTempPath(), $"xlibur-fidelity-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var stream = TestHelper.GetStreamFromResource(
                       TestHelper.GetResourcePath(@"Other\PivotTableReferenceFiles\ChartsheetAndPivotTable.xlsx")))
            using (var wb = new XLWorkbook(stream))
            {
                wb.Worksheets.First().Cell("A1").Value = "touched";
                wb.SaveAs(target);
            }

            using var fs = new FileStream(target, FileMode.Open, FileAccess.Read);
            using var ms = new MemoryStream();
            fs.CopyTo(ms);

            await Assert.That(PartExists(ms, "xl/chartsheets/sheet1.xml")).IsTrue();
        }
        finally
        {
            if (File.Exists(target))
                File.Delete(target);
        }
    }

    [Test]
    public async Task A_workbook_built_from_scratch_has_no_unmodelled_parts_to_lose()
    {
        // The pass-through above depends on there being an original package. There is none here,
        // which is the boundary of what preservation can do.
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Sheet1").Cell("A1").Value = 1;
            wb.SaveAs(ms, validate: true);
        }

        await Assert.That(PartExists(ms, "xl/chartsheets/sheet1.xml")).IsFalse();
    }

    #region Helpers

    private static MemoryStream LoadAndSave(string resourcePath)
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resourcePath));
        var ms = new MemoryStream();

        using (var wb = new XLWorkbook(stream))
            wb.SaveAs(ms);

        return ms;
    }

    private static bool PartExists(MemoryStream package, string partPath)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        return archive.Entries.Any(e =>
            e.FullName.Equals(partPath, StringComparison.OrdinalIgnoreCase));
    }

    private static string ReadPart(MemoryStream package, string partPath)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.First(e =>
            e.FullName.Equals(partPath, StringComparison.OrdinalIgnoreCase));

        using var entryStream = entry.Open();
        using var reader = new StreamReader(entryStream);
        return reader.ReadToEnd();
    }

    #endregion
}
