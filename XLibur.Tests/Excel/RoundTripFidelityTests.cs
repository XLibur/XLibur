using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;
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
    public async Task Slicers_and_their_caches_survive_a_round_trip()
    {
        using var saved = LoadAndSave(@"TryToLoad\SlicersOnPivotAndTable.xlsx");

        // slicer1 hangs off the table on sheet1, slicer2 off the pivot table on sheet2. Their caches
        // are cross-numbered: slicerCache1 serves the pivot slicer, slicerCache2 the table slicer.
        await Assert.That(PartExists(saved, "xl/slicers/slicer1.xml")).IsTrue();
        await Assert.That(PartExists(saved, "xl/slicers/slicer2.xml")).IsTrue();
        await Assert.That(PartExists(saved, "xl/slicerCaches/slicerCache1.xml")).IsTrue();
        await Assert.That(PartExists(saved, "xl/slicerCaches/slicerCache2.xml")).IsTrue();
    }

    [Test]
    public async Task Slicer_styling_XLibur_does_not_model_survives_a_round_trip()
    {
        using var saved = LoadAndSave(@"TryToLoad\SlicersOnPivotAndTable.xlsx");

        // The pivot slicer carries a renamed caption, a non-default built-in style and a single
        // selected item. None of that is modelled, so it only survives if the part is left alone.
        var pivotSlicer = ReadPart(saved, "xl/slicers/slicer2.xml");
        await Assert.That(pivotSlicer).Contains("caption=\"Region filter\"");
        await Assert.That(pivotSlicer).Contains("style=\"SlicerStyleDark3\"");

        // s="1" on a single <i> is the selection. The table slicer's cache instead carries an
        // x15:tableSlicerCache extension, which is the other of the two binding paths.
        var pivotCache = ReadPart(saved, "xl/slicerCaches/slicerCache1.xml");
        await Assert.That(pivotCache).Contains("<i x=\"0\" s=\"1\"/>");

        var tableCache = ReadPart(saved, "xl/slicerCaches/slicerCache2.xml");
        await Assert.That(tableCache).Contains("tableSlicerCache");
    }

    [Test]
    public async Task Slicer_references_survive_in_the_worksheet_xml()
    {
        using var saved = LoadAndSave(@"TryToLoad\SlicersOnPivotAndTable.xlsx");

        // Same trap as the form controls above: the worksheet part is rebuilt from the model on
        // every save, so surviving slicer parts are orphans unless the sheet keeps pointing at them.
        // Excel uses a different extension URI for a table slicer than for a pivot slicer.
        var tableSheet = ReadPart(saved, "xl/worksheets/sheet1.xml");
        await Assert.That(tableSheet).Contains("{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}");
        await Assert.That(tableSheet).Contains("slicerList");

        var pivotSheet = ReadPart(saved, "xl/worksheets/sheet2.xml");
        await Assert.That(pivotSheet).Contains("{A8765BA9-456A-4dab-B4F3-ACF838C121DE}");
        await Assert.That(pivotSheet).Contains("slicerList");

        // The extLst points at the slicer part by relationship id, so the relationship has to live too.
        await Assert.That(ReadPart(saved, "xl/worksheets/_rels/sheet1.xml.rels"))
            .Contains("../slicers/slicer1.xml");
        await Assert.That(ReadPart(saved, "xl/worksheets/_rels/sheet2.xml.rels"))
            .Contains("../slicers/slicer2.xml");
    }

    [Test]
    public async Task Slicer_cache_references_survive_in_the_workbook_xml()
    {
        using var saved = LoadAndSave(@"TryToLoad\SlicersOnPivotAndTable.xlsx");

        var workbookXml = ReadPart(saved, "xl/workbook.xml");

        // Two separate registries: x14:slicerCaches for the pivot slicer, x15:slicerCaches for the
        // table slicer. Losing either one orphans a cache part that still exists on disk.
        await Assert.That(workbookXml).Contains("{BBE1A952-AA13-448e-AADC-164F8A28A991}");
        await Assert.That(workbookXml).Contains("{46BE6895-7355-4a93-B00E-2C351335B9C9}");

        // Excel writes a #N/A defined name per slicer cache. XLibur models defined names and rewrites
        // the whole block, so these are the most exposed part of the whole arrangement.
        await Assert.That(workbookXml).Contains("Slicer_Region");
        await Assert.That(workbookXml).Contains("Slicer_Region1");

        await Assert.That(ReadPart(saved, "xl/_rels/workbook.xml.rels"))
            .Contains("slicerCaches/slicerCache1.xml");
        await Assert.That(ReadPart(saved, "xl/_rels/workbook.xml.rels"))
            .Contains("slicerCaches/slicerCache2.xml");
    }

    [Test]
    public async Task Slicers_survive_an_unrelated_edit()
    {
        // The user story is "edit a sheet that has nothing to do with the slicer, save, still there".
        using var stream = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"TryToLoad\SlicersOnPivotAndTable.xlsx"));
        using var saved = new MemoryStream();

        using (var wb = new XLWorkbook(stream))
        {
            wb.Worksheet("Data").Cell("E1").Value = "touched";
            wb.SaveAs(saved);
        }

        await Assert.That(PartExists(saved, "xl/slicers/slicer1.xml")).IsTrue();
        await Assert.That(PartExists(saved, "xl/slicers/slicer2.xml")).IsTrue();
        await Assert.That(ReadPart(saved, "xl/worksheets/sheet1.xml")).Contains("slicerList");
        await Assert.That(ReadPart(saved, "xl/worksheets/sheet2.xml")).Contains("slicerList");
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

    [Test]
    public async Task A_drawing_holding_only_slicers_survives_a_round_trip_byte_for_byte()
    {
        // Both sheets of this fixture carry a drawing whose only content is a slicer frame — no
        // picture, no chart, no legacy shape. XLibur models none of what is in them, so by the
        // preservation rule at the top of this file the parts should come through untouched.
        //
        // Every other slicer test asserts with Contains, which cannot see a part that was rewritten
        // rather than passed through. That is why this went unnoticed: the frame keeps every element
        // and attribute, and only the serialisation changes.
        using var original = Resource(@"TryToLoad\SlicersOnPivotAndTable.xlsx");
        var before1 = PartBytes(original, "xl/drawings/drawing1.xml");
        var before2 = PartBytes(original, "xl/drawings/drawing2.xml");

        using var saved = LoadAndSave(@"TryToLoad\SlicersOnPivotAndTable.xlsx");

        // CollectionOrdering.Matching, because IsEquivalentTo ignores order by default — it holds
        // {1,2,3} equivalent to {3,2,1}. For a byte-for-byte claim that is not what is meant.
        await Assert.That(PartBytes(saved, "xl/drawings/drawing1.xml"))
            .IsEquivalentTo(before1, CollectionOrdering.Matching);
        await Assert.That(PartBytes(saved, "xl/drawings/drawing2.xml"))
            .IsEquivalentTo(before2, CollectionOrdering.Matching);
    }

    [Test]
    public async Task Deleting_every_picture_still_drops_the_emptied_drawing_part()
    {
        // The other side of the same guard. RemoveEmptyDrawingPart asks whether the drawing has any
        // children; the test above pins that asking must not rewrite the part, and this one pins
        // that the answer stays correct once XLibur has emptied the drawing itself.
        //
        // The two pull in opposite directions: the emptiness check reads the part's bytes to avoid
        // attaching its DOM, but the deletions above it happen *in* that DOM, so the bytes on disk
        // are stale by the time the question is asked. Answer from the stream in that case and an
        // emptied drawing part is left behind in the package.
        using var saved = new MemoryStream();

        using (var stream = TestHelper.GetStreamFromResource(
                   TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
        using (var wb = new XLWorkbook(stream))
        {
            var ws = wb.Worksheets.First();
            while (ws.Pictures.Count > 0)
                ws.Pictures.Delete(ws.Pictures.First());

            wb.SaveAs(saved);
        }

        // Sheet 1's drawing held two pictures and nothing else, so emptying it should take the part.
        await Assert.That(PartExists(saved, "xl/drawings/drawing1.xml")).IsFalse();

        // The other sheets kept their pictures, so their drawings must survive.
        await Assert.That(PartExists(saved, "xl/drawings/drawing2.xml")).IsTrue();

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheets.First().Pictures.Count).IsEqualTo(0);
    }

    #region Helpers

    private static MemoryStream Resource(string resourcePath)
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(resourcePath));
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms;
    }

    private static byte[] PartBytes(MemoryStream package, string partPath)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.First(e =>
            e.FullName.Equals(partPath, StringComparison.OrdinalIgnoreCase));

        using var entryStream = entry.Open();
        using var buffer = new MemoryStream();
        entryStream.CopyTo(buffer);
        return buffer.ToArray();
    }

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
