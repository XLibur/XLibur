using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using TUnit.Assertions.Enums;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XLibur.Excel;
using XLibur.Excel.Tables;

namespace XLibur.Tests.Excel.Slicers;

/// <summary>
/// Creating slicers, editing loaded ones, and taking both apart again.
/// </summary>
/// <remarks>
/// <para>
/// A created slicer needs six things or Excel offers to repair the file: the slicer definition, the
/// worksheet's <c>extLst</c> reference to it, the cache part, the workbook's <c>extLst</c>
/// registration of that cache, a <c>#N/A</c> defined name, and a drawing anchor. Every one but the
/// anchor is asserted here; anchoring waits on spec 16's shared factory, so a slicer written today
/// saves correctly and cannot yet be seen in Excel.
/// </para>
/// <para>
/// Creation is deliberately internal until the anchor lands, which is why these tests reach for
/// <c>SlicersInternal</c> rather than a public <c>Add</c>.
/// </para>
/// </remarks>
public class SlicerWriteTests
{
    private const string Fixture = @"TryToLoad\SlicersOnPivotAndTable.xlsx";

    // ── Editing a loaded slicer (N1) ────────────────────────────────────

    [Test]
    public async Task Editing_a_loaded_slicer_keeps_everything_else_in_its_part()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var slicer = wb.Worksheet("Pivot").Slicers.Single();
            slicer.Caption = "Pick a region";
            slicer.ColumnCount = 2;
            wb.SaveAs(saved);
        }

        var xml = ReadPart(saved, "xl/slicers/slicer2.xml");

        // The edits landed.
        await Assert.That(xml).Contains("caption=\"Pick a region\"");
        await Assert.That(xml).DoesNotContain("Region filter").Because("The old caption must be replaced, not doubled up.");
        await Assert.That(xml).Contains("columnCount=\"2\"");

        // Everything XLibur does not model survived, which is the whole point of patching rather
        // than regenerating. The uid and rowHeight are attributes no XLibur API can produce.
        await Assert.That(xml).Contains("style=\"SlicerStyleDark3\"");
        await Assert.That(xml).Contains("xr10:uid=\"{A37CCB28-2182-443B-9ED3-90B79AD62CDA}\"");
        await Assert.That(xml).Contains("rowHeight=\"247650\"");
        await Assert.That(xml).Contains("mc:Ignorable");
    }

    [Test]
    public async Task An_edited_slicer_reloads_with_the_new_values()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var slicer = wb.Worksheet("Pivot").Slicers.Single();
            slicer.Style = "SlicerStyleLight2";
            slicer.ShowCaption = false;
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var reloadedSlicer = reloaded.Worksheet("Pivot").Slicers.Single();

        await Assert.That(reloadedSlicer.Style).IsEqualTo("SlicerStyleLight2");
        await Assert.That(reloadedSlicer.ShowCaption).IsFalse();
        await Assert.That(reloadedSlicer.Caption).IsEqualTo("Region filter").Because("An untouched property is untouched.");
    }

    [Test]
    public async Task Editing_one_slicer_leaves_the_other_slicers_part_untouched()
    {
        using var original = Resource();
        var before = PartBytes(original, "xl/slicers/slicer1.xml");

        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Pivot").Slicers.Single().Caption = "Edited";
            wb.SaveAs(saved);
        }

        // slicer1 is the table slicer on the other sheet. Nobody assigned to it, so its part is not
        // even opened — the byte comparison is what proves the patcher's gate actually gates.
        await Assert.That(PartBytes(saved, "xl/slicers/slicer1.xml")).IsEquivalentTo(before, CollectionOrdering.Matching);
    }

    [Test]
    public async Task Setting_a_caption_back_to_the_name_drops_the_attribute()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var slicer = wb.Worksheet("Pivot").Slicers.Single();
            slicer.Caption = slicer.Name;
            wb.SaveAs(saved);
        }

        // Excel omits the caption when it matches the name and shows the name instead, so restating
        // it would be a difference from what Excel writes for the same state.
        await Assert.That(ReadPart(saved, "xl/slicers/slicer2.xml")).DoesNotContain("caption=");
    }

    // ── Adding alongside a slicer that is already there ─────────────────

    /// <summary>
    /// Adding a slicer to a sheet that already carries one must not disturb the slicer that was
    /// already there.
    /// </summary>
    /// <remarks>
    /// This is the guard for the defect the acceptance-criteria check found: the writer used to
    /// reuse the sheet's existing <c>xl/slicers/slicerN.xml</c> and append the new definition into
    /// it, which re-serialises a part Excel wrote. Nothing was lost from the XML and the validator
    /// was clean, but Excel then stopped drawing the original — the same silent class of failure
    /// the round-trip guarantee exists to prevent. The byte comparison is the assertion; the
    /// reload afterwards only shows the sheet still has both.
    /// </remarks>
    [Test]
    public async Task Adding_a_slicer_leaves_the_slicer_already_on_the_sheet_byte_for_byte_intact()
    {
        using var original = Resource();
        var before = PartBytes(original, "xl/slicers/slicer2.xml");

        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = (XLPivotTable)wb.Worksheet("Pivot").PivotTables.Single();
            SlicersOf(wb, "Pivot").AddPivotSlicer(pivotTable, "Region");
            wb.SaveAs(saved);
        }

        await Assert.That(PartBytes(saved, "xl/slicers/slicer2.xml")).IsEquivalentTo(before, CollectionOrdering.Matching)
            .Because("Excel's own slicer part must pass through untouched; the new slicer belongs in a part of its own.");
    }

    [Test]
    public async Task A_slicer_added_beside_an_existing_one_gets_its_own_part_and_list_entry()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = (XLPivotTable)wb.Worksheet("Pivot").PivotTables.Single();
            SlicersOf(wb, "Pivot").AddPivotSlicer(pivotTable, "Region");
            wb.SaveAs(saved);
        }

        // Every slicers part Excel writes holds exactly one slicer, and the sheet's slicerList
        // names one part per slicer. That is the shape the manual check confirmed working.
        foreach (var part in SlicerParts(saved))
        {
            var count = System.Text.RegularExpressions.Regex.Matches(ReadPart(saved, part), "<[^>]*:?slicer ").Count;
            await Assert.That(count).IsEqualTo(1).Because($"{part} should define exactly one slicer.");
        }

        var sheetXml = ReadPart(saved, "xl/worksheets/sheet2.xml");
        var refs = System.Text.RegularExpressions.Regex.Matches(sheetXml, "<x14:slicer r:id=").Count;
        await Assert.That(refs).IsEqualTo(2).Because("The sheet now points at two slicer parts.");
    }

    [Test]
    public async Task Both_slicers_are_there_after_adding_one_beside_another()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = (XLPivotTable)wb.Worksheet("Pivot").PivotTables.Single();
            SlicersOf(wb, "Pivot").AddPivotSlicer(pivotTable, "Region");
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var names = reloaded.Worksheet("Pivot").Slicers.Select(s => s.Name).OrderBy(n => n).ToArray();

        await Assert.That(names).IsEquivalentTo(new[] { "Region 1", "Region 2" });
    }

    [Test]
    public async Task A_workbook_with_a_slicer_added_beside_an_existing_one_is_schema_valid()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = (XLPivotTable)wb.Worksheet("Pivot").PivotTables.Single();
            SlicersOf(wb, "Pivot").AddPivotSlicer(pivotTable, "Region");
            wb.SaveAs(saved);
        }

        await AssertSchemaValid(saved);
    }

    // ── Creating a pivot slicer ─────────────────────────────────────────

    [Test]
    public async Task A_created_pivot_slicer_writes_all_five_of_its_non_anchor_pieces()
    {
        using var saved = new MemoryStream();
        string cacheName;

        using (var wb = Load())
        {
            var pivotTable = (XLPivotTable)wb.Worksheet("Pivot").PivotTables.Single();
            var slicer = SlicersOf(wb, "Pivot").AddPivotSlicer(pivotTable, "Region");
            cacheName = slicer.Cache.Name;
            wb.SaveAs(saved);
        }

        // 1. A slicer definition, in a part of its own or alongside the sheet's existing one.
        var slicerXml = string.Concat(SlicerParts(saved).Select(p => ReadPart(saved, p)));
        await Assert.That(slicerXml).Contains("name=\"Region 2\"")
            .Because("Region and Region 1 are taken, so the next free slicer name is Region 2.");
        await Assert.That(slicerXml).Contains($"cache=\"{cacheName}\"");

        // 2. The worksheet extLst reference, under the pivot slicer URI.
        var sheetXml = ReadPart(saved, "xl/worksheets/sheet2.xml");
        await Assert.That(sheetXml).Contains("{A8765BA9-456A-4dab-B4F3-ACF838C121DE}");

        // 3. A cache part, bound to the pivot table by name.
        var cacheXml = string.Concat(CacheParts(saved).Select(p => ReadPart(saved, p)));
        await Assert.That(cacheXml).Contains($"name=\"{cacheName}\"");
        await Assert.That(cacheXml).Contains("SalesPivot");

        // 4. The workbook registration, in the x14 registry rather than the x15 one.
        var workbookXml = ReadPart(saved, "xl/workbook.xml");
        await Assert.That(workbookXml).Contains("{BBE1A952-AA13-448e-AADC-164F8A28A991}");

        // 5. The #N/A defined name Excel writes per cache.
        // The prefix the writer puts on the element is not the point, so the assertion is on the
        // name and its #N/A value rather than on the serialised form.
        await Assert.That(workbookXml).Contains($"definedName name=\"{cacheName}\">#N/A<");
    }

    [Test]
    public async Task A_pivot_table_a_slicer_filters_is_stamped_with_a_version_that_supports_slicers()
    {
        // The bug this pins cost six rounds of manual Excel checking to find, so it is worth
        // stating precisely. XLibur left createdVersion and updatedVersion at 0 on a pivot table it
        // created, and the writer omits an attribute at its default, so they were absent altogether.
        // Excel reads a pivot table stamped version 0 as predating slicers and silently refuses to
        // draw one bound to it — no repair prompt, no validation error, no missing part. Every
        // automated gate passed; the slicer just was not there.
        //
        // Excel-authored files carry 8 here; XLibur writes its own baseline of 5. Anything at or
        // above 4, the version slicers arrived in, will do. Zero will not.
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var data = wb.AddWorksheet("Data");
            data.Cell("A1").Value = "Region";
            data.Cell("B1").Value = "Amount";
            data.Cell("A2").Value = "North";
            data.Cell("B2").Value = 10;
            data.Cell("A3").Value = "South";
            data.Cell("B3").Value = 20;

            var sheet = wb.AddWorksheet("Pivot");
            var pivot = sheet.PivotTables.Add("P", sheet.Cell("A3"), data.Range("A1:B3"));
            pivot.RowLabels.Add("Region");
            pivot.Values.Add("Amount");
            sheet.Slicers.Add(pivot, "Region");

            wb.SaveAs(saved);
        }

        var xml = ReadPart(saved, "xl/pivotTables/pivotTable.xml");

        await Assert.That(xml).Contains("createdVersion=\"5\"")
            .Because("A pivot table stamped version 0 is one no slicer can attach to.");
        await Assert.That(xml).Contains("updatedVersion=\"5\"");
        await Assert.That(xml).Contains("minRefreshableVersion=\"3\"");
    }

    [Test]
    public async Task A_loaded_pivot_table_keeps_the_version_its_file_declares()
    {
        // The default above must not overwrite what a file says: an Excel-authored pivot table
        // carries 8, and saving it back as 5 would be XLibur claiming authorship of someone else's
        // pivot table and downgrading it on the way through.
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Data").Cell("H1").Value = "touched";
            wb.SaveAs(saved);
        }

        var xml = ReadPart(saved, "xl/pivotTables/pivotTable1.xml");

        await Assert.That(xml).Contains("createdVersion=\"8\"");
        await Assert.That(xml).Contains("updatedVersion=\"8\"");
    }

    [Test]
    public async Task A_created_pivot_slicer_binds_to_the_pivot_cache_identifier()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = (XLPivotTable)wb.Worksheet("Pivot").PivotTables.Single();
            SlicersOf(wb, "Pivot").AddPivotSlicer(pivotTable, "Region");
            wb.SaveAs(saved);
        }

        // The slicer cache quotes the pivot cache's own identifier, which lives in an extension of
        // the pivot cache definition and is not the renumbered cacheId in workbook.xml. The fixture
        // already carries one, so a slicer added to it has to reuse that rather than invent a
        // second — otherwise the new slicer points at a pivot cache that does not exist.
        var pivotCacheXml = ReadPart(saved, "xl/pivotCache/pivotCacheDefinition1.xml");
        await Assert.That(pivotCacheXml).Contains("pivotCacheId=\"973837003\"");

        var cacheXml = string.Concat(CacheParts(saved).Select(p => ReadPart(saved, p)));
        await Assert.That(cacheXml).Contains("pivotCacheId=\"973837003\"");
    }

    [Test]
    public async Task A_created_pivot_slicer_starts_with_every_item_selected()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = (XLPivotTable)wb.Worksheet("Pivot").PivotTables.Single();
            var slicer = SlicersOf(wb, "Pivot").AddPivotSlicer(pivotTable, "Region");

            // A slicer nobody has clicked filters nothing, and in the file that is every item
            // marked selected rather than an absent item list.
            await Assert.That(slicer.SelectedItems.Count).IsEqualTo(4);
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var created = reloaded.Worksheet("Pivot").Slicers.Single(s => s.Name == "Region 2");

        // The round trip has to agree with the model it came from, which is why the items are
        // populated when the slicer is created rather than when it is written.
        await Assert.That(created.SelectedItems.Select(i => i.GetText()).OrderBy(t => t))
            .IsEquivalentTo(new[] { "East", "North", "South", "West" });
    }

    // ── Creating a table slicer ─────────────────────────────────────────

    [Test]
    public async Task A_created_table_slicer_registers_under_the_table_uris()
    {
        using var saved = new MemoryStream();
        string cacheName;

        using (var wb = Load())
        {
            var table = (XLTable)wb.Worksheet("Data").Tables.Single();
            var slicer = SlicersOf(wb, "Data").AddTableSlicer(table, "Amount");
            cacheName = slicer.Cache.Name;
            wb.SaveAs(saved);
        }

        // A table slicer uses the other of each pair of URIs. Getting either wrong orphans the
        // cache as surely as leaving it out.
        await Assert.That(ReadPart(saved, "xl/worksheets/sheet1.xml"))
            .Contains("{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}");

        var workbookXml = ReadPart(saved, "xl/workbook.xml");
        await Assert.That(workbookXml).Contains("{46BE6895-7355-4a93-B00E-2C351335B9C9}");
        // The prefix the writer puts on the element is not the point, so the assertion is on the
        // name and its #N/A value rather than on the serialised form.
        await Assert.That(workbookXml).Contains($"definedName name=\"{cacheName}\">#N/A<");

        // The cache binds by table id and column id, neither of which XLibur models — both come
        // from what the table part was actually written as. Amount is the third column.
        var cacheXml = string.Concat(CacheParts(saved).Select(p => ReadPart(saved, p)));
        await Assert.That(cacheXml).Contains("tableSlicerCache");
        await Assert.That(cacheXml).Contains("column=\"3\"");
    }

    [Test]
    public async Task A_created_table_slicer_reloads_bound_to_its_column()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var table = (XLTable)wb.Worksheet("Data").Tables.Single();
            SlicersOf(wb, "Data").AddTableSlicer(table, "Amount");
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var created = reloaded.Worksheet("Data").Slicers.Single(s => s.SourceFieldName == "Amount");

        await Assert.That(created.SourceKind).IsEqualTo(XLSlicerSourceKind.Table);
        await Assert.That(created.Table?.Name).IsEqualTo("SalesTable");
        await Assert.That(created.HasSelection).IsFalse().Because("A fresh table slicer filters nothing.");
    }

    // ── The cascade ─────────────────────────────────────────────────────

    [Test]
    public async Task Deleting_a_pivot_table_takes_its_slicer_with_it()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            await Assert.That(wb.Worksheet("Pivot").Slicers.Count).IsEqualTo(1);

            wb.Worksheet("Pivot").PivotTables.Delete("SalesPivot");

            // The slicer had nothing else to filter, so it goes with the pivot table rather than
            // being left pointing at something that is no longer there.
            await Assert.That(wb.Worksheet("Pivot").Slicers.Count).IsEqualTo(0);
            wb.SaveAs(saved);
        }

        // Every trace has to go, or the saved file has an orphan Excel offers to repair.
        await Assert.That(PartExists(saved, "xl/slicers/slicer2.xml")).IsFalse();
        await Assert.That(PartExists(saved, "xl/slicerCaches/slicerCache1.xml")).IsFalse();

        var workbookXml = ReadPart(saved, "xl/workbook.xml");
        await Assert.That(workbookXml).DoesNotContain("Slicer_Region1");
        await Assert.That(ReadPart(saved, "xl/worksheets/sheet2.xml")).DoesNotContain("slicerList");
    }

    /// <summary>
    /// A slicer taken by the cascade takes its drawing with it.
    /// </summary>
    /// <remarks>
    /// The frame is what Excel draws a slicer through, and it names the slicer it belongs to. Left
    /// behind, it names one that no longer exists anywhere in the package — the sixth piece of a
    /// slicer outliving the other five. The definition, the part, the cache, the registration, the
    /// defined name and the <c>slicerList</c> entry were all being unpicked; the drawing was not.
    /// </remarks>
    [Test]
    public async Task Deleting_a_pivot_table_takes_its_slicers_drawing_with_it()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Pivot").PivotTables.Delete("SalesPivot");
            wb.SaveAs(saved);
        }

        var drawingXml = ReadPart(saved, "xl/drawings/drawing2.xml");
        await Assert.That(drawingXml).DoesNotContain("Region 1")
            .Because("The graphic frame names the slicer, so it cannot outlive it.");
        await Assert.That(drawingXml).DoesNotContain("/slicer")
            .Because("Nothing on this sheet is drawn through a slicer frame any more.");
    }

    [Test]
    public async Task Deleting_a_pivot_table_leaves_the_other_sheets_slicer_drawing_alone()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Pivot").PivotTables.Delete("SalesPivot");
            wb.SaveAs(saved);
        }

        // The table slicer on the other sheet is untouched by the cascade, so its frame stays.
        await Assert.That(ReadPart(saved, "xl/drawings/drawing1.xml")).Contains("Region");
    }

    [Test]
    public async Task Deleting_a_pivot_table_leaves_the_table_slicer_alone()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Pivot").PivotTables.Delete("SalesPivot");
            wb.SaveAs(saved);
        }

        // The table slicer on the other sheet has nothing to do with the pivot table, and both
        // slicers filter a field called Region — so a cascade keyed on anything looser than the
        // cache's own pivot table list would take this one out too.
        await Assert.That(PartExists(saved, "xl/slicers/slicer1.xml")).IsTrue();
        await Assert.That(PartExists(saved, "xl/slicerCaches/slicerCache2.xml")).IsTrue();
        await Assert.That(ReadPart(saved, "xl/workbook.xml")).Contains("name=\"Slicer_Region\">");
        await Assert.That(ReadPart(saved, "xl/worksheets/sheet1.xml")).Contains("slicerList");
    }

    [Test]
    public async Task A_slicer_shared_by_two_pivot_tables_survives_losing_one()
    {
        using var wb = Load();

        var pivotSheet = wb.Worksheet("Pivot");
        var first = (XLPivotTable)pivotSheet.PivotTables.Single();
        var slicer = pivotSheet.Slicers.Single();

        // A dashboard slicer drives several pivot tables through one cache. Losing one of them is
        // not the same as losing the slicer.
        var second = (XLPivotTable)wb.Worksheet("Data").PivotTables
            .Add("SecondPivot", wb.Worksheet("Data").Cell("H1"), first.PivotCache);
        ((XLSlicer)slicer).Cache.PivotTables.Add(second);
        ((XLSlicer)slicer).Cache.PivotTableNames.Add("SecondPivot");

        wb.Worksheet("Pivot").PivotTables.Delete("SalesPivot");

        await Assert.That(pivotSheet.Slicers.Count).IsEqualTo(1);
        await Assert.That(slicer.PivotTables.Select(pt => pt.Name)).IsEquivalentTo(new[] { "SecondPivot" });
    }

    // ── Schema ──────────────────────────────────────────────────────────

    [Test]
    public async Task A_workbook_with_a_created_pivot_slicer_is_schema_valid()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = (XLPivotTable)wb.Worksheet("Pivot").PivotTables.Single();
            SlicersOf(wb, "Pivot").AddPivotSlicer(pivotTable, "Region");
            wb.SaveAs(saved);
        }

        await AssertSchemaValid(saved);
    }

    [Test]
    public async Task A_workbook_with_a_created_table_slicer_is_schema_valid()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var table = (XLTable)wb.Worksheet("Data").Tables.Single();
            SlicersOf(wb, "Data").AddTableSlicer(table, "Amount");
            wb.SaveAs(saved);
        }

        await AssertSchemaValid(saved);
    }

    [Test]
    public async Task A_workbook_that_lost_its_slicer_to_the_cascade_is_schema_valid()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Pivot").PivotTables.Delete("SalesPivot");
            wb.SaveAs(saved);
        }

        await AssertSchemaValid(saved);
    }

    [Test]
    public async Task A_workbook_with_an_edited_slicer_is_schema_valid()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var slicer = wb.Worksheet("Pivot").Slicers.Single();
            slicer.Caption = "Pick a region";
            slicer.ColumnCount = 3;
            slicer.RowHeightPt = 22;
            wb.SaveAs(saved);
        }

        await AssertSchemaValid(saved);
    }

    #region Helpers

    /// <summary>
    /// The worksheet's own slicer collection. Creation is internal until anchoring lands, so this
    /// is how the tests reach it.
    /// </summary>
    private static XLSlicers SlicersOf(XLWorkbook wb, string sheetName) =>
        ((XLWorksheet)wb.Worksheet(sheetName)).SlicersInternal;

    /// <summary>
    /// The fixture, opened over a copy that outlives this call.
    /// </summary>
    /// <remarks>
    /// The workbook keeps hold of the stream it was opened from and reads it again on save — that
    /// is the mechanism the whole round trip depends on — so the stream cannot be disposed when
    /// this returns. A <see cref="MemoryStream"/> needs no deterministic disposal, so handing one
    /// over and letting it go is safe.
    /// </remarks>
    private static XLWorkbook Load()
    {
        var stream = Resource();
        stream.Position = 0;
        return new XLWorkbook(stream);
    }

    private static MemoryStream Resource()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Fixture));
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms;
    }

    private static async Task AssertSchemaValid(MemoryStream package)
    {
        package.Position = 0;
        using var doc = SpreadsheetDocument.Open(package, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2010)
            .Validate(doc)
            .Select(error => $"{error.Path?.XPath}: {error.Description}")
            .ToList();

        // Joined rather than asserted as a collection so a failure names what is wrong. Excel
        // repairs a file it cannot parse instead of saying where it broke, which makes the
        // validator the only thing that will tell you.
        await Assert.That(string.Join(Environment.NewLine, errors)).IsEmpty();
    }

    private static string[] SlicerParts(MemoryStream package) => PartsUnder(package, "xl/slicers/");

    private static string[] CacheParts(MemoryStream package) => PartsUnder(package, "xl/slicerCaches/");

    private static string[] PartsUnder(MemoryStream package, string prefix)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        return archive.Entries
            .Where(e => e.FullName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .Select(e => e.FullName)
            .ToArray();
    }

    private static bool PartExists(MemoryStream package, string partPath)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        return archive.Entries.Any(e => e.FullName.Equals(partPath, StringComparison.OrdinalIgnoreCase));
    }

    private static string ReadPart(MemoryStream package, string partPath)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.First(e => e.FullName.Equals(partPath, StringComparison.OrdinalIgnoreCase));

        using var entryStream = entry.Open();
        using var reader = new StreamReader(entryStream);
        return reader.ReadToEnd();
    }

    private static byte[] PartBytes(MemoryStream package, string partPath)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.First(e => e.FullName.Equals(partPath, StringComparison.OrdinalIgnoreCase));

        using var entryStream = entry.Open();
        using var buffer = new MemoryStream();
        entryStream.CopyTo(buffer);
        return buffer.ToArray();
    }

    #endregion
}
