using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Slicers;

/// <summary>
/// The slicer read model, read against the Excel-authored fixture that settled the fidelity
/// question in PRD 5 task 1.
/// </summary>
/// <remarks>
/// <para>
/// <c>Resource/TryToLoad/SlicersOnPivotAndTable.xlsx</c> carries one slicer of each kind, which is
/// what makes it worth reading twice. Sheet <c>Data</c> holds a table, <c>SalesTable</c>, with a
/// table slicer bound to its <c>Region</c> column. Sheet <c>Pivot</c> holds <c>SalesPivot</c> with
/// a pivot slicer bound to the same field, renamed to "Region filter", styled
/// <c>SlicerStyleDark3</c> and filtered to a single item.
/// </para>
/// <para>
/// The parts are cross-numbered — <c>slicerCache1</c> serves the pivot slicer and
/// <c>slicerCache2</c> the table slicer — so a test that happens to pass by index rather than by
/// binding would show up here.
/// </para>
/// </remarks>
public class SlicerReadModelTests
{
    private const string Fixture = @"TryToLoad\SlicersOnPivotAndTable.xlsx";

    [Test]
    public async Task Each_worksheet_owns_the_slicers_drawn_on_it()
    {
        using var wb = Load();

        await Assert.That(wb.Worksheet("Data").Slicers.Count).IsEqualTo(1);
        await Assert.That(wb.Worksheet("Pivot").Slicers.Count).IsEqualTo(1);
    }

    [Test]
    public async Task A_pivot_slicer_reports_the_styling_XLibur_does_not_otherwise_model()
    {
        using var wb = Load();

        var slicer = wb.Worksheet("Pivot").Slicers.Single();

        // The caption is the user-visible heading and is not the name; the style is a built-in
        // Excel style XLibur has no model for beyond its name. Both are what a template author
        // opening someone else's workbook needs to see.
        await Assert.That(slicer.Name).IsEqualTo("Region 1");
        await Assert.That(slicer.Caption).IsEqualTo("Region filter");
        await Assert.That(slicer.Style).IsEqualTo("SlicerStyleDark3");
        await Assert.That(slicer.ShowCaption).IsTrue();
        await Assert.That(slicer.ColumnCount).IsEqualTo(1u);

        // rowHeight is written in EMU: 247650 EMU is 19.5 pt.
        await Assert.That(slicer.RowHeightPt).IsEqualTo(19.5);
    }

    [Test]
    public async Task A_slicer_with_no_caption_of_its_own_reports_its_name()
    {
        using var wb = Load();

        // Excel omits the caption attribute when it matches the name, and shows the name.
        var slicer = wb.Worksheet("Data").Slicers.Single();

        await Assert.That(slicer.Name).IsEqualTo("Region");
        await Assert.That(slicer.Caption).IsEqualTo("Region");
        await Assert.That(slicer.Style).IsNull().Because("The table slicer carries no style attribute.");
    }

    [Test]
    public async Task A_slicer_can_be_found_by_name()
    {
        using var wb = Load();

        var slicers = wb.Worksheet("Pivot").Slicers;

        await Assert.That(slicers.Slicer("Region 1").Caption).IsEqualTo("Region filter");
        await Assert.That(slicers.TryGetSlicer("Region 1", out var found)).IsTrue();
        await Assert.That(found!.Name).IsEqualTo("Region 1");

        // The name is not the caption, and looking one up by the other finds nothing.
        await Assert.That(slicers.TryGetSlicer("Region filter", out _)).IsFalse();
    }

    // ── Binding ─────────────────────────────────────────────────────────

    [Test]
    public async Task A_pivot_slicer_binds_to_the_pivot_table_its_cache_names()
    {
        using var wb = Load();

        var slicer = wb.Worksheet("Pivot").Slicers.Single();

        await Assert.That(slicer.SourceKind).IsEqualTo(XLSlicerSourceKind.PivotTable);
        await Assert.That(slicer.SourceFieldName).IsEqualTo("Region");
        await Assert.That(slicer.PivotTables.Select(pt => pt.Name)).IsEquivalentTo(new[] { "SalesPivot" });
        await Assert.That(slicer.Table).IsNull();
    }

    [Test]
    public async Task A_table_slicer_binds_to_the_table_column_its_cache_names()
    {
        using var wb = Load();

        var slicer = wb.Worksheet("Data").Slicers.Single();

        // The other of the two binding paths: an x15:tableSlicerCache extension naming a table id
        // and a column id, neither of which XLibur models, resolved back to the loaded table.
        await Assert.That(slicer.SourceKind).IsEqualTo(XLSlicerSourceKind.Table);
        await Assert.That(slicer.SourceFieldName).IsEqualTo("Region");
        await Assert.That(slicer.Table?.Name).IsEqualTo("SalesTable");
        await Assert.That(slicer.PivotTables).IsEmpty();
    }

    [Test]
    public async Task A_pivot_table_sees_the_slicers_that_filter_it()
    {
        using var wb = Load();

        var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();

        // The view crosses sheets by construction, which is the point of it being a view: the
        // slicer is owned by the sheet it is drawn on, and only the cache says what it filters.
        await Assert.That(pivotTable.Slicers.Select(s => s.Name)).IsEquivalentTo(new[] { "Region 1" });
        await Assert.That(pivotTable.Slicers.Single()).IsSameReferenceAs(wb.Worksheet("Pivot").Slicers.Single());
    }

    [Test]
    public async Task The_table_slicer_does_not_show_up_on_the_pivot_table()
    {
        using var wb = Load();

        // Both slicers filter a field called Region, and the two caches are cross-numbered against
        // the two slicer parts. Binding by name rather than by position is what keeps them apart.
        var tableSlicer = wb.Worksheet("Data").Slicers.Single();
        await Assert.That(wb.Worksheet("Pivot").PivotTables.Single().Slicers).DoesNotContain(tableSlicer);
    }

    // ── Selection ───────────────────────────────────────────────────────

    [Test]
    public async Task A_pivot_slicers_selection_is_read_from_its_cache()
    {
        using var wb = Load();

        var slicer = wb.Worksheet("Pivot").Slicers.Single();

        await Assert.That(slicer.HasSelection).IsTrue();
        await Assert.That(slicer.SelectedItems.Select(i => i.GetText())).IsEquivalentTo(new[] { "East" });
    }

    [Test]
    public async Task The_selected_items_are_the_ones_the_pivot_table_is_not_hiding()
    {
        using var wb = Load();

        // The cache marks a selected item with s="1" and writes nothing on the others, which is the
        // opposite of what the attribute's name suggests and is not stated by the schema. The pivot
        // field is the independent witness, read by a different reader out of a different part: the
        // item the slicer selects has to be the one the pivot table is showing, and the three it
        // does not select have to be the three the pivot table marks hidden. Reading the flag the
        // other way round would give three items here instead of one.
        var slicer = wb.Worksheet("Pivot").Slicers.Single();
        var pivotTable = (XLPivotTable)wb.Worksheet("Pivot").PivotTables.Single();

        var shown = pivotTable.PivotFields[0].Items
            .Where(item => !item.Hidden && item.ItemType == XLPivotItemType.Data)
            .Select(item => item.GetValue()?.GetText())
            .ToList();

        await Assert.That(shown.Count).IsEqualTo(1)
            .Because("The fixture's pivot field hides three of its four items.");
        await Assert.That(slicer.SelectedItems.Select(i => i.GetText())).IsEquivalentTo(shown);
    }

    [Test]
    public async Task An_unfiltered_table_slicer_reports_no_selection()
    {
        using var wb = Load();

        // Nobody clicked this one, so the table's auto filter has no filter column for it and every
        // item is showing.
        var slicer = wb.Worksheet("Data").Slicers.Single();

        await Assert.That(slicer.HasSelection).IsFalse();
        await Assert.That(slicer.SelectedItems).IsEmpty();
    }

    // ── Fidelity ────────────────────────────────────────────────────────

    [Test]
    public async Task Reading_a_slicer_leaves_its_part_byte_for_byte_identical()
    {
        // This is the regression gate for the whole read model. Slicer parts survive a round trip
        // because nothing opens them; reaching one through SlicersPart.Slicers would attach a DOM
        // that the SDK writes back over the original bytes on save, taking with it every attribute
        // XLibur does not model. The reader streams the parts detached instead, and this asserts it.
        using var original = Resource();
        var before = PartBytes(original, "xl/slicers/slicer2.xml");
        var beforeCache = PartBytes(original, "xl/slicerCaches/slicerCache1.xml");

        using var saved = LoadAndSave();

        await Assert.That(PartBytes(saved, "xl/slicers/slicer2.xml")).IsEquivalentTo(before);
        await Assert.That(PartBytes(saved, "xl/slicerCaches/slicerCache1.xml")).IsEquivalentTo(beforeCache);
    }

    [Test]
    public async Task Slicers_still_load_after_a_round_trip()
    {
        // Reading the model back out of the saved package proves the relationships and content types
        // came through too, not just the part bytes.
        using var saved = LoadAndSave();
        saved.Position = 0;
        using var wb = new XLWorkbook(saved);

        await Assert.That(wb.Worksheet("Data").Slicers.Count).IsEqualTo(1);
        await Assert.That(wb.Worksheet("Pivot").Slicers.Single().Style).IsEqualTo("SlicerStyleDark3");
    }

    [Test]
    public async Task The_round_tripped_package_is_schema_valid()
    {
        using var saved = LoadAndSave();
        saved.Position = 0;

        using var doc = SpreadsheetDocument.Open(saved, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2010)
            .Validate(doc)
            .Select(error => $"{error.Path?.XPath}: {error.Description}")
            .ToList();

        await Assert.That(errors).IsEmpty();
    }

    #region Helpers

    /// <summary>
    /// The fixture, opened over a copy that outlives this call. The workbook reads its original
    /// stream again on save, so the stream cannot be disposed when this returns.
    /// </summary>
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

    private static MemoryStream LoadAndSave()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Fixture));
        var ms = new MemoryStream();

        using (var wb = new XLWorkbook(stream))
            wb.SaveAs(ms);

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

    #endregion
}
