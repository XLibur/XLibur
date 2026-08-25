using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XLibur.Excel;
using XLibur.Excel.Tables;

namespace XLibur.Tests.Excel.Slicers;

/// <summary>
/// Where a slicer sits: reading the anchor a file carries, placing a created one, and moving either.
/// </summary>
/// <remarks>
/// A slicer is drawn by a <c>xdr:graphicFrame</c> in the sheet's drawing part, not by its own part,
/// so the anchor is the sixth of the six pieces a created slicer needs. It is built by
/// <c>DrawingAnchorFactory</c> and by nothing else here.
/// </remarks>
public class SlicerPositionTests
{
    private const string Fixture = @"TryToLoad\SlicersOnPivotAndTable.xlsx";

    // ── Reading ─────────────────────────────────────────────────────────

    [Test]
    public async Task A_loaded_slicer_reports_where_its_anchor_puts_it()
    {
        using var wb = Load();

        // Both fixture slicers hang off a two-cell anchor whose from marker is col 5, row 1 —
        // zero-based in the file, so F2 in the model.
        await Assert.That(wb.Worksheet("Pivot").Slicers.Single().Position.Address.ToString()).IsEqualTo("F2");
        await Assert.That(wb.Worksheet("Data").Slicers.Single().Position.Address.ToString()).IsEqualTo("F2");
    }

    [Test]
    public async Task Reading_a_position_does_not_count_as_an_edit()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            // Reading every property has to leave the slicer unassigned, or the patcher would open
            // parts nobody touched. The byte comparison below is what proves it.
            var slicer = wb.Worksheet("Pivot").Slicers.Single();
            _ = slicer.Position;
            _ = slicer.Caption;
            _ = slicer.Style;
            wb.SaveAs(saved);
        }

        using var original = Resource();
        await Assert.That(PartBytes(saved, "xl/slicers/slicer2.xml"))
            .IsEquivalentTo(PartBytes(original, "xl/slicers/slicer2.xml"));
    }

    // ── Placing a created slicer ────────────────────────────────────────

    [Test]
    public async Task A_created_pivot_slicer_is_placed_beside_its_pivot_table_not_at_A1()
    {
        using var wb = Load();

        var pivotTable = (XLPivotTable)wb.Worksheet("Pivot").PivotTables.Single();
        var slicer = SlicersOf(wb, "Pivot").AddPivotSlicer(pivotTable, "Region");

        // DrawingAnchorFactory silently anchors a marker-less drawing at A1. For a slicer that
        // would drop the panel on top of the data it filters, so XLibur always supplies a marker
        // and that fallback stays unreachable. The pivot table occupies A3:B5.
        await Assert.That(slicer.Position.Address.ToString()).IsNotEqualTo("A1");
        await Assert.That(slicer.Position.Address.ColumnNumber).IsEqualTo(4);
        await Assert.That(slicer.Position.Address.RowNumber).IsEqualTo(3);
    }

    [Test]
    public async Task A_created_table_slicer_is_placed_beside_its_table()
    {
        using var wb = Load();

        var table = (XLTable)wb.Worksheet("Data").Tables.Single();
        var slicer = SlicersOf(wb, "Data").AddTableSlicer(table, "Amount");

        // SalesTable occupies A1:C15, so two columns of clearance puts the slicer at E1.
        await Assert.That(slicer.Position.Address.ToString()).IsEqualTo("E1");
    }

    [Test]
    public async Task A_created_slicer_writes_a_graphic_frame_into_the_drawing()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = (XLPivotTable)wb.Worksheet("Pivot").PivotTables.Single();
            SlicersOf(wb, "Pivot").AddPivotSlicer(pivotTable, "Region");
            wb.SaveAs(saved);
        }

        // The frame names the slicer and sits under the slicer graphic-data URI; that pair is what
        // Excel resolves to draw the panel.
        var drawing = ReadPart(saved, "xl/drawings/drawing2.xml");
        await Assert.That(drawing).Contains("http://schemas.microsoft.com/office/drawing/2010/slicer");
        await Assert.That(drawing).Contains("name=\"Region 2\"");
        await Assert.That(drawing).Contains("oneCellAnchor");
    }

    [Test]
    public async Task A_created_slicer_on_a_sheet_with_no_drawing_gets_one()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            // A third sheet has no drawing part at all, so the slicer has to bring one with it —
            // and the sheet has to end up referencing it, or the part is an orphan.
            var sheet = wb.AddWorksheet("Extra");
            var pivotTable = (XLPivotTable)wb.Worksheet("Pivot").PivotTables.Single();
            ((XLWorksheet)sheet).SlicersInternal.AddPivotSlicer(pivotTable, "Region");
            wb.SaveAs(saved);
        }

        await Assert.That(PartExists(saved, "xl/drawings/drawing3.xml")).IsTrue();
        await Assert.That(ReadPart(saved, "xl/worksheets/sheet3.xml")).Contains("<x:drawing");
    }

    // ── Moving ──────────────────────────────────────────────────────────

    [Test]
    public async Task Moving_a_loaded_slicer_shifts_both_corners_and_keeps_its_size()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            // F2 to H4 is two columns right and two rows down.
            wb.Worksheet("Pivot").Slicers.Single().Position = wb.Worksheet("Pivot").Cell("H4");
            wb.SaveAs(saved);
        }

        var drawing = ReadPart(saved, "xl/drawings/drawing2.xml");

        // From was col 5 row 1; to was col 8 row 15. Both move by the same delta, so the panel
        // covers the same number of columns and rows as before.
        await Assert.That(drawing).Contains("<xdr:col>7</xdr:col>");
        await Assert.That(drawing).Contains("<xdr:row>3</xdr:row>");
        await Assert.That(drawing).Contains("<xdr:col>10</xdr:col>");
        await Assert.That(drawing).Contains("<xdr:row>17</xdr:row>");
    }

    [Test]
    public async Task Moving_a_slicer_keeps_the_rest_of_its_frame()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Pivot").Slicers.Single().Position = wb.Worksheet("Pivot").Cell("H4");
            wb.SaveAs(saved);
        }

        var drawing = ReadPart(saved, "xl/drawings/drawing2.xml");

        // The anchor is edited, not replaced. Excel's frame carries an mc:AlternateContent wrapper,
        // a fallback shape and a creationId, none of which XLibur models — replacing the anchor to
        // move a slicer would throw all of it away.
        await Assert.That(drawing).Contains("AlternateContent");
        await Assert.That(drawing).Contains("creationId");
        await Assert.That(drawing).Contains("editAs=\"oneCell\"");
        await Assert.That(drawing).Contains("This shape represents a slicer");
    }

    [Test]
    public async Task Moving_a_slicer_does_not_touch_its_slicer_part()
    {
        using var original = Resource();
        var before = PartBytes(original, "xl/slicers/slicer2.xml");

        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Pivot").Slicers.Single().Position = wb.Worksheet("Pivot").Cell("H4");
            wb.SaveAs(saved);
        }

        // Position lives in the drawing, so moving a slicer must leave the slicers part closed —
        // the two halves of an edit are gated separately.
        await Assert.That(PartBytes(saved, "xl/slicers/slicer2.xml")).IsEquivalentTo(before);
    }

    [Test]
    public async Task A_moved_slicer_reloads_at_its_new_cell()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Data").Slicers.Single().Position = wb.Worksheet("Data").Cell("J7");
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);

        await Assert.That(reloaded.Worksheet("Data").Slicers.Single().Position.Address.ToString()).IsEqualTo("J7");
    }

    [Test]
    public async Task A_created_slicer_reloads_at_the_cell_it_was_placed_at()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var table = (XLTable)wb.Worksheet("Data").Tables.Single();
            var slicer = SlicersOf(wb, "Data").AddTableSlicer(table, "Amount");
            slicer.Position = wb.Worksheet("Data").Cell("H3");
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var created = reloaded.Worksheet("Data").Slicers.Single(s => s.SourceFieldName == "Amount");

        await Assert.That(created.Position.Address.ToString()).IsEqualTo("H3");
    }

    // ── The public creation API ─────────────────────────────────────────

    [Test]
    public async Task Slicers_can_be_created_through_the_public_interface()
    {
        using var wb = Load();

        IXLSlicers slicers = wb.Worksheet("Pivot").Slicers;
        var pivotSlicer = slicers.Add(wb.Worksheet("Pivot").PivotTables.Single(), "Region");

        IXLSlicers tableSlicers = wb.Worksheet("Data").Slicers;
        var tableSlicer = tableSlicers.Add(wb.Worksheet("Data").Tables.Single(), "Amount");

        await Assert.That(pivotSlicer.SourceKind).IsEqualTo(XLSlicerSourceKind.PivotTable);
        await Assert.That(tableSlicer.SourceKind).IsEqualTo(XLSlicerSourceKind.Table);
        await Assert.That(wb.Worksheet("Pivot").Slicers.Count).IsEqualTo(2);
    }

    [Test]
    public async Task Adding_a_slicer_on_a_field_the_cache_does_not_have_is_refused()
    {
        using var wb = Load();
        var slicers = wb.Worksheet("Pivot").Slicers;
        var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();

        await Assert.That(() => slicers.Add(pivotTable, "NoSuchField")).Throws<ArgumentException>();
    }

    // ── Schema ──────────────────────────────────────────────────────────

    [Test]
    public async Task A_created_slicers_drawing_is_schema_valid()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = (XLPivotTable)wb.Worksheet("Pivot").PivotTables.Single();
            SlicersOf(wb, "Pivot").AddPivotSlicer(pivotTable, "Region");

            var table = (XLTable)wb.Worksheet("Data").Tables.Single();
            SlicersOf(wb, "Data").AddTableSlicer(table, "Amount");

            wb.SaveAs(saved);
        }

        await AssertSchemaValid(saved);
    }

    [Test]
    public async Task A_moved_slicers_drawing_is_schema_valid()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Pivot").Slicers.Single().Position = wb.Worksheet("Pivot").Cell("H4");
            wb.SaveAs(saved);
        }

        await AssertSchemaValid(saved);
    }

    #region Helpers

    private static XLSlicers SlicersOf(XLWorkbook wb, string sheetName) =>
        ((XLWorksheet)wb.Worksheet(sheetName)).SlicersInternal;

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

        await Assert.That(string.Join(Environment.NewLine, errors)).IsEmpty();
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
