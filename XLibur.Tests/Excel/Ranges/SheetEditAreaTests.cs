using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// The <c>Area</c> a sheet listener is handed for an insert is the edited range's <b>leading edge</b>
/// extended by <c>Shift - 1</c> lines, which is precisely the area <c>XLRangeInsertHelper</c> shifts
/// the cells by. Its extent on the shift axis is therefore <c>Shift</c>, whatever the edited range's
/// own extent, so a listener that moves something by the area's extent moves it with its own cell.
/// <para>
/// It used to be the <em>whole</em> range extended by <c>Shift - 1</c>, which agreed with the cells
/// only for a range one line tall on the shift axis — a hyperlink on <c>A1:A5</c> travelled seven
/// rows while its cell travelled three. That was D15 in <c>DEFECTS.md</c>; spec 33 pinned it rather
/// than fixing it, and these tests are the pins, re-pointed.
/// </para>
/// </summary>
public class SheetEditAreaTests
{
    /// <summary>
    /// The area's height is read off a hyperlink placed well below the edit: <c>XLHyperlinks</c>
    /// shifts a fully-covered area by <c>insertedArea.Height</c>, so how far the hyperlink travels
    /// <em>is</em> the listener's area height. Every case now recovers the shift, including the
    /// multi-row ranges that used to travel too far.
    /// </summary>
    [Test]
    [Arguments("A1:A1", 3, 3)]
    [Arguments("A1:A5", 3, 3)]
    [Arguments("A1:A3", 2, 2)]
    public async Task The_listener_area_for_a_row_insert_is_the_leading_edge_extended_by_shift_minus_one(
        string rangeAddress, int shift, int expectedAreaHeight)
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        ws.Cell("A20")!.SetValue("x").SetHyperlink(new XLHyperlink("https://example.invalid/"));

        ws.Range(rangeAddress)!.InsertRowsAbove(shift);

        await Assert.That(ws.Cell(20 + expectedAreaHeight, 1)!.HasHyperlink).IsTrue();
    }

    /// <summary>The column mirror: the area's width is read off a hyperlink well to the right.</summary>
    [Test]
    [Arguments("A1:A1", 3, 3)]
    [Arguments("A1:E1", 3, 3)]
    [Arguments("A1:C1", 2, 2)]
    public async Task The_listener_area_for_a_column_insert_is_the_leading_edge_extended_by_shift_minus_one(
        string rangeAddress, int shift, int expectedAreaWidth)
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        ws.Cell("T1")!.SetValue("x").SetHyperlink(new XLHyperlink("https://example.invalid/"));

        ws.Range(rangeAddress)!.InsertColumnsBefore(shift);

        await Assert.That(ws.Cell(1, 20 + expectedAreaWidth)!.HasHyperlink).IsTrue();
    }

    /// <summary>
    /// The consequence, in the terms D15 was reported in: the cells move by <c>Shift</c> and the
    /// listeners move by the area's height, so the two agreeing is what keeps a hyperlink on a range
    /// taller than one row with its own cell.
    /// </summary>
    [Test]
    public async Task A_multi_row_edited_range_keeps_a_hyperlink_with_its_cell()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        ws.Cell("A5")!.SetValue("x").SetHyperlink(new XLHyperlink("https://example.invalid/"));

        ws.Range("A1:A5")!.InsertRowsAbove(3);

        await Assert.That(ws.Cell("A8")!.GetString()).IsEqualTo("x");
        await Assert.That(ws.Cell("A8")!.HasHyperlink).IsTrue();
        await Assert.That(ws.Cell("A12")!.HasHyperlink).IsFalse();
    }

    /// <summary>The column mirror of <see cref="A_multi_row_edited_range_keeps_a_hyperlink_with_its_cell"/>.</summary>
    [Test]
    public async Task A_multi_column_edited_range_keeps_a_hyperlink_with_its_cell()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        ws.Cell("E1")!.SetValue("x").SetHyperlink(new XLHyperlink("https://example.invalid/"));

        ws.Range("A1:E1")!.InsertColumnsBefore(3);

        await Assert.That(ws.Cell("H1")!.GetString()).IsEqualTo("x");
        await Assert.That(ws.Cell("H1")!.HasHyperlink).IsTrue();
        await Assert.That(ws.Cell("L1")!.HasHyperlink).IsFalse();
    }

    /// <summary>
    /// A delete needs no such correction — the deleted area <em>is</em> the whole edited range and the
    /// shift is that range's own line count — but nothing pinned the two together, so this does.
    /// </summary>
    [Test]
    public async Task A_multi_row_delete_keeps_a_hyperlink_with_its_cell()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        ws.Cell("A10")!.SetValue("x").SetHyperlink(new XLHyperlink("https://example.invalid/"));

        ws.Range("A1:A5")!.Delete(XLShiftDeletedCells.ShiftCellsUp);

        await Assert.That(ws.Cell("A5")!.GetString()).IsEqualTo("x");
        await Assert.That(ws.Cell("A5")!.HasHyperlink).IsTrue();
    }

    /// <summary>The column mirror of <see cref="A_multi_row_delete_keeps_a_hyperlink_with_its_cell"/>.</summary>
    [Test]
    public async Task A_multi_column_delete_keeps_a_hyperlink_with_its_cell()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        ws.Cell("J1")!.SetValue("x").SetHyperlink(new XLHyperlink("https://example.invalid/"));

        ws.Range("A1:E1")!.Delete(XLShiftDeletedCells.ShiftCellsLeft);

        await Assert.That(ws.Cell("E1")!.GetString()).IsEqualTo("x");
        await Assert.That(ws.Cell("E1")!.HasHyperlink).IsTrue();
    }

    /// <summary>
    /// The inserted rectangle is derived three times over and all three must agree:
    /// <c>XLRangeInsertHelper</c> builds it as <c>insertedRange</c> to move the cells,
    /// <c>SheetEdit.Area</c> rebuilds it for hyperlinks, and <c>SheetEdit.CoverageArea</c> rebuilds it
    /// from <c>Range</c> and <c>Shift</c> for conditional formats and data validations. Nothing
    /// asserts that directly, so this asserts it through the features: everything on one cell, an
    /// edited range taller than the shift, and they must all arrive on the same cell. Any one of the
    /// three drifting fails here.
    /// </summary>
    [Test]
    public async Task Cells_hyperlinks_and_coverage_move_by_the_same_inserted_rectangle()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        ws.Cell("A5")!.SetValue("x").SetHyperlink(new XLHyperlink("https://example.invalid/"));
        ws.Cell("A5")!.AddConditionalFormat().WhenNotBlank().Fill.SetBackgroundColor(XLColor.Red);
        ws.Range("A5:A5")!.CreateDataValidation().InputTitle = "Rule";

        ws.Range("A1:A5")!.InsertRowsAbove(3);

        await Assert.That(ws.Cell("A8")!.GetString()).IsEqualTo("x");
        await Assert.That(ws.Cell("A8")!.HasHyperlink).IsTrue();
        await Assert.That(ws.ConditionalFormats.Single().Ranges.Single().RangeAddress.ToString()).IsEqualTo("A8:A8");
        await Assert.That(ws.DataValidations.Single().Ranges.Single().RangeAddress.ToString()).IsEqualTo("A8:A8");
    }
}
