using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// The <c>Area</c> a sheet listener is handed is the edited range extended by <c>Shift - 1</c> lines,
/// so the shift magnitude is <b>not</b> recoverable from the area when the edited range spans more
/// than one line on the shift axis. This is why <c>SheetEdit</c> carries <c>Range</c> and
/// <c>Shift</c> as well as <c>Area</c>, and it is the premise spec 33 task 3 step 1 exists to
/// settle. If these tests fail, the premise is wrong and the port should be narrowed back to
/// <c>(sheet, area)</c>.
/// </summary>
public class SheetEditAreaTests
{
    /// <summary>
    /// The area's height is read off a hyperlink placed well below the edit: <c>XLHyperlinks</c>
    /// shifts a fully-covered area by <c>insertedArea.Height</c>, so how far the hyperlink travels
    /// <em>is</em> the listener's area height. A one-row range recovers the shift (3 == 3); a
    /// five-row range does not (7 != 3).
    /// </summary>
    [Test]
    [Arguments("A1:A1", 3, 3)]
    [Arguments("A1:A5", 3, 7)]
    [Arguments("A1:A3", 2, 4)]
    public async Task The_listener_area_is_the_whole_range_extended_by_shift_minus_one(
        string rangeAddress, int shift, int expectedAreaHeight)
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        ws.Cell("A20")!.SetValue("x").SetHyperlink(new XLHyperlink("https://example.invalid/"));

        ws.Range(rangeAddress)!.InsertRowsAbove(shift);

        await Assert.That(ws.Cell(20 + expectedAreaHeight, 1)!.HasHyperlink).IsTrue();
    }

    /// <summary>
    /// The consequence, pinned so that a later spec fixing it has to change this test deliberately:
    /// the cells move by <c>Shift</c> and the listeners move by the area's height, so for a range
    /// taller than one row a hyperlink parts company with its own cell. Recorded as D15 in
    /// <c>DEFECTS.md</c>; spec 33 preserves it rather than fixing it.
    /// </summary>
    [Test]
    public async Task A_multi_row_edited_range_detaches_a_hyperlink_from_its_cell()
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet("S");
        ws.Cell("A5")!.SetValue("x").SetHyperlink(new XLHyperlink("https://example.invalid/"));

        ws.Range("A1:A5")!.InsertRowsAbove(3);

        await Assert.That(ws.Cell("A8")!.GetString()).IsEqualTo("x");
        await Assert.That(ws.Cell("A8")!.HasHyperlink).IsFalse();
        await Assert.That(ws.Cell("A12")!.HasHyperlink).IsTrue();
    }
}
