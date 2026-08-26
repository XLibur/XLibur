using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Drawings;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// What a two-cell picture anchor does under every kind of structural edit.
/// <para>
/// The picture is the control in spec 33's evidence: it was the one drawing anchor that already
/// moved, because <c>XLMarker</c> stored a one-cell <c>IXLRange</c> purely so the range repository
/// would shift it — the seam smuggled past, in a comment. These cases were read off that mechanism
/// on the unmodified tree and are what <c>GridShift</c> reproduces, so that charts, notes, panes and
/// pivot tables move the way the one working anchor already did.
/// </para>
/// <para>
/// Spec 33 task 6 deletes the workaround and moves the picture onto <c>GridShift</c> too. Every case
/// here must survive that unchanged, which is what makes it a gate rather than a description.
/// </para>
/// </summary>
public class PictureAnchorShiftTests
{
    private static IXLPicture TwoCell(IXLWorksheet ws, string from, string to)
    {
        using var stream = System.Reflection.Assembly.GetAssembly(typeof(XLibur.Examples.BasicTable))!
            .GetManifestResourceStream("XLibur.Examples.Resources.SampleImage.jpg")!;
        return ws.AddPicture(stream, "p").MoveTo(ws.Cell(from)!, ws.Cell(to)!);
    }

    [Test]
    [Arguments("insert above", 1, 3, "C7", "J23")]
    [Arguments("insert inside", 10, 3, "C4", "J23")]
    [Arguments("insert below", 30, 3, "C4", "J20")]
    [Arguments("delete above", 1, -2, "C2", "J18")]
    [Arguments("delete inside", 5, -2, "C4", "J18")]
    [Arguments("delete covering the whole anchor", 1, -25, "C1", "J1")]
    public async Task A_row_edit_moves_a_two_cell_picture_anchor(
        string label, int at, int shift, string expectedFrom, string expectedTo)
    {
        _ = label;
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        var picture = TwoCell(ws, "C4", "J20");

        if (shift > 0)
            ws.Row(at).InsertRowsAbove(shift);
        else
            ws.Rows(at, at - shift - 1).Delete();

        await Assert.That(picture.TopLeftCell.Address.ToString()).IsEqualTo(expectedFrom);
        await Assert.That(picture.BottomRightCell.Address.ToString()).IsEqualTo(expectedTo);
        await Assert.That(ws.Pictures.Count).IsEqualTo(1);
    }

    [Test]
    [Arguments("insert left", 1, 2, "E4", "L20")]
    [Arguments("insert inside", 5, 2, "C4", "L20")]
    [Arguments("delete left", 1, -2, "A4", "H20")]
    public async Task A_column_edit_moves_a_two_cell_picture_anchor(
        string label, int at, int shift, string expectedFrom, string expectedTo)
    {
        _ = label;
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        var picture = TwoCell(ws, "C4", "J20");

        if (shift > 0)
            ws.Column(at).InsertColumnsBefore(shift);
        else
            ws.Columns(at, at - shift - 1).Delete();

        await Assert.That(picture.TopLeftCell.Address.ToString()).IsEqualTo(expectedFrom);
        await Assert.That(picture.BottomRightCell.Address.ToString()).IsEqualTo(expectedTo);
    }

    /// <summary>
    /// A drawing outside the edited columns does not move: inserting cells into <c>B1:B5</c> and
    /// shifting them down leaves an anchor in columns C to J alone. This is the cross-axis coverage
    /// rule, and it is why the transform takes the edited range and not just the shift.
    /// </summary>
    [Test]
    public async Task A_partial_insert_outside_the_anchor_columns_does_not_move_it()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        var picture = TwoCell(ws, "C4", "J20");

        ws.Range("B1:B5")!.InsertRowsAbove(3);

        await Assert.That(picture.TopLeftCell.Address.ToString()).IsEqualTo("C4");
        await Assert.That(picture.BottomRightCell.Address.ToString()).IsEqualTo("J20");
    }

    /// <summary>
    /// A delete starting <b>exactly</b> on the anchor's first row leaves that anchor unreadable:
    /// the repository path marks the range's first address invalid rather than clamping it, so
    /// reading <c>TopLeftCell</c> throws. Every neighbouring case clamps.
    /// <para>
    /// WRONG on purpose, and recorded as D16 in <c>DEFECTS.md</c>. Spec 33 task 6 moves the picture
    /// onto <c>GridShift</c>, which clamps here like everywhere else, and re-points this test.
    /// </para>
    /// </summary>
    [Test]
    public async Task A_delete_starting_on_the_anchor_row_leaves_it_invalid_yet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("S");
        var picture = TwoCell(ws, "C4", "J20");

        ws.Rows(4, 20).Delete();

        await Assert.That(() => picture.TopLeftCell)
            .Throws<System.ArgumentOutOfRangeException>();
    }
}
