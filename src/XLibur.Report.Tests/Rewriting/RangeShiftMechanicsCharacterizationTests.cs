using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Report.Tests.Rewriting;

/// <summary>
/// What the core library does to a range object when the sheet under it moves.
/// </summary>
/// <remarks>
/// <para>
/// The rule the rewriting work keeps running into: anything the core holds as a <em>live range</em>
/// moves for free, and anything it holds as a value does not. A picture anchor is a live range, which
/// is why pictures needed no code at all; a chart series reference, a pivot cache source and a pivot
/// table's position are values, which is why all three needed rewriting.
/// </para>
/// <para>
/// These tests pin the live half, because <c>&lt;&lt;Pivot&gt;&gt;</c> relies on it: a destination
/// written in template coordinates is held as a one-cell range so that the report's own inserts and
/// deletes carry it, rather than being arithmetic that has to be redone every time the engine gains a
/// way of moving something.
/// </para>
/// </remarks>
public class RangeShiftMechanicsCharacterizationTests
{
    private static (XLWorkbook Workbook, IXLWorksheet Sheet) Sheet()
    {
        var workbook = new XLWorkbook();
        return (workbook, workbook.AddWorksheet("Data"));
    }

    [Test]
    public async Task ACellRangeFollowsRowsInsertedAboveIt()
    {
        var (workbook, sheet) = Sheet();
        using var _ = workbook;

        var cell = sheet.Range("A10:A10");
        sheet.Row(2).InsertRowsBelow(3);

        await Assert.That(cell.RangeAddress.ToString()).IsEqualTo("A13:A13");
    }

    [Test]
    public async Task ACellRangeFollowsARowDeletedAboveIt()
    {
        var (workbook, sheet) = Sheet();
        using var _ = workbook;

        var cell = sheet.Range("A10:A10");
        sheet.Row(2).Delete();

        await Assert.That(cell.RangeAddress.ToString()).IsEqualTo("A9:A9");
    }

    [Test]
    public async Task ACellRangeFollowsAColumnDeletedToItsLeft()
    {
        var (workbook, sheet) = Sheet();
        using var _ = workbook;

        var cell = sheet.Range("D1:D1");
        sheet.Column(2).Delete();

        await Assert.That(cell.RangeAddress.ToString()).IsEqualTo("C1:C1");
    }

    /// <summary>
    /// A range whose own row goes reports itself invalid rather than silently pointing at whatever
    /// moved up into its place — which is what lets a tag tell "the column I named was deleted" from
    /// "the column I named moved".
    /// </summary>
    [Test]
    public async Task ACellRangeWhoseOwnColumnIsDeletedBecomesInvalid()
    {
        var (workbook, sheet) = Sheet();
        using var _ = workbook;

        var cell = sheet.Range("B1:B1");
        sheet.Column(2).Delete();

        await Assert.That(cell.RangeAddress.IsValid).IsFalse();
    }

    /// <summary>
    /// Rows inserted below a range's last row are outside it, so the range does not grow to take them
    /// in. Which is why a pivot's source cannot be a live range captured before expansion: it has to
    /// be taken once the rows exist.
    /// </summary>
    [Test]
    public async Task ARangeDoesNotGrowToCoverRowsInsertedBelowIt()
    {
        var (workbook, sheet) = Sheet();
        using var _ = workbook;

        var range = sheet.Range("A1:C2");
        sheet.Row(2).InsertRowsBelow(3);

        await Assert.That(range.RangeAddress.ToString()).IsEqualTo("A1:C2");
    }

    [Test]
    public async Task ARangeGrowsToCoverRowsInsertedInsideIt()
    {
        var (workbook, sheet) = Sheet();
        using var _ = workbook;

        var range = sheet.Range("A1:C5");
        sheet.Row(2).InsertRowsBelow(3);

        await Assert.That(range.RangeAddress.ToString()).IsEqualTo("A1:C8");
    }

    [Test]
    public async Task ARangeShrinksWhenOneOfItsColumnsGoes()
    {
        var (workbook, sheet) = Sheet();
        using var _ = workbook;

        var range = sheet.Range("A1:C5");
        sheet.Column(2).Delete();

        await Assert.That(range.RangeAddress.ToString()).IsEqualTo("A1:B5");
    }
}
