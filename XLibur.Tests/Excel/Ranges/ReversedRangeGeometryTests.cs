using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Patterns;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Regression tests for spec 36: a range whose corners are given in reverse order
/// (<c>ws.Range("B5:E2")</c>) must behave identically to the equivalent forward range
/// (<c>ws.Range("B2:E5")</c>) everywhere the object model exposes geometry. Each test here
/// pins one of the five defects the spec found, all caused by the area conversion swapping
/// both corners when only one axis was inverted instead of normalising each axis on its own.
/// </summary>
public class ReversedRangeGeometryTests
{
    /// <summary>
    /// Defect 1 (fatal): a conditional format on a range with reversed rows and forward
    /// columns made <see cref="XLWorkbook.SaveAs(string)"/> throw on every save.
    /// </summary>
    [Test]
    public async Task SavingConditionalFormatOnRangeWithReversedRowsDoesNotThrow()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Range("B5:E2").AddConditionalFormat().WhenGreaterThan(5).Fill.SetBackgroundColor(XLColor.Red);

        using var ms = new MemoryStream();
        await Assert.That(() => wb.SaveAs(ms)).ThrowsNothing();
    }

    /// <summary>
    /// Defect 2: a data validation created on a reversed range survived in memory but wrote an
    /// invalid reference on save, and came back as nothing after reload.
    /// </summary>
    [Test]
    public async Task DataValidationOnReversedRangeSurvivesSaveAndReload()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Range("B5:E2").CreateDataValidation().WholeNumber.Between(0, 100);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        using var wb2 = new XLWorkbook(ms);
        var ws2 = wb2.Worksheet("Sheet1");
        var reloadedValidations = ws2.DataValidations.ToList();

        await Assert.That(reloadedValidations.Count).IsEqualTo(1);
        var reloadedRanges = reloadedValidations[0].Ranges.ToList();
        await Assert.That(reloadedRanges.Count).IsEqualTo(1);
        await Assert.That(reloadedRanges[0].RangeAddress.ToString()).IsEqualTo("B2:E5");
    }

    /// <summary>
    /// Defect 3: applying a style to a reversed range wrote no cells, while assigning a value
    /// to the same range wrote all of them.
    /// </summary>
    [Test]
    public async Task StyleAppliedToReversedRangeStylesItsCells()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var range = ws.Range("B5:E2");
        range.Style.Fill.SetBackgroundColor(XLColor.Yellow);

        foreach (var cell in ws.Range("B2:E5").Cells())
        {
            await Assert.That(cell.Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);
        }
    }

    /// <summary>
    /// Defect 4: <c>RowCount()</c> and <c>ColumnCount()</c>
    /// returned negative numbers for a reversed range, while the range address's spans and
    /// <c>Cells().Count()</c> were correct.
    /// </summary>
    [Test]
    public async Task RowCountAndColumnCountOnReversedRangeReturnPositiveMagnitudes()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var range = ws.Range("B5:E2");

        await Assert.That(range.RowCount()).IsEqualTo(4);
        await Assert.That(range.ColumnCount()).IsEqualTo(4);
        await Assert.That(range.Cells().Count()).IsEqualTo(16);
    }

    /// <summary>
    /// Defect 5: <see cref="IXLRanges.Consolidate"/> returned an empty collection for a
    /// reversed range where it should have returned the equivalent forward range.
    /// </summary>
    [Test]
    public async Task ConsolidateIncludesReversedRange()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var ranges = new XLRanges { ws.Range("B5:E2") };

        var consolidated = ranges.Consolidate().ToList();

        await Assert.That(consolidated.Count).IsEqualTo(1);
        await Assert.That(consolidated[0].RangeAddress.ToString()).IsEqualTo("B2:E5");
    }

    private sealed class AddressableStub : IXLAddressable
    {
        public AddressableStub(IXLRangeAddress rangeAddress) => RangeAddress = rangeAddress;
        public IXLRangeAddress RangeAddress { get; }
    }

    /// <summary>
    /// A sixth, related defect (spec user story 10, not one of the five named ones): the
    /// range-index QuadTree (behind data validations, conditional formats and <c>XLRanges</c>
    /// once a worksheet holds 20 or more of them) compared a reversed range's raw corners
    /// directly against each quadrant's bounds and, at the leaf, against the query via
    /// <c>XLRangeAddress.Intersects</c> - which itself assumes both sides are already
    /// normalised. E8194:E8190 (rows reversed) straddles the boundary between the QuadTree's
    /// two level-1 row bands (1..8192 and 8193..), so <c>Covers</c> is false for both children
    /// and the range is stored at the root - isolating the leaf-level <c>Intersects</c> check,
    /// which returned false for a query cell that is inside the range.
    /// </summary>
    [Test]
    public async Task QuadTreeFindsReversedRangeThatSpansAQuadrantBoundary()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var entryAddress = ws.Range("E8194:E8190").RangeAddress;
        var queryAddress = ws.Range("E8192").RangeAddress;

        var quadrant = new Quadrant();
        quadrant.Add(new AddressableStub(entryAddress));

        var found = quadrant.GetIntersectedRanges(queryAddress).ToList();

        await Assert.That(found.Count).IsEqualTo(1);
    }
}
