using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
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

    /// <summary>
    /// Follow-up finding, flagged by the branch's own code review: below the QuadTree's
    /// 20-range promotion threshold, <c>XLRangeIndex</c> compares ranges with a linear
    /// scan through <c>IXLRangeAddress.Intersects</c> directly - which, like the QuadTree's own
    /// pre-fix comparisons, assumes both sides are already normalised. A data validation's
    /// *first* range is indexed from the already-normalised <c>Ranges</c> projection, so this
    /// only shows up for a range added to an existing validation's coverage after the fact,
    /// where the raw (possibly reversed) address reaches the index directly.
    /// </summary>
    [Test]
    public async Task DataValidationIndexFindsReversedRangeBeforePromotion()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("J1").CreateDataValidation().WholeNumber.Between(0, 100);
        ws.Cell("J2").CreateDataValidation().WholeNumber.Between(0, 100);
        var dv = ws.Cell("A1").CreateDataValidation();
        dv.WholeNumber.Between(0, 100);
        dv.AddRange(ws.Range("B5:E2")); // reversed rows, added to already-registered coverage

        var found = ws.DataValidations.GetAllInRange(ws.Range("C3").RangeAddress).ToList();

        await Assert.That(found.Count).IsEqualTo(1);
    }

    /// <summary>
    /// Follow-up finding: merged ranges are backed by the same range index, and a lone merge
    /// never reaches the 20-range promotion threshold, so <c>Merge()</c> on a reversed range hit
    /// the flat-list defect above. Separately, <c>Quadrant.CoversAnyRange</c> and
    /// <c>GetIntersectedRanges(IXLAddress)</c> - reached only once an index promotes - compared a
    /// stored range's raw corners against a point directly via
    /// <c>XLRangeAddress.Contains(in XLAddress)</c>, the same unguarded assumption one level
    /// down. Twenty merges past the reversed one forces promotion, isolating that second path.
    /// </summary>
    [Test]
    public async Task MergedReversedRangeIsRecognisedBeforeAndAfterPromotion()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Range("B5:E2").Merge(); // reversed rows, forward columns

        await Assert.That(ws.Cell("C3").IsMerged()).IsTrue();
        var mergedRange = ws.Cell("C3").MergedRange();
        await Assert.That(mergedRange).IsNotNull();
        await Assert.That(mergedRange!.RowCount()).IsEqualTo(4);
        await Assert.That(mergedRange.ColumnCount()).IsEqualTo(4);

        for (var i = 1; i <= 20; i++)
            ws.Range(i, 10, i, 11).Merge(checkIntersect: false);

        await Assert.That(ws.Cell("C3").IsMerged()).IsTrue();
        var mergedRangeAfterPromotion = ws.Cell("C3").MergedRange();
        await Assert.That(mergedRangeAfterPromotion).IsNotNull();
        await Assert.That(mergedRangeAfterPromotion!.RowCount()).IsEqualTo(4);
        await Assert.That(mergedRangeAfterPromotion.ColumnCount()).IsEqualTo(4);
    }

    /// <summary>
    /// Follow-up finding: <c>XLTable.DataRange</c> computes its cells via the relative
    /// <c>Range(int,int,int,int)</c> overload, which anchored its offsets to
    /// <c>RangeAddress.FirstAddress</c> directly - the top-left corner only when the address
    /// happens to be normalised. For "B5:E2" that corner is row 5, so the computed data range
    /// fell entirely outside the table's own bounds and <c>GetRange</c>'s bounds check (the same
    /// unguarded assumption, one level down) threw on save.
    /// </summary>
    [Test]
    public async Task TableOverReversedRangeSavesWithTheForwardRangeAsRef()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Range("B5:E2").CreateTable("MyTable");

        using var ms = new MemoryStream();
        await Assert.That(() => wb.SaveAs(ms)).ThrowsNothing();

        using var wb2 = new XLWorkbook(ms);
        var table = wb2.Worksheet("Sheet1").Table(0);
        await Assert.That(table.RangeAddress.ToString()).IsEqualTo("B2:E5");
    }

    /// <summary>
    /// Follow-up finding, from the codebase-wide audit: <c>XLRangeColumn.CellCount()</c> and
    /// <c>XLRangeRow.CellCount()</c> each duplicated the exact defect already fixed on
    /// <c>XLRangeBase.RowCount()</c>/<c>ColumnCount()</c> - computing directly from
    /// <c>RangeAddress.LastAddress - RangeAddress.FirstAddress</c> instead of delegating to the
    /// now-fixed base method.
    /// </summary>
    [Test]
    public async Task RangeColumnAndRangeRowCellCountOnReversedRangeReturnPositiveMagnitudes()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var range = ws.Range("B5:E2");

        await Assert.That(range.Column(1).CellCount()).IsEqualTo(4);
        await Assert.That(range.Row(1).CellCount()).IsEqualTo(4);
    }

    /// <summary>
    /// Follow-up finding, from the branch's own code review: <c>RowCount()</c>/<c>ColumnCount()</c>
    /// were normalised onto the rectangle but the members that *address* a cell relative to the
    /// range - <c>Cell(row, column)</c>, and <c>FirstCell()</c>/<c>LastCell()</c> through it - still
    /// anchored on <c>RangeAddress.FirstAddress</c>. The two then disagreed, and the range walked
    /// off its own bottom edge: <c>LastCell()</c> on "B5:E2" is row 5 + 4 - 1 = 8.
    /// </summary>
    [Test]
    public async Task CellsAddressedRelativeToAReversedRangeStayInsideIt()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var range = ws.Range("B5:E2");

        await Assert.That(range.FirstCell().Address.ToString()).IsEqualTo("B2");
        await Assert.That(range.LastCell().Address.ToString()).IsEqualTo("E5");
        await Assert.That(range.Cell(2, 3).Address.ToString()).IsEqualTo("D3");
    }

    /// <summary>
    /// The same defect one level up: <c>Rows()</c>/<c>Columns()</c> walk 1..RowCount() through
    /// <c>Row(int)</c>/<c>Column(int)</c>, which anchored on <c>RangeAddress.FirstAddress</c> too.
    /// For "B5:E2" that produced B5:E5..B8:E8 - three rows entirely outside the range - so a style
    /// or value written through <c>Rows()</c> landed on unrelated cells.
    /// </summary>
    [Test]
    public async Task RowsAndColumnsOfAReversedRangeAreTheForwardRangeMembers()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        var rows = ws.Range("B5:E2").Rows().Select(r => r.RangeAddress.ToString()!).ToList();
        await Assert.That(rows).IsEquivalentTo(new[] { "B2:E2", "B3:E3", "B4:E4", "B5:E5" });

        // Columns reversed instead of rows, to pin the other axis independently.
        var columns = ws.Range("E2:B5").Columns().Select(c => c.RangeAddress.ToString()!).ToList();
        await Assert.That(columns).IsEquivalentTo(new[] { "B2:B5", "C2:C5", "D2:D5", "E2:E5" });
    }

    /// <summary>
    /// The user-visible consequence of the two defects above: writing through <c>Rows()</c> used to
    /// style rows 5..8 instead of 2..5, silently mutating cells the caller never named.
    /// </summary>
    [Test]
    public async Task StylingThroughRowsOfAReversedRangeTouchesOnlyItsOwnCells()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");

        foreach (var row in ws.Range("B5:E2").Rows())
            row.Style.Fill.SetBackgroundColor(XLColor.Yellow);

        foreach (var cell in ws.Range("B2:E5").Cells())
            await Assert.That(cell.Style.Fill.BackgroundColor).IsEqualTo(XLColor.Yellow);

        await Assert.That(ws.Cell("B8").Style.Fill.BackgroundColor).IsNotEqualTo(XLColor.Yellow);
    }

    /// <summary>
    /// The used-cell probes search the normalised rectangle but converted the absolute row/column
    /// they found back to a relative index against <c>RangeAddress.FirstAddress</c>, so on a
    /// reversed range the two halves disagreed and the resulting index was negative.
    /// </summary>
    [Test]
    public async Task UsedRowsAndColumnsOfAReversedRangeAreInsideIt()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Cell("C3").Value = 1;
        ws.Cell("D4").Value = 1;
        var range = ws.Range("B5:E2");

        await Assert.That(range.FirstRowUsed()!.RangeAddress.ToString()).IsEqualTo("B3:E3");
        await Assert.That(range.LastRowUsed()!.RangeAddress.ToString()).IsEqualTo("B4:E4");
        await Assert.That(range.FirstColumnUsed()!.RangeAddress.ToString()).IsEqualTo("C2:C5");
        await Assert.That(range.LastColumnUsed()!.RangeAddress.ToString()).IsEqualTo("D2:D5");
    }

    /// <summary>
    /// A merged reversed range reaches the same relative addressing through
    /// <c>MergedRange()</c>, so its last cell used to sit outside the merge.
    /// </summary>
    [Test]
    public async Task LastCellOfAMergedReversedRangeIsInsideTheMerge()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Range("B5:E2").Merge();

        var mergedRange = ws.Cell("C3").MergedRange();

        await Assert.That(mergedRange).IsNotNull();
        await Assert.That(mergedRange!.FirstCell().Address.ToString()).IsEqualTo("B2");
        await Assert.That(mergedRange.LastCell().Address.ToString()).IsEqualTo("E5");
    }

    /// <summary>
    /// Follow-up finding: <c>XLRangeBase.SheetRange</c> throws for a <c>#REF!</c> address, and a
    /// range destroyed by a delete keeps exactly that - <c>XLRangeShiftHelper</c> assigns
    /// <c>XLWorksheet.InvalidAddress</c> to both corners. Routing the counts through it therefore
    /// turned a working public-API call into a throw, and added throw sites to the save and copy
    /// paths that call these on stored ranges. The normalisation has to tolerate <c>#REF!</c>.
    /// </summary>
    [Test]
    public async Task CountsOnARangeDestroyedByADeleteDoNotThrow()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var range = ws.Range("A2:B3");

        ws.Rows(1, 5).Delete(); // swallows the range whole, leaving it #REF!

        await Assert.That(() => _ = range.RowCount()).ThrowsNothing();
        await Assert.That(() => _ = range.ColumnCount()).ThrowsNothing();
    }

    /// <summary>
    /// Follow-up finding: the calc engine's <c>Reference</c> used to reject an un-normalised area
    /// with an <see cref="System.ArgumentException"/>. Removing that check left its own internals
    /// - which iterate first-to-last, and size <c>Apply</c>'s rectangle off <c>FirstAddress</c>
    /// with the absolute spans - trusting an invariant nothing enforced any more, so a reversed
    /// area would have produced a silently wrong formula result rather than an exception. The
    /// constructors normalise instead of refusing.
    /// </summary>
    [Test]
    public async Task ReferenceNormalisesAReversedAreaInsteadOfStoringItVerbatim()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var reversed = (XLRangeAddress)ws.Range("B5:E2").RangeAddress;
        var alsoReversed = (XLRangeAddress)ws.Range("H4:G1").RangeAddress;

        var single = new Reference(reversed);
        await Assert.That(single[0].ToString()).IsEqualTo("B2:E5");

        var fromList = new Reference(new List<XLRangeAddress> { reversed, alsoReversed });
        await Assert.That(fromList[0].ToString()).IsEqualTo("B2:E5");
        await Assert.That(fromList[1].ToString()).IsEqualTo("G1:H4");

        // XLRanges orders what it stores, so compare as a set rather than by index.
        var fromRanges = new Reference(new XLRanges { ws.Range("B5:E2"), ws.Range("H4:G1") });
        var areas = new List<string> { fromRanges[0].ToString()!, fromRanges[1].ToString()! };
        await Assert.That(areas).IsEquivalentTo(new[] { "B2:E5", "G1:H4" });
    }

    /// <summary>
    /// The normalised areas have to be usable, not merely tidy: iterating a reference built from a
    /// reversed area yields its cells rather than nothing, which is what the first-to-last loop in
    /// <c>GetCellsValues</c> did while the invariant went unenforced.
    /// </summary>
    [Test]
    public async Task ReferenceBuiltFromAReversedAreaCoversItsCells()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        var reference = new Reference((XLRangeAddress)ws.Range("B5:E2").RangeAddress);

        await Assert.That(reference.NumberOfCells).IsEqualTo(16);
        await Assert.That(reference[0].Contains(ws.Cell("C3").Address)).IsTrue();
    }

    /// <summary>
    /// User story 9: a reversed range used as a formula reference evaluates rather than
    /// throwing. The calc engine's <c>Reference</c> type used to reject an un-normalised
    /// <c>XLRangeAddress</c> defensively; that precondition is now removed because every path
    /// that reaches it already normalises first (formula-text parsing per axis in
    /// <c>AstNode.BuildAddress</c>, or geometry sourced from <c>Area</c>).
    /// </summary>
    [Test]
    public async Task FormulaReferencingReversedRangeEvaluates()
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Sheet1");
        ws.Range("B2:E5").Value = 2;
        ws.Cell("G1").FormulaA1 = "=SUM(B5:E2)";

        await Assert.That(ws.Cell("G1").GetValue<double>()).IsEqualTo(32d);
    }
}
