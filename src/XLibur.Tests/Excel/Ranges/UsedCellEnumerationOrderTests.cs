using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Ranges;

/// <summary>
/// Pins the enumeration contract of <c>CellsUsed()</c>: row-major order, every used cell once.
/// </summary>
/// <remarks>
/// <c>XLCells.GetUsedCells</c> streams a single range straight off the slice merge, on the grounds
/// that the merge already yields ascending row-major with no duplicates, and falls back to a sorted,
/// deduplicated path for everything else. These tests hold both halves to the same observable
/// contract so the two cannot drift, and they are deliberately characterisation tests - they pass
/// against the sort-everything implementation that preceded the fast path, which is what makes them
/// worth having.
/// </remarks>
public class UsedCellEnumerationOrderTests
{
    /// <summary>The addresses in the order they were yielded, joined into one string.</summary>
    /// <remarks>
    /// A joined string rather than a collection assertion on purpose: TUnit's <c>IsEquivalentTo</c>
    /// ignores order unless it is handed <c>CollectionOrdering.Matching</c>, and order is the whole
    /// point of these tests. <c>IsEqualTo</c> over a string cannot be order-blind by accident.
    /// </remarks>
    private static string Addresses(IEnumerable<IXLCell> cells) =>
        string.Join(",", cells.Select(c => c.Address.ToStringRelative()));

    /// <summary>The fast path: one range, default options.</summary>
    [Test]
    public async Task SingleRangeYieldsRowMajor()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        // Assigned deliberately out of order, so a pass-through of insertion order would fail.
        ws.Cell(3, 2).Value = "C";
        ws.Cell(1, 5).Value = "A";
        ws.Cell(2, 1).Value = "B";
        ws.Cell(1, 2).Value = "Z";

        await Assert.That(Addresses(ws.CellsUsed())).IsEqualTo("B1,E1,A2,B3");
    }

    /// <summary>
    /// A column-shaped range and a row-shaped one over the same sheet, which between them make the
    /// concatenation of per-range results non-row-major. The slow path has to restore the order.
    /// </summary>
    [Test]
    public async Task DisjointRangesYieldRowMajorAcrossRanges()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        for (var row = 1; row <= 3; row++)
        {
            ws.Cell(row, 1).Value = $"A{row}";
            ws.Cell(row, 3).Value = $"C{row}";
        }

        var ranges = new XLRanges();
        ranges.Add(ws.Range("C1:C3"));
        ranges.Add(ws.Range("A1:A3"));

        await Assert.That(Addresses(ranges.CellsUsed()))
            .IsEqualTo("A1,C1,A2,C2,A3,C3");
    }

    /// <summary>Overlapping ranges must yield each shared cell exactly once, still in row-major order.</summary>
    [Test]
    public async Task OverlappingRangesYieldEachCellOnce()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        for (var row = 1; row <= 4; row++)
        {
            for (var col = 1; col <= 2; col++)
                ws.Cell(row, col).Value = row * 10 + col;
        }

        var ranges = new XLRanges();
        ranges.Add(ws.Range("A1:B3"));
        ranges.Add(ws.Range("A2:B4")); // rows 2-3 are in both

        await Assert.That(Addresses(ranges.CellsUsed()))
            .IsEqualTo("A1,B1,A2,B2,A3,B3,A4,B4");
    }

    /// <summary>
    /// The same range added twice: every cell is produced twice by the enumeration and must be
    /// yielded once. This is the case the visited-set existed for.
    /// </summary>
    [Test]
    public async Task SameRangeTwiceYieldsEachCellOnce()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell(1, 1).Value = 1;
        ws.Cell(1, 2).Value = 2;
        ws.Cell(2, 1).Value = 3;

        var ranges = new XLRanges();
        ranges.Add(ws.Range("A1:B2"));
        ranges.Add(ws.Range("A1:B2"));

        await Assert.That(Addresses(ranges.CellsUsed())).IsEqualTo("A1,B1,A2");
    }

    /// <summary>
    /// A merged range contributes cells from outside the slices, and those arrive after the slice
    /// cells rather than in position - the reason the sorted path exists. The single-range fast path
    /// must not be taken here.
    /// </summary>
    [Test]
    public async Task MergedRangeCandidatesAreYieldedInRowMajorOrder()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        // A value low on the sheet, and a merge higher up whose non-anchor cells hold nothing.
        ws.Cell(3, 1).Value = "anchor";
        ws.Range("B1:C1").Merge();

        var options = XLCellsUsedOptions.AllContents | XLCellsUsedOptions.MergedRanges;

        await Assert.That(Addresses(ws.CellsUsed(options))).IsEqualTo("B1,C1,A3");
    }

    /// <summary>
    /// Asking for merged ranges on a sheet that has none must not change what is yielded. The fast
    /// path tests the sheet as well as the options, and this is what that test is for.
    /// </summary>
    [Test]
    public async Task MergedRangeOptionOnSheetWithoutMergesMatchesDefault()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell(1, 2).Value = "B1";
        ws.Cell(2, 1).Value = "A2";

        var withOption = Addresses(ws.CellsUsed(XLCellsUsedOptions.AllContents | XLCellsUsedOptions.MergedRanges));

        await Assert.That(withOption).IsEqualTo(Addresses(ws.CellsUsed()));
    }

    /// <summary>
    /// A data validation makes a cell non-empty without putting anything in the value slice, so it
    /// arrives through the candidate sequence.
    /// </summary>
    [Test]
    public async Task DataValidationCandidatesAreIncludedAndOrdered()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws.Cell(4, 1).Value = "value";
        ws.Range("B2:B2").CreateDataValidation().WholeNumber.EqualOrGreaterThan(1);

        var options = XLCellsUsedOptions.AllContents | XLCellsUsedOptions.DataValidation;

        await Assert.That(Addresses(ws.CellsUsed(options))).IsEqualTo("B2,A4");
    }

    /// <summary>
    /// A predicate must be applied on the streaming path exactly as it was on the sorted one.
    /// </summary>
    [Test]
    public async Task PredicateFiltersOnTheSingleRangePath()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        for (var row = 1; row <= 4; row++)
            ws.Cell(row, 1).Value = row;

        // GetValue<int> rather than a modulo over GetDouble: the cells hold whole numbers, and an
        // exact equality test against a floating-point result is fragile even when it happens to
        // hold (S1244).
        var even = ws.CellsUsed(c => c.GetValue<int>() % 2 == 0);

        await Assert.That(Addresses(even)).IsEqualTo("A2,A4");
    }

    /// <summary>
    /// The streaming path must not evaluate past what the caller asks for. Against the sorted
    /// implementation this walked and buffered the whole sheet to answer with its first element.
    /// </summary>
    [Test]
    public async Task FirstUsedCellIsReachableWithoutWalkingTheSheet()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        for (var row = 1; row <= 5_000; row++)
            ws.Cell(row, 1).Value = row;

        var visited = 0;
        var first = ws.CellsUsed(c =>
        {
            visited++;
            return true;
        }).First();

        await Assert.That(first.Address.ToStringRelative()).IsEqualTo("A1");

        // The predicate runs twice per cell on this path (once as the slice filter, once in the
        // used-cell test), so the bound is per-cell rather than exactly one call.
        await Assert.That(visited).IsLessThanOrEqualTo(4);
    }
}
