using System.Collections.Generic;
using System.Linq;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Excel.Coordinates;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// A ladder that attributes the cost of <c>CellsUsed()</c> to its individual layers, plus the range
/// shapes and the early-exit patterns that the aggregate number cannot distinguish.
///
/// Run with:
/// dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- --filter '*UsedCellEnumeration*'
/// </summary>
/// <remarks>
/// Spec 19 area 1, task 1.1. <c>XLiburReadBenchmarks</c> puts <c>CellsUsed()</c> at +2.39 s and
/// +1,020 MB over the load of the same 3.75 M cells, against +0.36 s and +0.00 MB for
/// <c>EnumerateUsedCells()</c>. Three candidate causes were identified by reading
/// <c>XLCells.GetUsedCells</c> and none of them was measured:
/// <list type="bullet">
/// <item>the <c>XLCell</c> wrapper minted per cell,</item>
/// <item>the <c>OrderBy</c>/<c>ThenBy</c>, which buffers and sorts the whole sheet before yielding
/// anything,</item>
/// <item>the <c>HashSet&lt;XLAddress&gt;</c> that retains every visited address to deduplicate.</item>
/// </list>
/// Each rung below adds exactly one of those to the rung beneath it, so a difference is that layer's
/// cost. The spec's rule is that no fix starts until this table says which one dominates.
/// <para>
/// The ladder measures <b>enumeration</b>, not reading: every variant checksums the address rather
/// than the value, so the rungs stay comparable across the two element types (<c>XLUsedCell</c> is a
/// snapshot that already carries its value; <c>XLCell</c> is a handle that would have to fetch one).
/// </para>
/// <para>
/// 50,000 x 10 rather than <c>XLiburReadBenchmarks</c>' 250,000 x 15: the per-cell costs under
/// examination are the same, and a fixture eight times smaller keeps an iteration short enough for
/// BenchmarkDotNet to take a useful number of them.
/// </para>
/// </remarks>
[MemoryDiagnoser]
public class UsedCellEnumerationBenchmarks
{
    private const int Rows = 50_000;
    private const int Cols = 10;

    /// <summary>Ranges in the two multi-range shapes. Ten is a plausible number for a caller unioning columns.</summary>
    private const int RangeCount = 10;

    private XLWorkbook _workbook = null!;
    private XLWorksheet _sheet = null!;
    private XLCellsCollection _cells = null!;
    private Area _full;

    // Concrete XLRanges rather than IXLRanges: the calls below then bind directly instead of
    // through interface dispatch. Kept as `= null!` like the other GlobalSetup-assigned fields in
    // this class rather than `XLRanges?`, because a nullable field would need a null-forgiving
    // operator at every use and warnings are errors here.
    private XLRanges _disjoint = null!;
    private XLRanges _overlapping = null!;

    [GlobalSetup]
    public void Setup()
    {
        SixLaborsV1FontBootstrap.Register();

        _workbook = new XLWorkbook();
        _sheet = (XLWorksheet)_workbook.AddWorksheet("Data");

        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
                _sheet.SetCellValue(r, c, (double)(r * Cols + c));
        }

        _cells = _sheet.Internals.CellsCollection;
        _full = new Area(1, 1, Rows, Cols);

        // Ten column bands that tile the sheet: the same cells as the single range, no duplicates.
        _disjoint = new XLRanges();
        for (var i = 0; i < RangeCount; i++)
            _disjoint.Add(_sheet.Range(1, i + 1, Rows, i + 1));

        // Ten full-width bands stacked so each shares half its rows with the next: every cell in the
        // overlap is produced twice, which is the case the HashSet exists for.
        _overlapping = new XLRanges();
        var band = Rows / RangeCount;
        for (var i = 0; i < RangeCount; i++)
        {
            var first = i * band / 2 + 1;
            _overlapping.Add(_sheet.Range(first, 1, first + band - 1, Cols));
        }
    }

    [GlobalCleanup]
    public void Cleanup() => _workbook.Dispose();

    // ---- the ladder -------------------------------------------------------------------------

    /// <summary>The floor: the slice enumerator, no wrapper, no LINQ.</summary>
    [Benchmark(Baseline = true)]
    public long L1_SliceOnly()
    {
        long sum = 0;
        foreach (var cell in _sheet.EnumerateUsedCells())
            sum += cell.Row + cell.Column;
        return sum;
    }

    /// <summary>+ the <c>XLCell</c> wrapper minted by <c>XLCellsCollection.GetCell</c>.</summary>
    [Benchmark]
    public long L2_PlusWrapper()
    {
        long sum = 0;
        foreach (var cell in _cells.GetCells(_full))
            sum += cell.Address.RowNumber + cell.Address.ColumnNumber;
        return sum;
    }

    /// <summary>+ the per-cell <c>IsEmpty</c> filter and the default predicate, as <c>GetUsedCellsInRange</c> applies them.</summary>
    [Benchmark]
    public long L3_PlusIsEmptyFilter()
    {
        long sum = 0;
        foreach (var cell in _cells.GetCells(_full))
        {
            if (!cell.IsEmpty(XLCellsUsedOptions.AllContents))
                sum += cell.Address.RowNumber + cell.Address.ColumnNumber;
        }

        return sum;
    }

    /// <summary>+ the <c>OrderBy</c>/<c>ThenBy</c> that buffers and sorts the whole sheet.</summary>
    [Benchmark]
    public long L4_PlusSort()
    {
        var ordered = _cells.GetCells(_full)
            .Where(c => !c.IsEmpty(XLCellsUsedOptions.AllContents))
            .OrderBy(c => c.Address.RowNumber)
            .ThenBy(c => c.Address.ColumnNumber);

        long sum = 0;
        foreach (var cell in ordered)
            sum += cell.Address.RowNumber + cell.Address.ColumnNumber;
        return sum;
    }

    /// <summary>+ the <c>HashSet&lt;XLAddress&gt;</c> dedup. This is <c>GetUsedCells</c> bar its GroupBy/SelectMany.</summary>
    [Benchmark]
    public long L5_PlusHashSet()
    {
        var ordered = _cells.GetCells(_full)
            .Where(c => !c.IsEmpty(XLCellsUsedOptions.AllContents))
            .OrderBy(c => c.Address.RowNumber)
            .ThenBy(c => c.Address.ColumnNumber);

        var visited = new HashSet<XLAddress>();
        long sum = 0;
        foreach (var cell in ordered)
        {
            if (visited.Add(cell.Address))
                sum += cell.Address.RowNumber + cell.Address.ColumnNumber;
        }

        return sum;
    }

    /// <summary>The real thing, for the gap to L5.</summary>
    [Benchmark]
    public long L6_CellsUsed()
    {
        long sum = 0;
        foreach (var cell in _sheet.CellsUsed())
            sum += cell.Address.RowNumber + cell.Address.ColumnNumber;
        return sum;
    }

    // ---- range shapes -----------------------------------------------------------------------

    /// <summary>Ten disjoint ranges covering the same cells. The dedup cannot fire; the sort still runs.</summary>
    [Benchmark]
    public long Shape_TenDisjointRanges()
    {
        long sum = 0;
        foreach (var cell in _disjoint.CellsUsed())
            sum += cell.Address.RowNumber + cell.Address.ColumnNumber;
        return sum;
    }

    /// <summary>Ten ranges each sharing half its rows with the next. The dedup fires for every overlapped cell.</summary>
    [Benchmark]
    public long Shape_TenOverlappingRanges()
    {
        long sum = 0;
        foreach (var cell in _overlapping.CellsUsed())
            sum += cell.Address.RowNumber + cell.Address.ColumnNumber;
        return sum;
    }

    // ---- early exit -------------------------------------------------------------------------

    /// <summary>
    /// The pattern the eager sort punishes hardest: the caller wants the first used cell and the
    /// implementation walks and sorts 500,000 of them to hand it over.
    /// </summary>
    [Benchmark]
    // XLCells implements IEnumerable<XLCell> and IEnumerable<IXLCell>, so First() needs the
    // type argument the repo already spells out at its other call sites.
    public int EarlyExit_CellsUsedFirst() => _sheet.CellsUsed().First<XLCell>().Address.RowNumber;

    /// <summary>The same intent through the streaming enumerator, for scale.</summary>
    [Benchmark]
    public int EarlyExit_EnumerateFirst()
    {
        foreach (var cell in _sheet.EnumerateUsedCells())
            return cell.Row;
        return 0;
    }
}
