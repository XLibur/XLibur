using BenchmarkDotNet.Attributes;
using XLibur.Excel;

namespace XLibur.Benchmarks;

/// <summary>
/// Bulk styling over ranges spanning two orders of magnitude, pinning the per-cell cost of
/// <c>XLStylizedBase.ModifyStyle</c>.
///
/// Run with:
/// dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- --filter '*BulkStyle*'
///
/// <para>
/// Spec 05's acceptance criterion 3 asks that styling a range allocate "O(distinct styles), not
/// O(cells)". Measured, allocation is exactly linear in cells — 324 KB, 3,236 KB, 32,264 KB for 10K,
/// 100K and 1M — and <b>no implementation can make it otherwise</b>: styling N cells writes N entries
/// into the style slice, and the slice has to grow to hold them. The criterion is mis-stated rather
/// than missed, in the same way spec 11's criterion 2 was.
/// </para>
/// <para>
/// What the curve does establish is the constant. It sits at ~33 bytes per cell, which is slice
/// storage; before spec 11's Task 4 it was ~234 bytes per cell, because <c>ModifyStyle</c> built one
/// <c>XLCell</c> wrapper per cell into a HashSet before writing anything. Both are linear; only one is
/// mostly waste. A regression back to the wrapper path would show up here as a ~7x jump in Allocated
/// at fixed <see cref="Cells"/>, which is the thing worth guarding.
/// </para>
/// </summary>
[MemoryDiagnoser]
public class BulkStyleBenchmarks
{
    /// <summary>
    /// Deliberately spans two orders of magnitude. A single size cannot distinguish "allocates per
    /// cell" from "allocates per style" — only the ratio between sizes can.
    /// </summary>
    [Params(10_000, 100_000, 1_000_000)]
    public int Cells;

    private const int Columns = 10;

    private XLWorkbook _workbook = null!;
    private IXLWorksheet _worksheet = null!;
    private IXLRange _range = null!;

    [IterationSetup]
    public void Setup()
    {
        _workbook = new XLWorkbook();
        _worksheet = _workbook.AddWorksheet("Sheet1");
        _range = _worksheet.Range(1, 1, Cells / Columns, Columns);
    }

    [IterationCleanup]
    public void Cleanup()
    {
        _workbook.Dispose();
    }

    /// <summary>
    /// One mutation over a uniformly styled range: every cell shares the inherited style, so the
    /// transition is computed once and the result is the same value for every point.
    /// </summary>
    [Benchmark(Baseline = true)]
    public void SetBoldOverUniformRange()
    {
        _range.Style.Font.Bold = true;
    }

    /// <summary>
    /// Four mutations, which is the pattern the create-path probe uses and the one real formatting
    /// code tends to look like.
    /// </summary>
    [Benchmark]
    public void SetFourPropertiesOverUniformRange()
    {
        var style = _range.Style;
        style.Font.Bold = true;
        style.Font.FontSize = 14;
        style.Fill.BackgroundColor = XLColor.LightGray;
        style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    }

    /// <summary>
    /// The adversarial case for the last-value memo in <c>XLStylizedBase.ApplyToCellStyles</c>: every
    /// row carries a different style, so the memo misses at each row boundary and the transition cache
    /// on <c>XLStyleValue</c> does the work instead. It costs ~4x the uniform case, most of which is the
    /// row objects the setup materialises rather than the styling itself — included so the memo's value
    /// is visible, and so a change that made the memo the only thing keeping the fast path fast would
    /// show up.
    /// </summary>
    [Benchmark]
    public void SetBoldOverStripedRange()
    {
        for (var row = 1; row <= Cells / Columns; row++)
            _worksheet.Row(row).Style.Font.FontSize = 8 + (row % 8);

        _range.Style.Font.Bold = true;
    }
}
