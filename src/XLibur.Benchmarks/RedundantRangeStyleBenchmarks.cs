using System;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;

namespace XLibur.Benchmarks;

/// <summary>
/// Setting a style value through a range, or a worksheet, whose cells already hold that value.
///
/// Run with:
/// dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- --filter '*RedundantRangeStyle*'
/// </summary>
/// <remarks>
/// Before #505 a setter on a range or a worksheet skipped the write when the value equalled the
/// container's own record of its style, so setting a value again through a range the caller held
/// cost nothing. That record is not the cells' style - a range rebuilt after a garbage collection
/// starts from its parent's - so the skip is gone and the write always reaches the cells. These
/// benchmarks price that case against <see cref="SetBoldOnce"/>, the same range styled for the first
/// time, which both versions have always paid in full.
/// </remarks>
[MemoryDiagnoser]
public class RedundantRangeStyleBenchmarks
{
    [Params(10_000, 100_000)]
    public int Cells { get; set; }

    private const int Columns = 10;

    // Assigned per iteration by the setups, so never observed null despite the declaration.
    private XLWorkbook _workbook = null!;
    private IXLWorksheet _worksheet = null!;
    private IXLRange _range = null!;

    private void CreateRange()
    {
        _workbook = new XLWorkbook();
        _worksheet = _workbook.AddWorksheet("Sheet1");
        _range = _worksheet.Range(1, 1, Cells / Columns, Columns);
    }

    [IterationSetup(Target = nameof(SetBoldOnce))]
    public void SetupUnstyled() => CreateRange();

    /// <summary>
    /// The range is styled here, off the measured path, and held in a field, so the range object
    /// the benchmark writes through is the one whose record already carries the value.
    /// </summary>
    [IterationSetup(Targets = [nameof(SetBoldAgainThroughHeldRange), nameof(SetBackgroundColourAgainThroughHeldRange)])]
    public void SetupStyledRange()
    {
        CreateRange();
        _range.Style.Font.Bold = true;
        _range.Style.Fill.BackgroundColor = XLColor.LightGray;
    }

    [IterationSetup(Target = nameof(SetBoldAgainOnWorksheet))]
    public void SetupStyledWorksheet()
    {
        CreateRange();
        _range.Style.Font.Italic = true; // materialise the cells, so the sheet has something to walk
        _worksheet.Style.Font.Bold = true;
    }

    /// <summary>
    /// Disposes the workbook and collects it outside the measured region, as
    /// <c>CellStylingBenchmarks</c> does, so one iteration's garbage is not charged to the next.
    /// </summary>
    [IterationCleanup]
    public void Cleanup()
    {
        _workbook.Dispose();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
    }

    /// <summary>First write over unstyled cells: the cost both versions pay.</summary>
    [Benchmark(Baseline = true)]
    public void SetBoldOnce()
    {
        _range.Style.Font.Bold = true;
    }

    /// <summary>The same value again, through the range object that set it.</summary>
    [Benchmark]
    public void SetBoldAgainThroughHeldRange()
    {
        _range.Style.Font.Bold = true;
    }

    /// <summary>The fill setter, whose non-cell path now decides the pattern per cell.</summary>
    [Benchmark]
    public void SetBackgroundColourAgainThroughHeldRange()
    {
        _range.Style.Fill.BackgroundColor = XLColor.LightGray;
    }

    /// <summary>The same value again on a worksheet whose record already carries it.</summary>
    [Benchmark]
    public void SetBoldAgainOnWorksheet()
    {
        _worksheet.Style.Font.Bold = true;
    }
}
