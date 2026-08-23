using System;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// What <see cref="IXLStyle.Batch"/> buys over the same six assignments made directly, on a cell.
///
/// Run with:
/// dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- --filter '*BatchStyling*'
/// </summary>
/// <remarks>
/// Spec 23 task 5. <c>Batch</c> exists for performance: a batch of N property assignments should
/// cost one style resolution and one style-slice write rather than N of each. Spec 23 replaced the
/// parallel <c>XLDeferred*</c> object graph that achieved that with a pending key on the one style
/// facade, so the claim has to be re-measured against the same shape it was built for.
/// <para>
/// <see cref="DirectPerCell"/> is the thing <c>Batch</c> is meant to beat, not a baseline in the
/// sense of "unchanged code" — the number that matters across the merge-base is
/// <see cref="BatchPerCell"/>, and the ratio between the two is what says batching still pays.
/// </para>
/// </remarks>
[MemoryDiagnoser]
public class BatchStylingBenchmarks
{
    private const int Rows = 50_000;

    private const string NumberFormat = "#,##0.00";

    private XLWorkbook _workbook = null!;
    private IXLWorksheet _worksheet = null!;

    [GlobalSetup]
    public void GlobalSetup() => SixLaborsV1FontBootstrap.Register();

    [IterationSetup]
    public void IterationSetup()
    {
        _workbook = new XLWorkbook();
        _worksheet = _workbook.AddWorksheet("Data");
    }

    /// <summary>
    /// Disposes the workbook and collects it here, outside the measured region — see the same
    /// cleanup on <see cref="CellStylingBenchmarks"/> for why the previous iteration's garbage
    /// otherwise lands on whichever variant runs next.
    /// </summary>
    [IterationCleanup]
    public void IterationCleanup()
    {
        _workbook.Dispose();

        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
    }

    [Benchmark(Baseline = true)]
    public void DirectPerCell()
    {
        for (var r = 0; r < Rows; r++)
        {
            var style = _worksheet.Cell(r + 1, 1).Style;
            style.Font.Bold = true;
            style.Font.FontSize = 12;
            style.Font.FontColor = XLColor.Red;
            style.Fill.BackgroundColor = XLColor.Green;
            style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            style.NumberFormat.Format = NumberFormat;
        }
    }

    [Benchmark]
    public void BatchPerCell()
    {
        for (var r = 0; r < Rows; r++)
        {
            _worksheet.Cell(r + 1, 1).Style.Batch(s =>
            {
                s.Font.Bold = true;
                s.Font.FontSize = 12;
                s.Font.FontColor = XLColor.Red;
                s.Fill.BackgroundColor = XLColor.Green;
                s.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                s.NumberFormat.Format = NumberFormat;
            });
        }
    }
}
