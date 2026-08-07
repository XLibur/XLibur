using System;
using System.Diagnostics;
using System.IO;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// What reading one dirty formula cell costs, against reading a hundred and against reading them
/// all.
///
/// Run with: <c>dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile dirtyread</c>
/// </summary>
/// <remarks>
/// Spec 19 area 5 task 5.3, which is spec 04's task 6 — the benchmark that spec has always needed
/// and never had.
/// <para>
/// Spec 04's case is that <c>XLCell.Evaluate</c> tries <c>TryEvaluateSingleCell</c> and, on any
/// <c>GettingDataException</c>, falls back to <c>Recalculate(wb, null)</c>, which builds the whole
/// calculation chain and dependency tree and walks every formula in the workbook — so reading one
/// cell can cost the workbook. Nothing has ever measured whether that path is reached in practice
/// or what it costs when it is.
/// </para>
/// <para>
/// The experiment needs no instrumentation, because the shape of the answer is visible from the
/// outside: <b>if reading one cell costs about what reading all of them costs, the cliff is real and
/// total.</b> If reading one is proportional to one, the fast path holds. Reading a hundred
/// distinguishes a per-read cliff from a once-per-workbook one — if the first read recalculates
/// everything, the ninety-nine that follow are free and the hundred-cell figure sits on top of the
/// one-cell figure rather than a hundred times it.
/// </para>
/// <para>
/// Two workbook shapes, because the fallback is triggered by a formula whose precedent is itself a
/// dirty formula: <see cref="Shape.ValuePrecedents"/> is the load-then-read case every export tool
/// produces, and <see cref="Shape.ChainedPrecedents"/> is the one spec 04 describes.
/// </para>
/// <para>
/// Bytes are exact. Times are one pass each and are reported because the effect being looked for is
/// orders of magnitude, not percentages; nothing subtle should be read from them.
/// </para>
/// </remarks>
public static class DirtyFormulaReadProfile
{
    private const int Rows = 100_000;
    private const int SampleReads = 100;

    private enum Shape
    {
        /// <summary>Every formula reads plain numbers, so no precedent is ever dirty.</summary>
        ValuePrecedents,

        /// <summary>Every formula reads the previous row's formula, so precedents are dirty.</summary>
        ChainedPrecedents,
    }

    public static void Run()
    {
        SixLaborsV1FontBootstrap.Register();

        Console.WriteLine();
        Console.WriteLine($"Reading dirty formula cells from a freshly loaded workbook, {Rows:N0} formulas.");
        Console.WriteLine("Each row is a fresh load, so no read benefits from a previous one.");
        Console.WriteLine();
        Console.WriteLine("  | shape            | cells read | elapsed ms | allocated | per cell read |");
        Console.WriteLine("  |------------------|------------|------------|-----------|---------------|");

        foreach (var shape in new[] { Shape.ValuePrecedents, Shape.ChainedPrecedents })
        {
            var file = Build(shape);

            // Warm the path once so the first measured row does not absorb JIT.
            Measure(file, 1);

            foreach (var reads in new[] { 1, 10, SampleReads, 1_000, Rows })
            {
                var (ms, allocated) = Measure(file, reads);
                Console.WriteLine(
                    $"  | {shape,-16} | {reads,10:N0} | {ms,10:F1} | {allocated / 1048576.0,6:N1} MB | " +
                    $"{allocated / (double)reads / 1024.0,10:N1} KB |");
            }

            Console.WriteLine("  |------------------|------------|------------|-----------|---------------|");
        }

        ChainDepthSweep();

        Console.WriteLine();
        Console.WriteLine("  If the one-cell row costs about what the all-cells row costs, reading a single dirty");
        Console.WriteLine("  formula recalculates the workbook and spec 04's cliff is real. If the hundred-cell row");
        Console.WriteLine("  sits on top of the one-cell row rather than at a hundred times it, the cost is paid");
        Console.WriteLine("  once per workbook rather than once per read.");
    }

    /// <summary>
    /// One deep read at several chain lengths, to put a complexity class on the fallback.
    /// </summary>
    /// <remarks>
    /// Doubling the chain doubles a linear cost and quadruples a quadratic one. The distinction
    /// matters more than the absolute figure: a linear fallback is merely expensive, a quadratic one
    /// stops being usable at a size real workbooks reach.
    /// </remarks>
    private static void ChainDepthSweep()
    {
        Console.WriteLine();
        Console.WriteLine("  one deep read against chain length (chained shape)");
        Console.WriteLine("  | formulas | elapsed ms | allocated | ms / formula | vs previous |");
        Console.WriteLine("  |----------|------------|-----------|--------------|-------------|");

        var previous = 0.0;
        foreach (var rows in new[] { 12_500, 25_000, 50_000, 100_000 })
        {
            var file = BuildChain(rows);
            var (ms, allocated) = MeasureDeepRead(file, rows);
            Console.WriteLine(
                $"  | {rows,8:N0} | {ms,10:F1} | {allocated / 1048576.0,6:N1} MB | {ms / rows,12:F3} | " +
                $"{(previous > 0 ? $"{ms / previous,10:F2}x" : "         -"),11} |");
            previous = ms;
        }

        Console.WriteLine();
        Console.WriteLine("  Doubling the chain: ~2x is linear, ~4x is quadratic.");
    }

    private static (double Ms, long Allocated) MeasureDeepRead(byte[] file, int rows)
    {
        var stream = new MemoryStream(file, writable: false);

        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);

        using var workbook = new XLWorkbook(stream);
        var ws = workbook.Worksheet(1);

        var before = GC.GetTotalAllocatedBytes(precise: true);
        var sw = Stopwatch.StartNew();
        var value = ws.Cell(rows, 6).GetDouble();
        sw.Stop();

        GC.KeepAlive(value);
        return (sw.Elapsed.TotalMilliseconds, GC.GetTotalAllocatedBytes(precise: true) - before);
    }

    private static byte[] BuildChain(int rows)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Data");

        for (var row = 1; row <= rows; row++)
        {
            ws.Cell(row, 1).Value = row;
            ws.Cell(row, 2).Value = row * 2;
            ws.Cell(row, 3).Value = row * 3;
            ws.Cell(row, 4).Value = row * 4;
            ws.Cell(row, 5).Value = row * 5;
            ws.Cell(row, 6).FormulaA1 = row == 1 ? "SUM(A1:E1)" : $"SUM(A{row}:E{row})+F{row - 1}";
        }

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    /// <summary>
    /// Load the file and read <paramref name="reads"/> formula cells, spread across the sheet so a
    /// sample is not confined to one region.
    /// </summary>
    private static (double Ms, long Allocated) Measure(byte[] file, int reads)
    {
        var stream = new MemoryStream(file, writable: false);

        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);

        // The load is outside the measurement: this is the cost of reading, not of parsing.
        using var workbook = new XLWorkbook(stream);
        var ws = workbook.Worksheet(1);

        var before = GC.GetTotalAllocatedBytes(precise: true);
        var sw = Stopwatch.StartNew();

        // Spread the sample so that it always ends on the last row rather than starting on the
        // first. Sampling from row 1 upwards made the single-read case meaningless for the chained
        // shape: row 1 is the only formula in it with no formula precedent, so reading just that one
        // measured the fast path and reported it as the deep case.
        double checksum = 0;
        for (var taken = 0; taken < reads; taken++)
        {
            var row = (int)((long)Rows * (taken + 1) / reads);
            checksum += ws.Cell(row, 6).GetDouble();
        }

        sw.Stop();
        var allocated = GC.GetTotalAllocatedBytes(precise: true) - before;

        GC.KeepAlive(checksum);
        return (sw.Elapsed.TotalMilliseconds, allocated);
    }

    private static byte[] Build(Shape shape)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Data");

        for (var row = 1; row <= Rows; row++)
        {
            ws.Cell(row, 1).Value = row;
            ws.Cell(row, 2).Value = row * 2;
            ws.Cell(row, 3).Value = row * 3;
            ws.Cell(row, 4).Value = row * 4;
            ws.Cell(row, 5).Value = row * 5;

            ws.Cell(row, 6).FormulaA1 = shape == Shape.ValuePrecedents || row == 1
                ? $"SUM(A{row}:E{row})"
                : $"SUM(A{row}:E{row})+F{row - 1}";
        }

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }
}
