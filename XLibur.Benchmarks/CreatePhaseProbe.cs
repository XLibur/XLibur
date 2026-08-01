using System;
using System.Diagnostics;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// Per-operation allocation cost of the public cell-building API, on a worksheet with no merged
/// ranges, tables or formulas.
///
/// Run with: dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile create
///
/// <see cref="SaveAllocationProfile"/> answers "how much does the whole workbook cost"; this answers
/// "which API call is paying for it". The two together are what Spec 11's acceptance criteria are
/// stated against.
/// </summary>
public static class CreatePhaseProbe
{
    private const int Rows = 50_000;
    private const int Cols = 10;
    private const int Ops = Rows * Cols;

    public static void Run()
    {
        SixLaborsV1FontBootstrap.Register();

        // Warm up JIT and the process-wide style/colour repositories so one-time costs do not
        // land on whichever probe happens to run first.
        foreach (var probe in Probes())
            probe.Action();

        Console.WriteLine();
        Console.WriteLine($"Per-operation cost over {Ops:N0} operations ({Rows:N0} rows x {Cols} cols)");
        Console.WriteLine("| Probe                                   |    Total |  Bytes/op |   ns/op |");
        Console.WriteLine("|-----------------------------------------|----------|-----------|---------|");

        foreach (var probe in Probes())
        {
            ForceGC();

            var before = GC.GetTotalAllocatedBytes(precise: true);
            var watch = Stopwatch.StartNew();
            probe.Action();
            watch.Stop();
            var bytes = GC.GetTotalAllocatedBytes(precise: true) - before;

            Console.WriteLine(
                $"| {probe.Name,-39} | {bytes / 1048576.0,5:F1} MB | {bytes / (double)Ops,9:F1} | {watch.Elapsed.TotalNanoseconds / Ops,7:F1} |");
        }

        Console.WriteLine();
        Console.WriteLine("Bytes/op is exact. ns/op is single-shot — use BenchmarkDotNet for time claims.");
        Console.WriteLine("The bulk-style row is a total: subtract the ws.SetCellValue row, which populates");
        Console.WriteLine("identically, to get the styling-only cost per cell.");
    }

    private static (string Name, Action Action)[] Probes() =>
    [
        ("ws.Cell(r,c) discarded", CellWrapperOnly),
        ("ws.Cell(r,c).Value = double", CellValueDouble),
        ("ws.SetCellValue(r,c, double)", SetCellValueDouble),
        ("ws.Cell(r,c).Value = string (shared)", CellValueString),
        ("ws.Cell(r,c).Value = DateTime", CellValueDateTime),
        ("...Value = double, sheet has 1 merge", CellValueDoubleWithMerge),
        ("ws.Cell(r,c).Style discarded", StyleWrapperOnly),
        ("ws.Cell(r,c).Style + 1 font mutation", StyleOneMutation),
        ("ws.Cell(r,c).Style + 4 mutations", StyleFourMutations),
        ("ws.Range(all).Style.Bold + populate", RangeBulkStyle),
    ];

    /// <summary>
    /// Bulk styling goes through <c>XLStylizedBase.ModifyStyle</c>. Measured over the same cell
    /// count as the per-cell probes so the two are directly comparable — note the styling itself is
    /// one statement, not 500,000.
    /// </summary>
    /// <remarks>
    /// The range has to be populated first, so this row is a <em>total</em>: it includes the same
    /// work as the <c>ws.SetCellValue(r,c, double)</c> probe, which populates identically and is
    /// therefore the matched baseline. Styling-only cost is this row minus that one.
    /// </remarks>
    private static void RangeBulkStyle()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("s");

        // Identical to SetCellValueDouble, so subtracting that probe isolates the styling.
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
                ws.SetCellValue(r, c, r * 1.5);
        }

        ws.Range(1, 1, Rows, Cols).Style.Font.Bold = true;
    }

    private static void CellWrapperOnly()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("s");
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
                _ = ws.Cell(r, c);
        }
    }

    private static void CellValueDouble()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("s");
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
                ws.Cell(r, c).Value = r * 1.5;
        }
    }

    private static void SetCellValueDouble()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("s");
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
                ws.SetCellValue(r, c, r * 1.5);
        }
    }

    private static void CellValueString()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("s");
        // A fixed pool, so the shared-string table does its work without the probe itself
        // allocating a fresh string per cell.
        string[] pool = ["Active", "Pending", "Closed", "Review", "Draft"];
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
                ws.Cell(r, c).Value = pool[(r + c) % pool.Length];
        }
    }

    private static void CellValueDateTime()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("s");
        var baseDate = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
                ws.Cell(r, c).Value = baseDate.AddDays((r + c) % 1500);
        }
    }

    /// <summary>
    /// A single merged title row is an extremely common sheet layout, and it takes every
    /// subsequent cell write off the "no merged ranges" fast path — so the cost of the
    /// merged-range test itself, not just the cost of skipping it, has to stay low.
    /// </summary>
    private static void CellValueDoubleWithMerge()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("s");
        ws.Range(1, 1, 1, 4).Merge();

        for (var r = 2; r <= Rows + 1; r++)
        {
            for (var c = 1; c <= Cols; c++)
                ws.Cell(r, c).Value = r * 1.5;
        }
    }

    private static void StyleWrapperOnly()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("s");
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
                _ = ws.Cell(r, c).Style;
        }
    }

    private static void StyleOneMutation()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("s");
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
                ws.Cell(r, c).Style.Font.Bold = true;
        }
    }

    private static void StyleFourMutations()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("s");
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
            {
                var s = ws.Cell(r, c).Style;
                s.Font.Bold = true;
                s.Font.FontColor = XLColor.DarkRed;
                s.Fill.BackgroundColor = XLColor.LightBlue;
                s.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
        }
    }

    // ReSharper disable once InconsistentNaming
    private static void ForceGC()
    {
#pragma warning disable S1215 // Intentionally forcing GC to isolate probes
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, true, true);
#pragma warning restore S1215
    }
}
