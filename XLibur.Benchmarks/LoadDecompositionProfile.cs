using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// Where the load path's allocation goes: how much the model keeps, how much is garbage, and which
/// column shapes are responsible.
///
/// Run with: <c>dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile loadalloc</c>
/// </summary>
/// <remarks>
/// Spec 19 area 4, task 4.1. <c>LoadWorkbook</c> allocates 334.54 MB parsing 3.75 M cells — 93.5 B
/// each — and after spec 02 removed the transient per-cell strings nobody has said what the rest is.
/// The spec's rule is that no other task in this area starts until this table exists, because two
/// predecessors (05 and 18) published attributions that did not survive re-measurement.
/// <para>
/// The first split is retained against transient, because it decides whether there is anything to
/// win at all. Allocation the model still holds after a forced gen2 collection is storage: the value
/// slice, the shared strings, the formula objects. Allocation that has already died is garbage, and
/// only garbage is reachable by a code change that does not alter what a workbook costs to hold.
/// </para>
/// <para>
/// The second split is by column shape, measured by ablation rather than by instrumenting the
/// loader: each variant differs from the baseline in exactly one dimension, so a difference prices
/// that dimension. Crossing rather than guessing is what spec 18 task 3 had to do to separate sheet
/// geometry from string uniqueness, after two plausible single-cause explanations both turned out to
/// be half right.
/// </para>
/// <para>
/// 100,000 rows rather than <c>XLiburReadBenchmarks</c>' 250,000: the per-cell costs are the same
/// and eight fixtures have to be built. Figures are reported per cell so they compare directly.
/// Bytes are exact (<c>GC.GetTotalAllocatedBytes(precise: true)</c>); the elapsed column is a single
/// pass and must not be used to claim a change moved.
/// </para>
/// </remarks>
public static class LoadDecompositionProfile
{
    private const int Rows = 100_000;

    public static void Run()
    {
        SixLaborsV1FontBootstrap.Register();

        var variants = new (string Label, Func<byte[]> Build)[]
        {
            ("baseline (as XLiburReadBenchmarks: 3 uniq str, 5 num, 3 date, 3 str, 1 formula)", () => Build(Shape.Baseline)),
            ("no formula column (col 15 numeric instead)", () => Build(Shape.NoFormula)),
            ("formula column, one repeated text (SUM($D$1:$H$1))", () => Build(Shape.SharedFormulaText)),
            ("no unique strings (cols 1/13/14 repeat 10 values)", () => Build(Shape.RepeatedStrings)),
            ("no dates (cols 9-11 numeric)", () => Build(Shape.NoDates)),
            ("all 15 columns numeric", () => Build(Shape.AllNumeric)),
            ("1 numeric column (per-row cost)", () => Build(Shape.SingleNumeric)),
        };

        Console.WriteLine();
        Console.WriteLine($"Load allocation, {Rows:N0} rows. Bytes exact; elapsed is one pass.");
        Console.WriteLine();
        Console.WriteLine("  | variant                                                          | cells     | allocated | retained | transient | B/cell | ret B/cell | file KB | ms   |");
        Console.WriteLine("  |------------------------------------------------------------------|-----------|-----------|----------|-----------|--------|------------|---------|------|");

        var results = new List<(string Label, long Allocated, long Retained, long Cells)>();

        foreach (var (label, build) in variants)
        {
            var bytes = build();
            var (allocated, retained, cells, ms) = MeasureLoad(bytes);
            results.Add((label, allocated, retained, cells));

            Console.WriteLine(
                $"  | {label,-66} | {cells,9:N0} | {allocated / 1048576.0,7:N1} MB | " +
                $"{retained / 1048576.0,6:N1} MB | {(allocated - retained) / 1048576.0,7:N1} MB | " +
                $"{(double)allocated / cells,6:F1} | {(double)retained / cells,10:F1} | " +
                $"{bytes.Length / 1024.0,7:N0} | {ms,4:F0} |");
        }

        Report(results);
    }

    private static void Report(List<(string Label, long Allocated, long Retained, long Cells)> r)
    {
        var baseline = r[0];
        Console.WriteLine();
        Console.WriteLine("  What each dimension costs, as the baseline minus the variant that removes it.");
        Console.WriteLine("  Cell counts are equal except for the last, so these are like-for-like.");
        Console.WriteLine();
        Console.WriteLine("  | dimension removed        | allocated saved | retained saved | per cell of that kind |");
        Console.WriteLine("  |--------------------------|-----------------|----------------|-----------------------|");

        Diff("formula column (1 of 15)", r[1], 1);
        Diff("distinct formula text", r[2], 1);
        Diff("unique strings (3 of 15)", r[3], 3);
        Diff("dates (3 of 15)", r[4], 3);
        Diff("everything but numbers", r[5], 9);

        void Diff(string label, (string Label, long Allocated, long Retained, long Cells) variant, int columns)
        {
            var alloc = baseline.Allocated - variant.Allocated;
            var ret = baseline.Retained - variant.Retained;
            var affected = (double)Rows * columns;
            Console.WriteLine(
                $"  | {label,-24} | {alloc / 1048576.0,12:N1} MB | {ret / 1048576.0,11:N1} MB | " +
                $"{alloc / affected,10:F1} B alloc, {ret / affected,5:F1} B kept |");
        }

        Console.WriteLine();
        Console.WriteLine("  Retained is measured with the workbook alive after a forced gen2 collection, so it is");
        Console.WriteLine("  what holding the workbook costs. Transient is allocation that had already died: the");
        Console.WriteLine("  only part a loader change can remove without changing what a workbook costs to hold.");
    }

    private static (long Allocated, long Retained, long Cells, double Ms) MeasureLoad(byte[] fileBytes)
    {
        // The stream wraps the existing array rather than copying it, and is created before the
        // baseline reading so it is not counted as load allocation.
        var stream = new MemoryStream(fileBytes, writable: false);

        Collect();
        var retainedBefore = GC.GetTotalMemory(forceFullCollection: true);
        var allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);

        var sw = Stopwatch.StartNew();
        var workbook = new XLWorkbook(stream);
        sw.Stop();

        var allocatedAfter = GC.GetTotalAllocatedBytes(precise: true);

        Collect();
        var retainedAfter = GC.GetTotalMemory(forceFullCollection: true);

        var cells = workbook.Worksheets.Sum(ws => (long)ws.CellsUsed().Count());

        GC.KeepAlive(workbook);
        workbook.Dispose();
        GC.KeepAlive(stream);

        return (allocatedAfter - allocatedBefore, retainedAfter - retainedBefore, cells, sw.Elapsed.TotalMilliseconds);
    }

    private static void Collect()
    {
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
    }

    private enum Shape
    {
        Baseline,
        NoFormula,
        SharedFormulaText,
        RepeatedStrings,
        NoDates,
        AllNumeric,
        SingleNumeric,
    }

    private static byte[] Build(Shape shape)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Data");

#pragma warning disable S2245 // Deterministic seed for a reproducible fixture.
        var random = new Random(42);
#pragma warning restore S2245
        var baseDate = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        string[] regions = { "North", "South", "East", "West", "Central" };
        string[] statuses = { "Active", "Pending", "Closed", "Review", "Draft" };

        for (var i = 0; i < Rows; i++)
        {
            var row = i + 1;
            var seed = random.Next(10000);

            if (shape == Shape.SingleNumeric)
            {
                ws.Cell(row, 1).Value = seed;
                continue;
            }

            var numeric = shape == Shape.AllNumeric;
            var uniqueStrings = shape != Shape.RepeatedStrings && !numeric;

            // Cols 1-3: strings, of which col 1 is the unique one in the baseline.
            if (numeric)
            {
                ws.Cell(row, 1).Value = seed;
                ws.Cell(row, 2).Value = i % 12;
                ws.Cell(row, 3).Value = i % regions.Length;
            }
            else
            {
                ws.Cell(row, 1).Value = uniqueStrings ? $"Item {i}-{seed}" : $"Item {i % 10}";
                ws.Cell(row, 2).Value = $"Cat-{i % 12}";
                ws.Cell(row, 3).Value = regions[i % regions.Length];
            }

            // Cols 4-8: numbers in every shape.
            ws.Cell(row, 4).Value = Math.Round(random.NextDouble() * 10000, 2);
            ws.Cell(row, 5).Value = random.Next(1, 5000);
            ws.Cell(row, 6).Value = Math.Round(random.NextDouble(), 4);
            ws.Cell(row, 7).Value = Math.Round(random.NextDouble() * 1000, 2);
            ws.Cell(row, 8).Value = random.Next(0, 100);

            // Cols 9-11: dates, unless the shape removes them.
            if (shape is Shape.NoDates or Shape.AllNumeric)
            {
                ws.Cell(row, 9).Value = random.Next(0, 2000);
                ws.Cell(row, 10).Value = random.Next(0, 2000);
                ws.Cell(row, 11).Value = random.Next(0, 2000);
            }
            else
            {
                ws.Cell(row, 9).Value = baseDate.AddDays(random.Next(0, 2000));
                ws.Cell(row, 10).Value = baseDate.AddDays(random.Next(0, 2000));
                ws.Cell(row, 11).Value = baseDate.AddDays(random.Next(0, 2000));
            }

            // Cols 12-14: more strings, two of them unique in the baseline.
            if (numeric)
            {
                ws.Cell(row, 12).Value = i % statuses.Length;
                ws.Cell(row, 13).Value = row;
                ws.Cell(row, 14).Value = seed;
            }
            else
            {
                ws.Cell(row, 12).Value = statuses[i % statuses.Length];
                ws.Cell(row, 13).Value = uniqueStrings ? $"Note for row {row} with seed {seed}" : "Note";
                ws.Cell(row, 14).Value = uniqueStrings ? $"CODE-{seed:D5}" : "CODE-00000";
            }

            // Col 15: the formula, unless the shape removes it.
            if (shape is Shape.NoFormula or Shape.AllNumeric)
                ws.Cell(row, 15).Value = seed;
            else if (shape == Shape.SharedFormulaText)
                ws.Cell(row, 15).FormulaA1 = "SUM($D$1:$H$1)";
            else
                ws.Cell(row, 15).FormulaA1 = $"SUM(D{row}:H{row})";
        }

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }
}
