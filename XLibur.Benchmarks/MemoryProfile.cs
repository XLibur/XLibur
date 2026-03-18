using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using JetBrains.Profiler.SelfApi;
using XLibur.Excel;

namespace XLibur.Benchmarks;

/// <summary>
/// Standalone profiling harness for dotMemory analysis.
/// Run with: dotnet run -c Release -- profile [load|read|both]
///
/// Modes:
///   load  - Profile workbook loading only (XLWorkbook constructor)
///   read  - Profile loading & reading all cells as strings
///   both  - Take snapshots at each phase (default)
///
/// Output: .dmw workspace files in C:\profiles\
/// </summary>
public static class MemoryProfile
{
    private const int RowCount = 250_000;
    private const int ColCount = 15;

    public static void Run(string[] args)
    {
        var mode = args.Length > 1 ? args[1].ToLowerInvariant() : "both";
        var outputDir = args.Length > 2 ? args[2] : @"C:\profiles";

        Directory.CreateDirectory(outputDir);

        Console.WriteLine($"Mode: {mode}");
        Console.WriteLine($"Output: {outputDir}");
        Console.WriteLine($"Generating test file ({RowCount:N0} rows x {ColCount} cols)...");

        var fileBytes = GenerateTestFile();
        Console.WriteLine($"Test file size: {fileBytes.Length / 1024.0 / 1024.0:F1} MB");

        // Force GC to get a clean baseline
        ForceGC();

        Console.WriteLine("Initializing dotMemory...");
        DotMemory.Init();

        var config = new DotMemory.Config()
            .SaveToDir(outputDir);

        Console.WriteLine("Attaching dotMemory...");
        DotMemory.Attach(config);

        try
        {
            ForceGC();
            DotMemory.GetSnapshot("Baseline");
            Console.WriteLine("[Snapshot] Baseline captured");

            switch (mode)
            {
                case "load":
                    ProfileLoad(fileBytes);
                    break;
                case "read":
                    ProfileLoadAndRead(fileBytes);
                    break;
                default:
                    ProfileBoth(fileBytes);
                    break;
            }
        }
        finally
        {
            Console.WriteLine("Detaching dotMemory...");
            DotMemory.Detach();
            Console.WriteLine("Done. Open .dmw file from {0}", outputDir);
        }
    }

    private static void ProfileLoad(byte[] fileBytes)
    {
        Console.WriteLine("Loading workbook...");
        var sw = Stopwatch.StartNew();

        var workbook = new XLWorkbook(new MemoryStream(fileBytes));
        sw.Stop();
        Console.WriteLine($"Load completed in {sw.Elapsed.TotalSeconds:F1}s");

        ForceGC();
        DotMemory.GetSnapshot("After Load");
        Console.WriteLine("[Snapshot] After Load captured");

        workbook.Dispose();
        ForceGC();
        DotMemory.GetSnapshot("After Dispose");
        Console.WriteLine("[Snapshot] After Dispose captured");
    }

    private static void ProfileLoadAndRead(byte[] fileBytes)
    {
        Console.WriteLine("Loading workbook...");
        var sw = Stopwatch.StartNew();

        using var workbook = new XLWorkbook(new MemoryStream(fileBytes));
        sw.Stop();
        Console.WriteLine($"Load completed in {sw.Elapsed.TotalSeconds:F1}s");

        Console.WriteLine("Reading all cells...");
        sw.Restart();

        var ws = workbook.Worksheets.First();
        var count = 0;
        for (var row = 1; row <= RowCount; row++)
        {
            for (var col = 1; col <= ColCount; col++)
            {
                _ = ws.Cell(row, col).GetValue<string>();
                count++;
            }
        }

        sw.Stop();
        Console.WriteLine($"Read {count:N0} cells in {sw.Elapsed.TotalSeconds:F1}s");

        ForceGC();
        DotMemory.GetSnapshot("After Load+Read");
        Console.WriteLine("[Snapshot] After Load+Read captured");
    }

    private static void ProfileBoth(byte[] fileBytes)
    {
        Console.WriteLine("=== Phase 1: Load ===");
        var sw = Stopwatch.StartNew();

        var workbook = new XLWorkbook(new MemoryStream(fileBytes));
        sw.Stop();
        Console.WriteLine($"Load completed in {sw.Elapsed.TotalSeconds:F1}s");

        ForceGC();
        DotMemory.GetSnapshot("After Load");
        Console.WriteLine("[Snapshot] After Load captured");

        Console.WriteLine("=== Phase 2: Read All Cells ===");
        sw.Restart();

        var ws = workbook.Worksheets.First();
        var count = 0;
        for (var row = 1; row <= RowCount; row++)
        {
            for (var col = 1; col <= ColCount; col++)
            {
                _ = ws.Cell(row, col).GetValue<string>();
                count++;
            }
        }

        sw.Stop();
        Console.WriteLine($"Read {count:N0} cells in {sw.Elapsed.TotalSeconds:F1}s");

        ForceGC();
        DotMemory.GetSnapshot("After Read");
        Console.WriteLine("[Snapshot] After Read captured");

        Console.WriteLine("=== Phase 3: Dispose ===");
        workbook.Dispose();

        ForceGC();
        DotMemory.GetSnapshot("After Dispose");
        Console.WriteLine("[Snapshot] After Dispose captured");
    }

    private static byte[] GenerateTestFile()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Data");

#pragma warning disable S2245 // Deterministic seed for reproducible benchmarks
        var random = new Random(42);
#pragma warning restore S2245
        var baseDate = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        string[] regions = { "North", "South", "East", "West", "Central" };
        string[] statuses = { "Active", "Pending", "Closed", "Review", "Draft" };

        for (var i = 0; i < RowCount; i++)
        {
            var row = i + 1;
            var seed = random.Next(10000);

            ws.Cell(row, 1).Value = $"Item {i}-{seed}";
            ws.Cell(row, 2).Value = $"Cat-{i % 12}";
            ws.Cell(row, 3).Value = regions[i % regions.Length];

            ws.Cell(row, 4).Value = Math.Round(random.NextDouble() * 10000, 2);
            ws.Cell(row, 5).Value = random.Next(1, 5000);
            ws.Cell(row, 6).Value = Math.Round(random.NextDouble(), 4);
            ws.Cell(row, 7).Value = Math.Round(random.NextDouble() * 1000, 2);
            ws.Cell(row, 8).Value = random.Next(0, 100);

            ws.Cell(row, 9).Value = baseDate.AddDays(random.Next(0, 2000));
            ws.Cell(row, 10).Value = baseDate.AddDays(random.Next(0, 2000));
            ws.Cell(row, 11).Value = baseDate.AddDays(random.Next(0, 2000));

            ws.Cell(row, 12).Value = statuses[i % statuses.Length];
            ws.Cell(row, 13).Value = $"Note for row {row} with seed {seed}";
            ws.Cell(row, 14).Value = $"CODE-{seed:D5}";

            ws.Cell(row, 15).FormulaA1 = $"SUM(D{row}:H{row})";
        }

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    // ReSharper disable once InconsistentNaming
    private static void ForceGC()
    {
#pragma warning disable S1215 // Intentionally forcing GC for accurate memory profiling
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, true, true);
#pragma warning restore S1215
    }
}
