using System;
using System.IO;
using System.IO.Compression;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Excel.Streaming;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// The forward-only writer on the same 50K x 3 workload as
/// <see cref="XLiburWorkbookBenchmarks.CreateAndSave"/> and
/// <see cref="OpenXmlWorkbookBenchmarks.CreateAndSave"/>, so the three are directly comparable
/// in the joined summary: the in-memory model, the streaming writer, and the raw OpenXML SDK
/// floor.
/// </summary>
/// <remarks>
/// Peak memory is what this API is for, and 50K rows is far too small to show it - a workload
/// that fits comfortably in memory cannot demonstrate bounded memory. That measurement lives in
/// <see cref="StreamingMemoryProfile"/>, at a million rows.
/// </remarks>
[MemoryDiagnoser]
[Config(typeof(JoinSummaryConfig))]
public class StreamingWriteBenchmarks
{
    private const int RowCount = 50_000;

    private BenchmarkData _data = null;
    private string[] _strings = null;
    private double[] _numbers = null;
    private DateTime[] _dates = null;

    [GlobalSetup]
    public void Setup()
    {
        SixLaborsV1FontBootstrap.Register();
        _data = BenchmarkData.Create(RowCount);
        _strings = _data.Strings;
        _numbers = _data.Numbers;
        _dates = _data.Dates;
    }

    [Benchmark(Baseline = true)]
    public void StreamingWrite() => Write(new XLStreamingOptions());

    [Benchmark]
    public void StreamingWriteInlineStrings() =>
        Write(new XLStreamingOptions { StringStorage = XLStreamingStringStorage.Inline });

    [Benchmark]
    public void StreamingWriteFastestCompression() =>
        Write(new XLStreamingOptions { CompressionLevel = CompressionLevel.Fastest });

    private void Write(XLStreamingOptions options)
    {
        using var stream = new MemoryStream();
        using var workbook = XLStreamingWorkbook.Create(stream, options);

        var sheet = workbook.AddWorksheet("Data");
        sheet.AppendRow("Name", "Amount", "Date");

        for (var i = 0; i < RowCount; i++)
            sheet.AppendRow(_strings[i], _numbers[i], _dates[i]);

        using (var row = sheet.AddRow())
        {
            row.Cell("Total");
            row.Formula($"SUM(B2:B{RowCount + 1})");
        }

        workbook.Finish();
    }
}

/// <summary>
/// The acceptance measurement for the streaming writer: a million rows of ten columns, written
/// to disk, reporting peak managed heap alongside what the in-memory model costs for the same
/// data.
/// </summary>
/// <remarks>
/// Run with: <c>dotnet run -c Release -- profile streaming [rowCount]</c>.
/// Not a BenchmarkDotNet benchmark - peak heap over one long run is the quantity of interest,
/// not a distribution over many short ones, and running the in-memory comparison under BDN
/// would mean holding a multi-gigabyte workbook through the warmup iterations too.
/// </remarks>
public static class StreamingMemoryProfile
{
    private const int DefaultRowCount = 1_000_000;
    private const int ColumnCount = 10;

    public static void Run(string[] args)
    {
        SixLaborsV1FontBootstrap.Register();

        var rowCount = args.Length > 2 && int.TryParse(args[2], out var parsed) ? parsed : DefaultRowCount;
        Console.WriteLine($"Streaming write: {rowCount:N0} rows x {ColumnCount} cols");
        Console.WriteLine();

        MeasureStreaming(rowCount, new XLStreamingOptions(), "XLStreamingWorkbook (shared strings)");

        // Every row here carries a distinct string, which is the worst case for the shared
        // string table and the one case where a streaming write is not flat in memory. Inline
        // storage is the documented escape hatch; measuring both makes the trade concrete.
        MeasureStreaming(rowCount,
            new XLStreamingOptions { StringStorage = XLStreamingStringStorage.Inline },
            "XLStreamingWorkbook (inline strings)");

        // The in-memory comparison is run at a tenth of the size: XLWorkbook needs roughly a
        // gigabyte per 400K rows at this width, so the full run would page or die on most
        // machines - which is the entire point of the exercise.
        var comparisonRows = Math.Max(1, rowCount / 10);
        Console.WriteLine();
        Console.WriteLine($"For comparison, XLWorkbook at {comparisonRows:N0} rows (a tenth of the size):");
        MeasureInMemory(comparisonRows);
    }

    private static void MeasureStreaming(int rowCount, XLStreamingOptions options, string label)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlibur-streaming-profile-{Guid.NewGuid():N}.xlsx");
        try
        {
            ForceGc();
            var baseline = GC.GetTotalMemory(forceFullCollection: true);
            var start = DateTime.UtcNow;

            long peak;
            using (var workbook = XLStreamingWorkbook.Create(path, options))
            {
                var sheet = workbook.AddWorksheet("Data");
                var values = new XLCellValue[ColumnCount];

                peak = baseline;
                for (var r = 0; r < rowCount; r++)
                {
                    values[0] = $"Item {r}";
                    for (var c = 1; c < ColumnCount; c++)
                        values[c] = r * ColumnCount + c;

                    sheet.AppendRow(values, null);

                    if (r % 100_000 == 0)
                        peak = Math.Max(peak, GC.GetTotalMemory(forceFullCollection: false));
                }

                peak = Math.Max(peak, GC.GetTotalMemory(forceFullCollection: false));

                // Finish() serialises the shared string table, which under SharedStrings is the
                // largest live structure of the whole write - sampling before it would report a
                // peak that excludes the most expensive moment.
                workbook.Finish();
                peak = Math.Max(peak, GC.GetTotalMemory(forceFullCollection: false));
            }

            var elapsed = DateTime.UtcNow - start;
            var fileSize = new FileInfo(path).Length;

            Report(label, peak - baseline, elapsed, fileSize);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static void MeasureInMemory(int rowCount)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlibur-inmemory-profile-{Guid.NewGuid():N}.xlsx");
        try
        {
            ForceGc();
            var baseline = GC.GetTotalMemory(forceFullCollection: true);
            var start = DateTime.UtcNow;

            long peak;
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.AddWorksheet("Data");
                for (var r = 0; r < rowCount; r++)
                {
                    sheet.Cell(r + 1, 1).Value = $"Item {r}";
                    for (var c = 1; c < ColumnCount; c++)
                        sheet.Cell(r + 1, c + 1).Value = r * ColumnCount + c;
                }

                peak = GC.GetTotalMemory(forceFullCollection: false);
                workbook.SaveAs(path);
                peak = Math.Max(peak, GC.GetTotalMemory(forceFullCollection: false));
            }

            var elapsed = DateTime.UtcNow - start;
            var fileSize = new FileInfo(path).Length;

            Report("XLWorkbook", peak - baseline, elapsed, fileSize);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static void Report(string label, long peakBytes, TimeSpan elapsed, long fileSize)
    {
        Console.WriteLine($"  {label}");
        Console.WriteLine($"    peak managed heap : {peakBytes / 1024.0 / 1024.0,8:F1} MB");
        Console.WriteLine($"    elapsed           : {elapsed.TotalSeconds,8:F2} s");
        Console.WriteLine($"    file size         : {fileSize / 1024.0 / 1024.0,8:F1} MB");
    }

    private static void ForceGc()
    {
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
    }
}
