using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using XLibur.Excel;
using XLibur.Excel.Streaming;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// What the zip costs the save path, and what the caller gets for it.
///
/// Run with: <c>dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile compression</c>
/// </summary>
/// <remarks>
/// Spec 19 area 3, tasks 3.1 and 3.2.
/// <para>
/// Task 3.1 needs both halves of the trade. <c>CreateAndSaveFastestCompression</c> established that
/// <see cref="CompressionLevel.Fastest"/> is worth 30% of the create-and-save benchmark, but a
/// default cannot be argued from time alone - the output gets bigger, and by how much depends on the
/// content. Three shapes are measured because one would not generalise: a narrow numeric grid, a
/// wide heavily-styled one, and a structurally rich template whose sheets are nearly empty.
/// </para>
/// <para>
/// Task 3.2 asks how much of the gap between the in-memory model's save and the streaming writer is
/// packaging rather than the model. Deflate is common to both, so the comparison is made at
/// <see cref="CompressionLevel.NoCompression"/>, where whatever remains is the packaging layer plus
/// the difference in how each reaches the data. The model's save is timed on a workbook built
/// outside the clock, so it is the save alone; the streaming figure is necessarily build and write
/// together, and the model's own build time is reported so the two can be put on the same footing.
/// </para>
/// <para>
/// Times are medians of several passes. Per this spec's own measurement protocol they locate cost
/// and must not be used to claim a change moved - <c>XLiburWorkbookBenchmarks</c> is for that. The
/// byte counts are exact.
/// </para>
/// </remarks>
public static class CompressionProfile
{
    private const int RowCount = 50_000;
    private const int Passes = 5;

    private static readonly CompressionLevel[] Levels =
    [
        CompressionLevel.NoCompression,
        CompressionLevel.Fastest,
        CompressionLevel.Optimal,
        CompressionLevel.SmallestSize,
    ];

    public static void Run()
    {
        SixLaborsV1FontBootstrap.Register();
        var data = BenchmarkData.Create(RowCount);

        Console.WriteLine();
        Console.WriteLine($"Compression trade-off, medians of {Passes} passes. Bytes are exact.");

        MeasureShape("narrow numeric grid (50,000 x 3)", () => BuildGrid(data));
        MeasureShape("styled grid (50,000 x 10, half the rows styled)", () => BuildFormatted(data));
        MeasureShape("template (10 sheets, 20 names, 26 validations)", BuildTemplate);

        ReportLevelHonoured(data);
        StreamingSplit(data);
    }

    /// <summary>
    /// Save at each level, rebuilding the workbook every pass.
    /// </summary>
    /// <remarks>
    /// The rebuild is not tidiness, it is the only way to measure this at all: after one SaveAs the
    /// workbook's load source is the stream it just wrote, and the next save copies that package and
    /// patches it rather than creating one, so the level is not applied to parts that already exist.
    /// Re-saving one workbook produced byte-identical output at all four levels, which is what
    /// exposed it - see <see cref="ReportLevelHonoured"/>.
    /// </remarks>
    private static void MeasureShape(string label, Func<XLWorkbook> build)
    {
        // Warm the path so the first level measured does not absorb JIT.
        using (var warm = build())
            warm.SaveAs(new MemoryStream(), new SaveOptions { CompressionLevel = CompressionLevel.Optimal });

        Console.WriteLine();
        Console.WriteLine($"  {label}");
        Console.WriteLine("  | level           | save ms | output KB | vs Optimal time | vs Optimal size |");
        Console.WriteLine("  |-----------------|---------|-----------|-----------------|-----------------|");

        // Measured once per level and kept: measuring again to compute the ratios would both double
        // the work and quote a different pass than the one printed.
        var results = Levels.Select(level => (Level: level, Result: SaveAtLevel(build, level))).ToArray();
        var optimal = results.Single(r => r.Level == CompressionLevel.Optimal).Result;

        foreach (var (level, (ms, bytes)) in results)
        {
            Console.WriteLine(
                $"  | {level,-15} | {ms,7:F1} | {bytes / 1024.0,9:N0} | " +
                $"{ms / optimal.Ms,14:F2}x | {(double)bytes / optimal.Bytes,14:F2}x |");
        }
    }

    private static (double Ms, long Bytes) SaveAtLevel(Func<XLWorkbook> build, CompressionLevel level)
    {
        var options = new SaveOptions { CompressionLevel = level };
        var times = new double[Passes];
        long bytes = 0;

        for (var i = 0; i < Passes; i++)
        {
            using var workbook = build();
            var output = new MemoryStream();
            var sw = Stopwatch.StartNew();
            workbook.SaveAs(output, options);
            sw.Stop();
            times[i] = sw.Elapsed.TotalMilliseconds;
            bytes = output.Length;
        }

        return (Median(times), bytes);
    }

    /// <summary>
    /// Whether <see cref="SaveOptions.CompressionLevel"/> reaches the output at all, for the three
    /// situations a caller can be in: a new workbook saved once, the same workbook saved again, and
    /// a workbook that was loaded from a file.
    /// </summary>
    private static void ReportLevelHonoured(BenchmarkData data)
    {
        Console.WriteLine();
        Console.WriteLine("  is SaveOptions.CompressionLevel honoured? (output KB)");
        Console.WriteLine("  | situation                        | NoCompression | Optimal | honoured |");
        Console.WriteLine("  |----------------------------------|---------------|---------|----------|");

        Row("new workbook, first save", level =>
        {
            using var wb = BuildGrid(data);
            var output = new MemoryStream();
            wb.SaveAs(output, new SaveOptions { CompressionLevel = level });
            return output.Length;
        });

        Row("same workbook, second save", level =>
        {
            using var wb = BuildGrid(data);
            wb.SaveAs(new MemoryStream(), new SaveOptions { CompressionLevel = CompressionLevel.Optimal });
            var output = new MemoryStream();
            wb.SaveAs(output, new SaveOptions { CompressionLevel = level });
            return output.Length;
        });

        var template = TemplateFixture.Build(sheetCount: 10, definedNames: 20, validations: 26, dataRows: 2_000);
        Row("workbook loaded from a stream", level =>
        {
            using var wb = new XLWorkbook(new MemoryStream(template, writable: false));
            var output = new MemoryStream();
            wb.SaveAs(output, new SaveOptions { CompressionLevel = level });
            return output.Length;
        });

        static void Row(string label, Func<CompressionLevel, long> save)
        {
            var none = save(CompressionLevel.NoCompression);
            var optimal = save(CompressionLevel.Optimal);
            var honoured = none > optimal * 1.5 ? "yes" : "NO";
            Console.WriteLine($"  | {label,-32} | {none / 1024.0,13:N0} | {optimal / 1024.0,7:N0} | {honoured,8} |");
        }
    }

    /// <summary>
    /// Task 3.2: the model path against the streaming writer at each level, on identical data.
    /// </summary>
    private static void StreamingSplit(BenchmarkData data)
    {
        Console.WriteLine();
        Console.WriteLine("  model versus streaming writer, same 50,000 x 3 data");
        Console.WriteLine("  | level           | model build+save ms | streaming ms | model save only ms | model KB | streaming KB |");
        Console.WriteLine("  |-----------------|---------------------|--------------|--------------------|----------|--------------|");

        // Warm both paths.
        using (var wb = BuildGrid(data))
            wb.SaveAs(new MemoryStream());
        using (var warm = new MemoryStream())
            StreamingWrite(data, CompressionLevel.Optimal, warm);

        foreach (var level in Levels)
        {
            var buildSave = new double[Passes];
            var saveOnly = new double[Passes];
            var streaming = new double[Passes];
            long modelBytes = 0;
            long streamBytes = 0;

            for (var i = 0; i < Passes; i++)
            {
                var sw = Stopwatch.StartNew();
                using var wb = BuildGrid(data);
                var built = sw.Elapsed.TotalMilliseconds;

                var output = new MemoryStream();
                var saveStart = sw.Elapsed.TotalMilliseconds;
                wb.SaveAs(output, new SaveOptions { CompressionLevel = level });
                sw.Stop();

                buildSave[i] = sw.Elapsed.TotalMilliseconds;
                saveOnly[i] = sw.Elapsed.TotalMilliseconds - saveStart;
                modelBytes = output.Length;
                _ = built;

                using var streamOut = new MemoryStream();
                var sw2 = Stopwatch.StartNew();
                StreamingWrite(data, level, streamOut);
                sw2.Stop();
                streaming[i] = sw2.Elapsed.TotalMilliseconds;
                streamBytes = streamOut.Length;
            }

            Console.WriteLine(
                $"  | {level,-15} | {Median(buildSave),19:F1} | {Median(streaming),12:F1} | " +
                $"{Median(saveOnly),18:F1} | {modelBytes / 1024.0,8:N0} | {streamBytes / 1024.0,12:N0} |");
        }

        Console.WriteLine();
        Console.WriteLine("  At NoCompression deflate is out of both columns, so what is left of the gap");
        Console.WriteLine("  is System.IO.Packaging against the streaming writer's own zip, plus the");
        Console.WriteLine("  difference between building a model and writing rows straight out.");
    }

    private static XLWorkbook BuildGrid(BenchmarkData data)
    {
        var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Data");

        ws.Cell(1, 1).Value = "Name";
        ws.Cell(1, 2).Value = "Amount";
        ws.Cell(1, 3).Value = "Date";

        for (var i = 0; i < RowCount; i++)
        {
            var row = i + 2;
            ws.Cell(row, 1).Value = data.Strings[i];
            ws.Cell(row, 2).Value = data.Numbers[i];
            ws.Cell(row, 3).Value = data.Dates[i];
        }

        var sumRow = RowCount + 2;
        ws.Cell(sumRow, 1).Value = "Total";
        ws.Cell(sumRow, 2).FormulaA1 = $"SUM(B2:B{RowCount + 1})";
        return workbook;
    }

    private static XLWorkbook BuildFormatted(BenchmarkData data)
    {
        var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Formatted");

        FormattedSheetBuilder.WriteHeaders(ws);

        for (var i = 0; i < RowCount; i++)
        {
            var row = i + 2;
            var idx = i % data.Strings.Length;

            FormattedSheetBuilder.WriteRowData(ws, data, row, i, idx);

            if (i % 2 == 0)
                FormattedSheetBuilder.ApplyRowFormatting(ws, row, i);
        }

        return workbook;
    }

    private static XLWorkbook BuildTemplate()
    {
        var bytes = TemplateFixture.Build(sheetCount: 10, definedNames: 20, validations: 26, dataRows: 0);
        return new XLWorkbook(new MemoryStream(bytes, writable: false));
    }

    private static void StreamingWrite(BenchmarkData data, CompressionLevel level, Stream output)
    {
        using var workbook = XLStreamingWorkbook.Create(output, new XLStreamingOptions { CompressionLevel = level });

        var sheet = workbook.AddWorksheet("Data");
        sheet.AppendRow("Name", "Amount", "Date");

        for (var i = 0; i < RowCount; i++)
            sheet.AppendRow(data.Strings[i], data.Numbers[i], data.Dates[i]);

        using (var row = sheet.AddRow())
        {
            row.Cell("Total");
            row.Formula($"SUM(B2:B{RowCount + 1})");
        }

        workbook.Finish();
    }

    private static double Median(double[] values)
    {
        var sorted = values.OrderBy(v => v).ToArray();
        return sorted[sorted.Length / 2];
    }
}
