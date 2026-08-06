using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;
using static XLibur.Benchmarks.TemplateFixture;

namespace XLibur.Benchmarks;

/// <summary>
/// Cost of the template-driven export shape: open an existing workbook, change a little, save it
/// again. Decomposed so the fixed per-request overhead — the part that scales with neither rows
/// written nor data changed — can be separated from the cost of the edit itself.
///
/// Run with: dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile template [path-to.xlsx]
///
/// The decomposition is the point. Profiling a template-driven export in a consuming application
/// put roughly 300 ms per request into opening and re-saving a 124 KB workbook carrying ~10
/// sheets, ~20 defined names and ~26 data validations, purely to write a single column of lookup
/// values. A combined number cannot say whether that lands in the parse or the serialise, nor
/// which structural feature drives it, so each probe below isolates one candidate.
/// </summary>
/// <remarks>
/// By default the fixture is generated. A synthetic fixture is a poor stand-in for a real
/// reporting template — shared strings, styles, tables, spilled dynamic-array formulas,
/// conditional formatting, images and external links all live in a real file and none are
/// reproduced here — so measurements against it understate real-world cost. Point the probe at an
/// actual file to reproduce production numbers, either by argument or by environment variable:
/// <code>
/// $env:XLIBUR_PERF_TEMPLATE = "C:\path\to\Template.xlsx"
/// </code>
/// </remarks>
public static class TemplateRoundTripProfile
{
    /// <summary>
    /// Timed passes per probe. Nine rather than a handful because the spread between passes on a
    /// loaded desktop is comfortably wider than the effect sizes worth chasing here — a five-pass
    /// median moved by ±15% run to run, enough to invent regressions that were not there.
    /// </summary>
    private const int Iterations = 9;

    public static void Run(string[] args)
    {
        SixLaborsV1FontBootstrap.Register();

        var template = ResolveTemplatePath(args);

        // Warm up the JIT, the style repository and the font engine so their one-time cost does
        // not land on whichever probe runs first. The fixture deliberately exercises every
        // structural feature the probes use, and runs more than once: an earlier version warmed
        // up on a bare two-sheet workbook and left enough un-jitted code that the first real probe
        // read ~30% slow, which is the same size as the effects being measured.
        var warmup = Build(sheetCount: 4, definedNames: 4, validations: 4, dataRows: 20);
        for (var i = 0; i < 3; i++)
        {
            using var state = Open(warmup);
            WriteGrid(ResolveDataSheet(state.Workbook), rows: 50, GridColumns, perCellNumberFormat: true);
            using var sink = new MemoryStream();
            state.Workbook.SaveAs(sink);
        }

        if (template is not null)
            Console.WriteLine($"fixture: {template} (generated-fixture parameters ignored)");

        if (args.Length > 2 && args[2].Equals("loop", StringComparison.OrdinalIgnoreCase))
        {
            Loop(template, phase: args.Length > 3 ? args[3] : "roundtrip");
            return;
        }

        RoundTripCost(template);
        OpenVersusSaveAttribution(template);
        LookupColumnRefresh(template);
        DataGridWrite(template);
        NumberFormatPerCellVersusPerColumn(template);
        FullCycle(template);

        Console.WriteLine();
        Console.WriteLine("Bytes are exact. Times are medians of "
            + $"{Iterations} passes after a warm-up — use BenchmarkDotNet for time claims.");
    }

    /// <summary>
    /// Repeats one phase for a fixed wall-clock window so an external profiler can be attached to
    /// it, e.g.
    /// <code>
    /// dotnet-trace collect --format speedscope -- XLibur.Benchmarks.exe profile template loop open
    /// </code>
    /// </summary>
    /// <remarks>
    /// The table-producing probes are the wrong shape to profile: they interleave six unrelated
    /// workloads with forced gen2 collections between every pass, so a captured trace mixes them
    /// together and the GC work swamps the code under study. This runs one phase and nothing else.
    /// </remarks>
    private static void Loop(string? template, string phase)
    {
        var bytes = GetFixture(template, sheetCount: 10, definedNames: 20, validations: 26, dataRows: 0);
        RoundTrip(bytes); // JIT warm-up, so start-up does not colour the trace.

        Console.WriteLine($"looping '{phase}' for {LoopSeconds}s...");
        var elapsed = Stopwatch.StartNew();
        var iterations = 0;

        while (elapsed.Elapsed.TotalSeconds < LoopSeconds)
        {
            switch (phase)
            {
                case "open":
                    using (var input = new MemoryStream(bytes, writable: false))
                    using (var workbook = new XLWorkbook(input))
                        _ = workbook.Worksheets.Count;
                    break;

                case "save":
                    using (var state = Open(bytes))
                    using (var sink = new MemoryStream())
                        state.Workbook.SaveAs(sink);
                    break;

                default:
                    RoundTrip(bytes);
                    break;
            }

            iterations++;
        }

        Console.WriteLine($"{iterations:N0} iterations in {elapsed.Elapsed.TotalSeconds:F1}s");
    }

    private const int LoopSeconds = 20;

    // ── 1. Round-trip cost ────────────────────────────────────────────────────

    /// <summary>
    /// Opens a workbook and saves it again with no modification whatsoever. This is the floor cost
    /// every caller pays just to touch a file, and in the profiled application it was the single
    /// largest contributor. The cases vary the workbook's structural weight to show which feature
    /// drives the cost.
    /// </summary>
    private static void RoundTripCost(string? template)
    {
        Header("open + save, unmodified");

        foreach (var (sheets, names, validations) in new[]
                 {
                     // The first three vary only the sheet count, so the marginal cost of one
                     // structurally empty worksheet — the part of the bill that scales with
                     // nothing a caller controls — falls out of the slope.
                     (1, 0, 0),
                     (10, 0, 0),
                     (40, 0, 0),
                     (10, 20, 0),
                     (10, 20, 26),
                     (10, 20, 100),
                 })
        {
            var bytes = GetFixture(template, sheets, names, validations, dataRows: 0);
            Row($"sheets={sheets} names={names} validations={validations}",
                Measure(() => RoundTrip(bytes)), bytes.Length);
        }
    }

    /// <summary>
    /// Attributes the round trip between parse and serialise. Save is reported as
    /// (open+save) − (open) rather than measured directly, because a second save of the same
    /// instance is not the same operation as the first: <c>SaveAs</c> adopts its destination as
    /// the workbook's origin, so the next save edits the package it just wrote instead of the
    /// template. Every probe therefore opens a fresh workbook per iteration.
    /// </summary>
    /// <remarks>
    /// The adoption is also a trap worth knowing about when writing probes. Because the previous
    /// destination becomes the origin, disposing it — which a <c>using</c> on a scratch
    /// <see cref="MemoryStream"/> does — makes the <em>next</em> <c>SaveAs</c> throw
    /// <see cref="ObjectDisposedException"/> ("Cannot access a closed Stream") from deep inside,
    /// where it reads that origin back. It is not that saving twice is unsupported: two saves to
    /// two streams that both stay alive succeed, and the stream a workbook was loaded from is
    /// left readable.
    /// </remarks>
    private static void OpenVersusSaveAttribution(string? template)
    {
        Header("open versus save attribution");

        var bytes = GetFixture(template, sheetCount: 10, definedNames: 20, validations: 26, dataRows: 0);

        var open = Measure(() =>
        {
            using var input = new MemoryStream(bytes, writable: false);
            using var workbook = new XLWorkbook(input);
            _ = workbook.Worksheets.Count;
        });

        var roundTrip = Measure(() => RoundTrip(bytes));

        Row("open only", open, bytes.Length);
        Row("open + save", roundTrip, bytes.Length);
        Row("save (by subtraction)", roundTrip - open, bytes.Length);
    }

    // ── 2. Single-column lookup refresh ───────────────────────────────────────

    /// <summary>
    /// The "refresh one lookup column" shape: clear a column, write N strings, repoint a defined
    /// name, save. The write itself is trivial, so anything above the round-trip floor is
    /// attributable to the edit.
    /// </summary>
    private static void LookupColumnRefresh(string? template)
    {
        Header("lookup refresh (open -> clear -> write -> repoint name -> save)");

        var bytes = GetFixture(template, sheetCount: 10, definedNames: 20, validations: 26, dataRows: 0);

        foreach (var values in new[] { 100, 1_000, 10_000 })
        {
            var items = Enumerable.Range(1, values).Select(i => $"Lookup value {i}").ToArray();

            Row($"values={values:N0}", Measure(() =>
            {
                using var input = new MemoryStream(bytes, writable: false);
                using var workbook = new XLWorkbook(input);
                var sheet = ResolveLookupSheet(workbook);

                var lastRow = sheet.LastRowUsed()?.RowNumber() ?? HeaderRow;
                if (lastRow > HeaderRow)
                    sheet.Range(HeaderRow + 1, 1, lastRow, 1).Clear(XLClearOptions.Contents);

                for (var i = 0; i < items.Length; i++)
                    sheet.Cell(HeaderRow + 1 + i, 1).Value = items[i];

                if (workbook.DefinedNames.TryGetValue(LookupRangeName, out var definedName))
                    definedName.SetRefersTo(sheet.Range(HeaderRow + 1, 1, HeaderRow + items.Length, 1));

                using var output = new MemoryStream();
                workbook.SaveAs(output);
            }), bytes.Length);
        }
    }

    // ── 3. Bulk grid write ────────────────────────────────────────────────────

    /// <summary>
    /// Writes a wide grid of mixed cell types, the shape of a data export. Isolates per-cell write
    /// throughput from the round-trip floor by timing the in-memory writes separately from the
    /// save; the open that precedes each stays outside the measurement.
    /// </summary>
    private static void DataGridWrite(string? template)
    {
        Header("grid write, split into cell writes and save");

        var bytes = GetFixture(template, sheetCount: 10, definedNames: 20, validations: 26, dataRows: 0);

        foreach (var rows in new[] { 1_000, 5_000, 20_000 })
        {
            var write = MeasureExcludingSetup(
                () => Open(bytes),
                state => WriteGrid(ResolveDataSheet(state.Workbook), rows, GridColumns, perCellNumberFormat: false));

            var save = MeasureExcludingSetup(
                () =>
                {
                    var state = Open(bytes);
                    WriteGrid(ResolveDataSheet(state.Workbook), rows, GridColumns, perCellNumberFormat: false);
                    return state;
                },
                state =>
                {
                    using var output = new MemoryStream();
                    state.Workbook.SaveAs(output);
                });

            Row($"rows={rows:N0} cols={GridColumns} (cells only)", write, bytes.Length);
            Row($"rows={rows:N0} cols={GridColumns} (save only)", save, bytes.Length);
        }
    }

    // ── 4. Per-cell versus per-column styling ─────────────────────────────────

    /// <summary>
    /// A/B on number-format application. Setting <c>Style.NumberFormat.Format</c> on every cell is
    /// the natural way to write a formatted column and is what the profiled application did;
    /// setting it once on the column is the alternative. If per-cell styling is materially slower,
    /// style resolution or the style cache is a bottleneck worth attention.
    /// </summary>
    private static void NumberFormatPerCellVersusPerColumn(string? template)
    {
        Header("date column, per-cell versus per-column number format");

        var bytes = GetFixture(template, sheetCount: 10, definedNames: 20, validations: 26, dataRows: 0);

        foreach (var rows in new[] { 1_000, 5_000, 20_000 })
        {
            var perCell = MeasureExcludingSetup(
                () => Open(bytes),
                state => WriteDateColumn(ResolveDataSheet(state.Workbook), rows, column: 1, perCellFormat: true));

            var perColumn = MeasureExcludingSetup(
                () => Open(bytes),
                state =>
                {
                    var sheet = ResolveDataSheet(state.Workbook);
                    sheet.Column(1).Style.NumberFormat.Format = DateFormat;
                    WriteDateColumn(sheet, rows, column: 1, perCellFormat: false);
                });

            Row($"rows={rows:N0} per-cell", perCell, bytes.Length);
            Row($"rows={rows:N0} per-column", perColumn, bytes.Length);
            Console.WriteLine(
                $"| {"  ratio per-cell / per-column",-49} | {SafeRatio(perCell.Milliseconds, perColumn.Milliseconds),7:F2}x | {SafeRatio(perCell.FastestMs, perColumn.FastestMs),7:F2}x | {SafeRatio(perCell.Bytes, perColumn.Bytes),7:F2}x |             |");
        }
    }

    // ── 5. End-to-end ─────────────────────────────────────────────────────────

    /// <summary>
    /// The complete shape the profiled application performs per request: open a template, clear
    /// stale rows, write the grid, save. Provided so a change inside the library can be judged
    /// against the whole operation rather than one stage.
    /// </summary>
    private static void FullCycle(string? template)
    {
        Header("full cycle (open -> clear -> write -> save)");

        var bytes = GetFixture(template, sheetCount: 10, definedNames: 20, validations: 26, dataRows: 200);

        foreach (var rows in new[] { 1_000, 5_000 })
        {
            Row($"rows={rows:N0}", Measure(() =>
            {
                using var input = new MemoryStream(bytes, writable: false);
                using var workbook = new XLWorkbook(input);
                var sheet = ResolveDataSheet(workbook);

                var lastRow = sheet.LastRowUsed()?.RowNumber() ?? HeaderRow;
                if (lastRow >= FirstDataRow)
                {
                    var lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 1;
                    sheet.Range(FirstDataRow, 1, lastRow, lastColumn).Clear(XLClearOptions.Contents);
                }

                WriteGrid(sheet, rows, GridColumns, perCellNumberFormat: true);

                using var output = new MemoryStream();
                workbook.SaveAs(output);
            }), bytes.Length);
        }
    }

    // ── fixture ───────────────────────────────────────────────────────────────

    private static string? ResolveTemplatePath(string[] args)
    {
        // `profile template <path>` wins over the environment variable. "loop" is the profiler
        // sub-command rather than a path, so it falls through to the environment variable.
        var argPath = args.Length > 2 && !args[2].Equals("loop", StringComparison.OrdinalIgnoreCase)
            ? args[2]
            : null;
        var path = argPath ?? Environment.GetEnvironmentVariable("XLIBUR_PERF_TEMPLATE");

        if (string.IsNullOrWhiteSpace(path))
            return null;

        if (!File.Exists(path))
            throw new FileNotFoundException($"Template fixture points at a missing file: {path}", path);

        return path;
    }

    /// <summary>
    /// Returns the fixture workbook: the external template when one was supplied, otherwise a
    /// generated one. The sheetCount / definedNames / validations parameters describe how to
    /// <em>build</em> a fixture, so they have no effect on an external template — cases that
    /// differ only in those values become repeat samples of one measurement.
    /// </summary>
    private static byte[] GetFixture(string? template, int sheetCount, int definedNames, int validations, int dataRows) =>
        template is not null
            ? File.ReadAllBytes(template)
            : Build(sheetCount, definedNames, validations, dataRows);

    private static void RoundTrip(byte[] bytes)
    {
        using var input = new MemoryStream(bytes, writable: false);
        using var workbook = new XLWorkbook(input);
        using var output = new MemoryStream();
        workbook.SaveAs(output);
    }

    /// <summary>
    /// Opening produces two disposables that must outlive the measured call, because
    /// <see cref="XLWorkbook"/> reads lazily from the stream it was constructed over.
    /// </summary>
    private readonly record struct OpenWorkbook(MemoryStream Input, XLWorkbook Workbook) : IDisposable
    {
        public void Dispose()
        {
            Workbook.Dispose();
            Input.Dispose();
        }
    }

    private static OpenWorkbook Open(byte[] bytes)
    {
        var input = new MemoryStream(bytes, writable: false);
        return new OpenWorkbook(input, new XLWorkbook(input));
    }

    // ── measurement ───────────────────────────────────────────────────────────

    /// <param name="Milliseconds">Median of the timed passes.</param>
    /// <param name="FastestMs">
    /// Fastest timed pass. The least-disturbed sample, and in practice the more stable estimator
    /// of the two: it is bounded below by the real work, whereas the median still drifts with
    /// whatever else the machine is doing. Compare this one across builds.
    /// </param>
    /// <param name="Bytes">Allocation of the median pass. Exact, and not subject to timing noise.</param>
    private readonly record struct Sample(double Milliseconds, double FastestMs, long Bytes)
    {
        public static Sample operator -(Sample left, Sample right) =>
            new(left.Milliseconds - right.Milliseconds,
                left.FastestMs - right.FastestMs,
                left.Bytes - right.Bytes);
    }

    /// <summary>
    /// Runs <paramref name="action"/> for <see cref="Iterations"/> timed passes and returns the
    /// median time with the allocation of the median pass. Median rather than mean so a single GC
    /// pause does not dominate the result.
    /// </summary>
    private static Sample Measure(Action action) =>
        MeasureExcludingSetup<object?>(() => null, _ => action());

    /// <summary>
    /// As <see cref="Measure"/>, but with per-iteration setup held outside the measurement.
    /// Setup is unavoidable for the split probes: a saved workbook has adopted its destination as
    /// its origin and would not save the same work a second time, so each pass
    /// has to open a fresh one, and that open would otherwise swamp what is being measured.
    /// </summary>
    private static Sample MeasureExcludingSetup<TState>(Func<TState> setup, Action<TState> action)
    {
        var times = new double[Iterations];
        var allocations = new long[Iterations];

        for (var i = 0; i < Iterations; i++)
        {
            var state = setup();
            try
            {
                ForceGC();

                var before = GC.GetTotalAllocatedBytes(precise: true);
                var watch = Stopwatch.StartNew();
                action(state);
                watch.Stop();

                times[i] = watch.Elapsed.TotalMilliseconds;
                allocations[i] = GC.GetTotalAllocatedBytes(precise: true) - before;
            }
            finally
            {
                (state as IDisposable)?.Dispose();
            }
        }

        // Sorted independently: the median allocation is wanted as a robust figure in its own
        // right, not as the allocation that happened to accompany the median time.
        Array.Sort(times);
        Array.Sort(allocations);
        return new Sample(times[Iterations / 2], times[0], allocations[Iterations / 2]);
    }

    private static void ForceGC()
    {
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
    }

    /// <summary>
    /// The sheet the grid probes write into: the generated fixture's data sheet when present,
    /// otherwise the first worksheet. Falling back positionally is what lets an arbitrary external
    /// template be used without knowing its sheet names.
    /// </summary>
    private static IXLWorksheet ResolveDataSheet(IXLWorkbook workbook) =>
        workbook.Worksheets.TryGetWorksheet(DataSheet, out var sheet)
            ? sheet
            : workbook.Worksheets.First();

    /// <summary>
    /// The sheet the lookup probe writes into: the generated fixture's first lookup sheet when
    /// present, otherwise the last worksheet — chosen so it is not the same sheet as
    /// <see cref="ResolveDataSheet"/> in a multi-sheet template.
    /// </summary>
    private static IXLWorksheet ResolveLookupSheet(IXLWorkbook workbook) =>
        workbook.Worksheets.TryGetWorksheet(FirstLookupSheet, out var sheet)
            ? sheet
            : workbook.Worksheets.Last();

    private static double SafeRatio(double numerator, double denominator) =>
        denominator <= 0 ? double.NaN : numerator / denominator;

    private static void Header(string title)
    {
        Console.WriteLine();
        Console.WriteLine(title);
        Console.WriteLine("| Probe                                             |   median |  fastest |    Alloc |   Fixture   |");
        Console.WriteLine("|---------------------------------------------------|----------|----------|----------|-------------|");
    }

    private static void Row(string label, Sample sample, int fixtureBytes) =>
        Console.WriteLine(
            $"| {label,-49} | {sample.Milliseconds,7:F1}m | {sample.FastestMs,7:F1}m | {sample.Bytes / 1048576.0,5:F1} MB | {fixtureBytes,8:N0} B |");
}
