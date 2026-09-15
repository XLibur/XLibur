using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Tracing;
using System.IO;
using System.Linq;
using System.Threading;
using JetBrains.Profiler.SelfApi;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.Coordinates;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// Where the first edit after a load spends its memory (#513): the whole dependency tree is built
/// by that edit, and kept until the workbook is disposed.
///
/// Run with: dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile firstedit [rows] [formulasPerRow] [dotmemory]
///
/// Defaults to 25,000 rows × 8 formulas. Every probe loads the workbook again, so none of them
/// sees a tree another probe built. With <c>dotmemory</c>, it also saves a dotMemory workspace
/// with snapshots after the load and after the first edit, in C:\profiles.
/// </summary>
public static class FirstEditProfile
{
    public static void Run(string[] args)
    {
        SixLaborsV1FontBootstrap.Register();
        var rows = args.Length > 2 ? int.Parse(args[2]) : 25_000;
        var formulasPerRow = args.Length > 3 ? int.Parse(args[3]) : 8;
        var formulas = rows * formulasPerRow;

        Console.WriteLine($"Building fixture: {rows:N0} rows x {formulasPerRow} formulas = {formulas:N0} formulas...");
        var package = FirstEditFixture.Build(rows, formulasPerRow);

        // Warm-up, so that the JIT is not in any measured number.
        var warmup = FirstEditFixture.Build(200, formulasPerRow);
        EditAfterLoad(warmup);
        BuildTree(warmup);
        ParseAll(warmup);

        Console.WriteLine();
        Console.WriteLine("| Probe | Allocated | per formula | Kept alive | per formula | ms |");
        Console.WriteLine("|---|---:|---:|---:|---:|---:|");
        Report("Load", formulas, Load(package));
        Report("First edit (public API)", formulas, EditAfterLoad(package));
        Report("DependencyTree.CreateFrom alone", formulas, BuildTree(package));
        Report("Parse every formula alone", formulas, ParseAll(package));
        Console.WriteLine();
        Console.WriteLine("Bytes are exact. Times are single-shot — use BenchmarkDotNet for time claims.");

        Console.WriteLine();
        Console.WriteLine("Allocation by type in DependencyTree.CreateFrom, sampled from GC AllocationTick events (~100 KB each):");
        SampleTreeBuild(package);

        if (args.Any(a => a.Equals("dotmemory", StringComparison.OrdinalIgnoreCase)))
            TakeDotMemorySnapshots(package);
    }

    private readonly record struct Probe(long Allocated, long KeptAlive, double Milliseconds);

    private static Probe Load(byte[] package)
    {
        ForceGC();
        var heapBefore = GC.GetTotalMemory(forceFullCollection: true);
        var allocBefore = GC.GetTotalAllocatedBytes(precise: true);
        var watch = Stopwatch.StartNew();
        var workbook = new XLWorkbook(new MemoryStream(package, writable: false));
        watch.Stop();
        var allocated = GC.GetTotalAllocatedBytes(precise: true) - allocBefore;
        ForceGC();
        var keptAlive = GC.GetTotalMemory(forceFullCollection: true) - heapBefore;
        workbook.Dispose();
        return new Probe(allocated, keptAlive, watch.Elapsed.TotalMilliseconds);
    }

    private static Probe EditAfterLoad(byte[] package)
    {
        using var workbook = new XLWorkbook(new MemoryStream(package, writable: false));
        var sheet = workbook.Worksheet(1);

        ForceGC();
        var heapBefore = GC.GetTotalMemory(forceFullCollection: true);
        var allocBefore = GC.GetTotalAllocatedBytes(precise: true);
        var watch = Stopwatch.StartNew();
        sheet.Cell(1, 1).Value = 5;
        watch.Stop();
        var allocated = GC.GetTotalAllocatedBytes(precise: true) - allocBefore;
        ForceGC();
        var keptAlive = GC.GetTotalMemory(forceFullCollection: true) - heapBefore;
        return new Probe(allocated, keptAlive, watch.Elapsed.TotalMilliseconds);
    }

    private static Probe BuildTree(byte[] package)
    {
        using var workbook = new XLWorkbook(new MemoryStream(package, writable: false));

        ForceGC();
        var heapBefore = GC.GetTotalMemory(forceFullCollection: true);
        var allocBefore = GC.GetTotalAllocatedBytes(precise: true);
        var watch = Stopwatch.StartNew();
        var tree = DependencyTree.CreateFrom(workbook);
        watch.Stop();
        var allocated = GC.GetTotalAllocatedBytes(precise: true) - allocBefore;
        ForceGC();
        var keptAlive = GC.GetTotalMemory(forceFullCollection: true) - heapBefore;
        GC.KeepAlive(tree);
        return new Probe(allocated, keptAlive, watch.Elapsed.TotalMilliseconds);
    }

    private static Probe ParseAll(byte[] package)
    {
        using var workbook = new XLWorkbook(new MemoryStream(package, writable: false));
        var texts = new List<string>();
        foreach (var sheet in workbook.WorksheetsInternal)
        {
            using var enumerator = sheet.Internals.CellsCollection.FormulaSlice.GetForwardEnumerator(Area.Full);
            while (enumerator.MoveNext())
                texts.Add(enumerator.Current.A1);
        }

        var engine = workbook.CalcEngine;
        ForceGC();
        var allocBefore = GC.GetTotalAllocatedBytes(precise: true);
        var watch = Stopwatch.StartNew();
        foreach (var text in texts)
            engine.TryParse(text, out _);
        watch.Stop();
        var allocated = GC.GetTotalAllocatedBytes(precise: true) - allocBefore;
        return new Probe(allocated, 0, watch.Elapsed.TotalMilliseconds);
    }

    private static void SampleTreeBuild(byte[] package)
    {
        using var workbook = new XLWorkbook(new MemoryStream(package, writable: false));
        ForceGC();

        using var sampler = new AllocationSampler();
        sampler.Start();
        var tree = DependencyTree.CreateFrom(workbook);
        GC.KeepAlive(tree);

        // Runtime events reach the listener on a dispatcher thread, a little after the allocation.
        Thread.Sleep(2_000);
        sampler.Stop();

        var byType = sampler.Snapshot();
        var total = byType.Sum(x => x.Value);
        Console.WriteLine("| Type | Sampled MB | Share |");
        Console.WriteLine("|---|---:|---:|");
        foreach (var (type, bytes) in byType.OrderByDescending(x => x.Value).Take(25))
            Console.WriteLine($"| `{type}` | {bytes / 1048576.0:F1} | {100.0 * bytes / total:F1}% |");
        Console.WriteLine($"| **All ({sampler.Samples:N0} samples)** | {total / 1048576.0:F1} | 100% |");
    }

    private static void TakeDotMemorySnapshots(byte[] package)
    {
        const string outputDir = @"C:\profiles";
        Directory.CreateDirectory(outputDir);
        Console.WriteLine();
        Console.WriteLine("Initializing dotMemory...");
        DotMemory.Init();
        DotMemory.Attach(new DotMemory.Config().SaveToDir(outputDir));
        try
        {
            using var workbook = new XLWorkbook(new MemoryStream(package, writable: false));
            ForceGC();
            DotMemory.GetSnapshot("After load");
            workbook.Worksheet(1).Cell(1, 1).Value = 5;
            ForceGC();
            DotMemory.GetSnapshot("After first edit");
        }
        finally
        {
            DotMemory.Detach();
        }

        Console.WriteLine($"dotMemory workspace saved to {outputDir}");
    }

    private static void Report(string name, int formulas, Probe probe)
    {
        var keptAlive = probe.KeptAlive == 0 ? "—" : $"{probe.KeptAlive / 1048576.0:F1} MB";
        var keptAlivePerFormula = probe.KeptAlive == 0 ? "—" : $"{(double)probe.KeptAlive / formulas:F0} B";
        Console.WriteLine(
            $"| {name} | {probe.Allocated / 1048576.0:F1} MB | {(double)probe.Allocated / formulas:F0} B | {keptAlive} | {keptAlivePerFormula} | {probe.Milliseconds:F0} |");
    }

    private static void ForceGC()
    {
#pragma warning disable S1215 // Intentionally forcing GC for accurate memory profiling
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
#pragma warning restore S1215
    }

    /// <summary>
    /// Sums the <c>GCAllocationTick</c> events of the runtime by type. The runtime raises one event
    /// about every 100 KB of allocation, naming the type of the object that crossed the threshold,
    /// so the sums are a statistical picture of where the bytes go, not exact counts.
    /// </summary>
    private sealed class AllocationSampler : EventListener
    {
        private const EventKeywords GcKeyword = (EventKeywords)0x1;
        private readonly object _gate = new();
        private readonly Dictionary<string, long> _bytesByType = new();
        private volatile bool _recording;

        public long Samples { get; private set; }

        public void Start() => _recording = true;

        public void Stop() => _recording = false;

        public Dictionary<string, long> Snapshot()
        {
            lock (_gate)
                return new Dictionary<string, long>(_bytesByType);
        }

        protected override void OnEventSourceCreated(EventSource eventSource)
        {
            if (eventSource.Name == "Microsoft-Windows-DotNETRuntime")
                EnableEvents(eventSource, EventLevel.Verbose, GcKeyword);
        }

        protected override void OnEventWritten(EventWrittenEventArgs eventData)
        {
            if (!_recording || eventData.EventName is null ||
                !eventData.EventName.StartsWith("GCAllocationTick", StringComparison.Ordinal) ||
                eventData.PayloadNames is null || eventData.Payload is null)
                return;

            var typeIndex = eventData.PayloadNames.IndexOf("TypeName");
            var amountIndex = eventData.PayloadNames.IndexOf("AllocationAmount64");
            if (amountIndex < 0)
                amountIndex = eventData.PayloadNames.IndexOf("AllocationAmount");
            if (typeIndex < 0 || amountIndex < 0)
                return;

            var type = eventData.Payload[typeIndex] as string ?? "?";
            var amount = Convert.ToInt64(eventData.Payload[amountIndex]);
            lock (_gate)
            {
                _bytesByType[type] = _bytesByType.GetValueOrDefault(type) + amount;
                Samples++;
            }
        }
    }
}
