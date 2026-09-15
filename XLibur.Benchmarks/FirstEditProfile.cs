using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Tracing;
using System.IO;
using System.Linq;
using System.Threading;
using ClosedXML.Parser;
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

        var shared = args.Any(a => a.Equals("shared", StringComparison.OrdinalIgnoreCase));

        Console.WriteLine($"Building fixture: {rows:N0} rows x {formulasPerRow} formulas = {formulas:N0} formulas{(shared ? ", shared down each column" : "")}...");
        var package = FirstEditFixture.Build(rows, formulasPerRow, shared);

        // Warm-up, so that the JIT is not in any measured number.
        var warmup = FirstEditFixture.Build(200, formulasPerRow, shared);
        EditAfterLoad(warmup);
        BuildTree(warmup);
        ParseAll(warmup);
        ConvertAll(warmup, out _);
        WalkAll(warmup);

        Console.WriteLine();
        Console.WriteLine("| Probe | Allocated | per formula | Kept alive | per formula | ms |");
        Console.WriteLine("|---|---:|---:|---:|---:|---:|");
        Report("Load", formulas, Load(package));
        Report("First edit (public API)", formulas, EditAfterLoad(package));
        Report("DependencyTree.CreateFrom alone", formulas, BuildTree(package));
        Report("Parse every formula alone", formulas, ParseAll(package));
        Report("Convert every formula to R1C1 alone", formulas, ConvertAll(package, out var distinctR1C1));
        Report("Walk every formula, building no AST", formulas, WalkAll(package));
        Console.WriteLine();
        Console.WriteLine($"Distinct R1C1 texts: {distinctR1C1:N0} of {formulas:N0} formulas.");
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

    /// <summary>
    /// The cost of getting the R1C1 text of every formula from its A1 text. A parse cache keyed on
    /// R1C1 text pays this for each formula whose R1C1 text the loader did not keep (#513, change 5).
    /// </summary>
    private static Probe ConvertAll(byte[] package, out int distinct)
    {
        using var workbook = new XLWorkbook(new MemoryStream(package, writable: false));
        var formulas = new List<(string Text, Point Point)>();
        foreach (var sheet in workbook.WorksheetsInternal)
        {
            using var enumerator = sheet.Internals.CellsCollection.FormulaSlice.GetForwardEnumerator(Area.Full);
            while (enumerator.MoveNext())
                formulas.Add((enumerator.Current.A1, enumerator.Point));
        }

        var converted = new string[formulas.Count];
        ForceGC();
        var allocBefore = GC.GetTotalAllocatedBytes(precise: true);
        var watch = Stopwatch.StartNew();
        for (var i = 0; i < formulas.Count; ++i)
        {
            FormulaText.TryConvert(formulas[i].Text, formulas[i].Point, FormulaNotation.R1C1, out var r1c1, out _);
            converted[i] = r1c1;
        }

        watch.Stop();
        var allocated = GC.GetTotalAllocatedBytes(precise: true) - allocBefore;
        distinct = new HashSet<string>(converted, StringComparer.Ordinal).Count;
        return new Probe(allocated, 0, watch.Elapsed.TotalMilliseconds);
    }

    /// <summary>
    /// The floor for extracting precedents while parsing (#513, change 6): the parser walks every
    /// formula with a factory that builds nothing, so what is left is the work of the tokenizer and
    /// the parser itself.
    /// </summary>
    private static Probe WalkAll(byte[] package)
    {
        using var workbook = new XLWorkbook(new MemoryStream(package, writable: false));
        var texts = new List<string>();
        foreach (var sheet in workbook.WorksheetsInternal)
        {
            using var enumerator = sheet.Internals.CellsCollection.FormulaSlice.GetForwardEnumerator(Area.Full);
            while (enumerator.MoveNext())
                texts.Add(enumerator.Current.A1);
        }

        ForceGC();
        var allocBefore = GC.GetTotalAllocatedBytes(precise: true);
        var watch = Stopwatch.StartNew();
        foreach (var text in texts)
            FormulaText.TryWalk(text, text, NoAstFactory.Instance, FormulaNotation.A1, out _, out _);
        watch.Stop();
        var allocated = GC.GetTotalAllocatedBytes(precise: true) - allocBefore;
        return new Probe(allocated, 0, watch.Elapsed.TotalMilliseconds);
    }

    /// <summary>
    /// A factory that builds nothing: every node is <c>0</c>.
    /// </summary>
    private sealed class NoAstFactory : IAstFactory<int, int, string>
    {
        public static readonly NoAstFactory Instance = new();

        public int LogicalValue(string context, SymbolRange range, bool value) => 0;
        public int NumberValue(string context, SymbolRange range, double value) => 0;
        public int TextValue(string context, SymbolRange range, string text) => 0;
        public int ErrorValue(string context, SymbolRange range, ReadOnlySpan<char> error) => 0;
        public int ArrayNode(string context, SymbolRange range, int rows, int columns, IReadOnlyList<int> elements) => 0;
        public int BlankNode(string context, SymbolRange range) => 0;
        public int LogicalNode(string context, SymbolRange range, bool value) => 0;
        public int ErrorNode(string context, SymbolRange range, ReadOnlySpan<char> error) => 0;
        public int SheetErrorNode(string context, SymbolRange range, int? workbookIndex, string sheet, ReadOnlySpan<char> error) => 0;
        public int NumberNode(string context, SymbolRange range, double value) => 0;
        public int TextNode(string context, SymbolRange range, string text) => 0;
        public int Reference(string context, SymbolRange range, ReferenceArea reference) => 0;
        public int SheetReference(string context, SymbolRange range, string sheet, ReferenceArea reference) => 0;
        public int BangReference(string context, SymbolRange range, ReferenceArea reference) => 0;
        public int Reference3D(string context, SymbolRange range, string firstSheet, string lastSheet, ReferenceArea reference) => 0;
        public int ExternalSheetReference(string context, SymbolRange range, int workbookIndex, string sheet, ReferenceArea reference) => 0;
        public int ExternalReference3D(string context, SymbolRange range, int workbookIndex, string firstSheet, string lastSheet, ReferenceArea reference) => 0;
        public int Function(string context, SymbolRange range, ReadOnlySpan<char> functionName, IReadOnlyList<int> arguments) => 0;
        public int Function(string context, SymbolRange range, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<int> args) => 0;
        public int ExternalFunction(string context, SymbolRange range, int workbookIndex, string sheetName, ReadOnlySpan<char> functionName, IReadOnlyList<int> arguments) => 0;
        public int ExternalFunction(string context, SymbolRange range, int workbookIndex, ReadOnlySpan<char> functionName, IReadOnlyList<int> arguments) => 0;
        public int CellFunction(string context, SymbolRange range, RowCol cell, IReadOnlyList<int> arguments) => 0;
        public int StructureReference(string context, SymbolRange range, StructuredReferenceArea area, string? firstColumn, string? lastColumn) => 0;
        public int StructureReference(string context, SymbolRange range, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn) => 0;
        public int ExternalStructureReference(string context, SymbolRange range, int workbookIndex, string table, StructuredReferenceArea area, string? firstColumn, string? lastColumn) => 0;
        public int Name(string context, SymbolRange range, string name) => 0;
        public int SheetName(string context, SymbolRange range, string sheet, string name) => 0;
        public int BangName(string context, SymbolRange range, string name) => 0;
        public int ExternalName(string context, SymbolRange range, int workbookIndex, string name) => 0;
        public int ExternalSheetName(string context, SymbolRange range, int workbookIndex, string sheet, string name) => 0;
        public int ExternalDynamicDataExchange(string context, SymbolRange range, int workbookIndex, string item) => 0;
        public int DynamicDataExchange(string context, SymbolRange range, string application, string topic, string item) => 0;
        public int BinaryNode(string context, SymbolRange range, BinaryOperation operation, int leftNode, int rightNode) => 0;
        public int Unary(string context, SymbolRange range, UnaryOperation operation, int node) => 0;
        public int Nested(string context, SymbolRange range, int node) => 0;
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
