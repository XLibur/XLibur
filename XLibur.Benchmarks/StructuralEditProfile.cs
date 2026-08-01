using System;
using System.Collections.Generic;
using System.Diagnostics;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// Cost of repeated single-row inserts, decomposed into the two things an insert pays for: the
/// range-shift notification pass (which visits live ranges) and the formula shift (which visits
/// formula cells, on every worksheet).
///
/// Run with: dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile structural
///
/// The decomposition is the point. Spec 05 workstream A has two independent halves — A2 attacks the
/// range pass, A3 the formula pass — and a combined number cannot tell you which one you just moved,
/// or which one is worth attacking first. Each probe isolates one by leaving the other at zero.
/// </summary>
public static class StructuralEditProfile
{
    private const int Inserts = 1_000;
    private const int LiveRanges = 1_000;
    private const int FormulaCells = 1_000;

    public static void Run()
    {
        SixLaborsV1FontBootstrap.Register();

        // Warm up: JIT, the style repository and the font engine, so their one-time cost does not
        // land on whichever probe runs first.
        RunProbe(("warmup", () => Insert(liveRanges: 10, formulaCells: 10)));

        Console.WriteLine();
        Console.WriteLine($"{Inserts:N0} single-row inserts, one InsertRowsAbove(1) at a time, all at row {InsertAtRow:N0}");
        Console.WriteLine("| Probe                                             |    Total |    ms |   us/insert |");
        Console.WriteLine("|---------------------------------------------------|----------|-------|-------------|");

        foreach (var probe in Probes())
            RunProbe(probe);

        Console.WriteLine();
        Console.WriteLine("Bytes are exact. Times are single-shot — use BenchmarkDotNet for time claims.");
    }

    private static (string Name, Action Action)[] Probes() =>
    [
        ($"{Inserts:N0} inserts, 0 ranges, 0 formulas", () => Insert(0, 0)),
        ($"{Inserts:N0} inserts, {LiveRanges:N0} ranges below (affected)", () => Insert(LiveRanges, 0, Placement.Below)),
        ($"{Inserts:N0} inserts, {LiveRanges:N0} ranges above (no-ops)", () => Insert(LiveRanges, 0, Placement.Above)),
        ($"{Inserts:N0} inserts, {FormulaCells:N0} formulas below", () => Insert(0, FormulaCells, Placement.Below)),
        ($"{Inserts:N0} inserts, {FormulaCells:N0} formulas above", () => Insert(0, FormulaCells, Placement.Above)),
        ($"{Inserts:N0} inserts, both below", () => Insert(LiveRanges, FormulaCells, Placement.Below)),
        ($"1 batch insert of {Inserts:N0}, both below", () => InsertBatch(LiveRanges, FormulaCells)),
    ];

    /// <summary>
    /// Where the content sits relative to the insert point. The insert row itself is held fixed across
    /// every probe, because moving it would also change how many rows <c>XLRowsCollection.ShiftRowsDown</c>
    /// has to renumber — a third cost that would otherwise be confounded with the two under study.
    /// <para>
    /// <see cref="Above"/> content cannot be affected by the insert, so those probes measure exactly the
    /// cost of *visiting* candidates that turn out to be no-ops. The gap between Above and Below is the
    /// most that a filter or an index could ever recover.
    /// </para>
    /// </summary>
    private enum Placement
    {
        Below,
        Above,
    }

    /// <summary>Fixed insert point, with room for <see cref="Placement.Above"/> content beneath it.</summary>
    private const int InsertAtRow = 1_005;

    private static void RunProbe((string Name, Action Action) probe)
    {
        ForceGC();

        var before = GC.GetTotalAllocatedBytes(precise: true);
        var watch = Stopwatch.StartNew();
        probe.Action();
        watch.Stop();
        var bytes = GC.GetTotalAllocatedBytes(precise: true) - before;

        if (probe.Name == "warmup")
            return;

        Console.WriteLine(
            $"| {probe.Name,-49} | {bytes / 1048576.0,5:F1} MB | {watch.Elapsed.TotalMilliseconds,5:F0} | {watch.Elapsed.TotalMicroseconds / Inserts,11:F1} |");
    }

    private static void Insert(int liveRanges, int formulaCells, Placement placement = Placement.Below)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var held = Populate(ws, liveRanges, formulaCells, placement);

        for (var i = 0; i < Inserts; i++)
            ws.Row(InsertAtRow).InsertRowsAbove(1);

        GC.KeepAlive(held);
    }

    private static void InsertBatch(int liveRanges, int formulaCells)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        var held = Populate(ws, liveRanges, formulaCells, Placement.Below);

        ws.Row(InsertAtRow).InsertRowsAbove(Inserts);

        GC.KeepAlive(held);
    }

    /// <summary>
    /// The range list has to be held: the range repository stores weak references, so ranges nobody
    /// keeps a reference to are collected and the shift pass never sees them — which would quietly
    /// benchmark an empty sheet.
    /// </summary>
    private static List<IXLRange> Populate(IXLWorksheet ws, int liveRanges, int formulaCells, Placement placement)
    {
        // Above content ends one row clear of the insert point; below content starts one row past it.
        var firstRow = placement == Placement.Above ? InsertAtRow - liveRanges - formulaCells - 1 : InsertAtRow + 1;

        var held = new List<IXLRange>(liveRanges);
        for (var r = firstRow; r < firstRow + liveRanges; r++)
            held.Add(ws.Range(r, 1, r, 5));

        for (var r = firstRow; r < firstRow + formulaCells; r++)
            ws.Cell(r, 8).FormulaA1 = $"A{r}+B{r}";

        return held;
    }

    private static void ForceGC()
    {
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
    }
}
