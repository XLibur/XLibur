using System;
using System.Diagnostics;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// Cost of the dependency-tree <c>MarkDirty</c> walk on the bulk-edit path: many single-cell
/// value writes against an already-built dependency tree, once the tree has real dependents to
/// walk and once it has none. Spec 40 replaces the walk's "already visited" check (which reused
/// a formula's dirty flag) with its own tracking; this profile is the before/after gate that
/// change must not regress.
///
/// Run with: dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile bulkedit
/// </summary>
public static class BulkEditDirtyWalkProfile
{
    private const int Edits = 200_000;
    private const int Chains = 200;
    private const int ChainDepth = 10;
    private const int UnsettledInputs = 20_000;
    private const int UnsettledDepth = 50;

    public static void Run()
    {
        SixLaborsV1FontBootstrap.Register();

        RunProbe(("warmup", () => EditRootsWithDependents(chains: 10, depth: 5, edits: 1_000), 1));

        Console.WriteLine();
        Console.WriteLine("Single-cell value writes; edit count is named per probe.");
        Console.WriteLine("| Probe                                             |    Total |    ms |   us/edit |");
        Console.WriteLine("|---------------------------------------------------|----------|-------|-----------|");

        foreach (var probe in Probes())
            RunProbe(probe);

        Console.WriteLine();
        Console.WriteLine("Bytes are exact. Times are single-shot — use BenchmarkDotNet for time claims.");
    }

    private static (string Name, Action Action, int Edits)[] Probes() =>
    [
        ($"{Edits:N0} edits, {Chains}x{ChainDepth} chains (real walk)", () => EditRootsWithDependents(Chains, ChainDepth, Edits), Edits),
        ($"{Edits:N0} edits, no dependents (walk finds nothing)", () => EditCellsWithNoDependents(Edits), Edits),
        ($"{UnsettledInputs:N0} unsettled edits, {UnsettledDepth}-deep shared model", () => EditSharedModelWithoutSettling(UnsettledInputs, UnsettledDepth), UnsettledInputs),
    ];

    /// <summary>
    /// The worst case for a walk that tracks visits per call rather than reusing the dirty flag:
    /// every input of a shared model is written with no read in between, so nothing ever cleans
    /// the model and every edit re-walks the same dependent closure.
    /// </summary>
    /// <remarks>
    /// This is the one workload spec 40's fix genuinely makes more expensive, and it is measured
    /// here rather than assumed away. Reusing the dirty flag as the visited marker also
    /// short-circuited <i>across</i> calls: once the model was dirty from the first edit, every
    /// later edit stopped at the first hop, making the sequence roughly O(N + M) instead of
    /// O(N x M). That saving was never sound — it is the same shortcut that pruned legitimate
    /// walks and produced stale values, which is why it is gone — so the numbers below are the
    /// price of correctness on this shape, not a regression to be optimised away by restoring it.
    /// Anyone tempted to reintroduce a cross-call skip needs a marker that distinguishes "this
    /// subtree is dirty because a walk reached it" from "dirty for some unrelated reason";
    /// the dirty flag alone cannot.
    /// </remarks>
    private static void EditSharedModelWithoutSettling(int inputs, int depth)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        for (var r = 1; r <= inputs; r++)
            ws.Cell(r, 1).Value = r;

        // Every input feeds the head of the model, and each link feeds the next, so a walk from
        // any single input has the whole depth-deep chain downstream of it.
        ws.Cell(1, 2).FormulaA1 = $"SUM(A1:A{inputs})";
        for (var r = 2; r <= depth; r++)
            ws.Cell(r, 2).FormulaA1 = $"B{r - 1}+1";

        foreach (var cell in ws.CellsUsed(x => x.HasFormula))
            _ = cell.Value;

        // No read between edits: the model stays dirty throughout.
        for (var r = 1; r <= inputs; r++)
            ws.Cell(r, 1).Value = r + 1;
    }

    /// <summary>
    /// Every root feeds a chain of <paramref name="depth"/> formulas, so each edit's walk has real
    /// work: it must traverse the whole chain and mark every link dirty.
    /// </summary>
    /// <remarks>
    /// Each edit is immediately followed by reading every cell of the edited chain, top to
    /// bottom, which cleans every link before the next edit without ever needing the chain's
    /// exception-driven <c>GettingDataException</c>/full-recalculate fallback (each read's own
    /// precedent was just cleaned by the read before it). Two things this deliberately avoids:
    /// <list type="bullet">
    /// <item>
    /// Not settling at all would let a column sampled twice in a row hit the pre-fix code's
    /// dirty-as-visited shortcut on the second edit (its first dependent is already dirty from
    /// the first edit and nothing ever cleaned it) and end the walk after one hop — understating
    /// the pre-fix baseline by skipping exactly the work spec 40 restores, rather than measuring
    /// the walk's genuine per-edit cost.
    /// </item>
    /// <item>
    /// Settling with a single read of the chain's tail instead of level-by-level makes every edit
    /// pay for one <c>GettingDataException</c> cascade back to the root through
    /// <c>XLCalculationChain</c>'s reordering, which dominates the measurement by roughly three
    /// orders of magnitude and buries the walk's own cost — identically before and after this
    /// spec's fix, since neither touches the calculation chain.
    /// </item>
    /// </list>
    /// </remarks>
    private static void EditRootsWithDependents(int chains, int depth, int edits)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");

        for (var c = 1; c <= chains; c++)
        {
            ws.Cell(1, c).Value = 1;
            for (var r = 2; r <= depth + 1; r++)
                ws.Cell(r, c).FormulaA1 = $"{ws.Cell(r - 1, c).Address.ColumnLetter}{r - 1}+1";
        }

        // Force evaluation so the dependency tree exists and every formula starts clean.
        foreach (var cell in ws.CellsUsed(x => x.HasFormula))
            _ = cell.Value;

#pragma warning disable S2245 // Deterministic seed for reproducible benchmarks
        var rnd = new Random(42);
#pragma warning restore S2245
        for (var i = 0; i < edits; i++)
        {
            var column = rnd.Next(1, chains + 1);
            ws.Cell(1, column).Value = i;
            for (var r = 2; r <= depth + 1; r++)
                _ = ws.Cell(r, column).Value;
        }
    }

    /// <summary>
    /// A dependency tree exists (seeded by one unrelated formula), but the edited cells have no
    /// dependents at all, so every walk this triggers finds nothing — isolating the walk's
    /// per-call overhead from any work it does when it finds dependents.
    /// </summary>
    private static void EditCellsWithNoDependents(int edits)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Sheet1");
        ws.Cell(1, 1).Value = 1;
        ws.Cell(2, 1).FormulaA1 = "A1+1";
        _ = ws.Cell(2, 1).Value;

        for (var i = 0; i < edits; i++)
            ws.Cell(100 + i % 1000, 50).Value = i;
    }

    private static void RunProbe((string Name, Action Action, int Edits) probe)
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
            $"| {probe.Name,-49} | {bytes / 1048576.0,5:F1} MB | {watch.Elapsed.TotalMilliseconds,5:F0} | {watch.Elapsed.TotalMicroseconds / probe.Edits,9:F2} |");
    }

    /// <summary>
    /// Settles the heap so a probe's <c>GetTotalAllocatedBytes</c> delta is its own allocation and
    /// not the previous probe's garbage. Forcing the collector is the measurement here.
    /// </summary>
    private static void ForceGC()
    {
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
    }
}
