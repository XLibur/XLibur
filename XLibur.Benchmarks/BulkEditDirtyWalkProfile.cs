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

    public static void Run()
    {
        SixLaborsV1FontBootstrap.Register();

        RunProbe(("warmup", () => EditRootsWithDependents(chains: 10, depth: 5, edits: 1_000)));

        Console.WriteLine();
        Console.WriteLine($"{Edits:N0} single-cell value writes per probe");
        Console.WriteLine("| Probe                                             |    Total |    ms |   us/edit |");
        Console.WriteLine("|---------------------------------------------------|----------|-------|-----------|");

        foreach (var probe in Probes())
            RunProbe(probe);

        Console.WriteLine();
        Console.WriteLine("Bytes are exact. Times are single-shot — use BenchmarkDotNet for time claims.");
    }

    private static (string Name, Action Action)[] Probes() =>
    [
        ($"{Edits:N0} edits, {Chains}x{ChainDepth} chains (real walk)", () => EditRootsWithDependents(Chains, ChainDepth, Edits)),
        ($"{Edits:N0} edits, no dependents (walk finds nothing)", () => EditCellsWithNoDependents(Edits)),
    ];

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

        var rnd = new Random(42);
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
            $"| {probe.Name,-49} | {bytes / 1048576.0,5:F1} MB | {watch.Elapsed.TotalMilliseconds,5:F0} | {watch.Elapsed.TotalMicroseconds / Edits,9:F2} |");
    }

    private static void ForceGC()
    {
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
    }
}
