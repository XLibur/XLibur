using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// Isolates the cost of resolving a <c>ReferenceNode</c> to a range during formula
/// evaluation. <c>RecalculateAllFormulas</c> marks every formula dirty and then
/// recalculates, so each iteration re-resolves every reference and the measurement is
/// repeatable.
/// </summary>
/// <remarks>
/// The three shapes separate the two costs that <c>ReferenceNode.GetReference</c> used to
/// pay on every evaluation — parsing the address string, and allocating a <c>Reference</c>:
/// <list type="bullet">
/// <item>
/// <see cref="UniqueSameSheet"/> gives every row its own formula text, so
/// <c>ExpressionCache</c> hands out one AST per row and each node is resolved exactly once.
/// Nothing can be reused across evaluations, so this measures the address parsing alone.
/// </item>
/// <item>
/// <see cref="SharedSameSheet"/> repeats one formula text down the column, so a single AST
/// (and therefore a single <c>ReferenceNode</c>) is resolved once per row. This is the shape
/// a per-node memo can serve, on the sheet-less path.
/// </item>
/// <item>
/// <see cref="SharedCrossSheet"/> does the same through a sheet prefix, which has to resolve
/// the worksheet before building the address — the path where a memo has to be keyed on the
/// resolved sheet rather than cached outright.
/// </item>
/// </list>
/// Row count is deliberately far below the 250K of <see cref="XLiburReadBenchmarks"/>: this
/// benchmark is measuring per-reference cost, not load throughput, and a smaller sheet keeps
/// iterations short enough for BenchmarkDotNet to take a useful number of them.
/// </remarks>
[MemoryDiagnoser]
public class FormulaEvaluationBenchmarks
{
    private const int RowCount = 20_000;

    /// <summary>
    /// Cells summed by every variant's formula.
    /// </summary>
    /// <remarks>
    /// Held equal across all three on purpose. Until spec 19 area 5 task 5.1 it was not: the
    /// same-sheet variant summed <c>$D$1:$E$1</c>, two cells, while the cross-sheet one summed
    /// twenty. The 3.1x gap between them was therefore a gap between summing two cells and summing
    /// twenty as much as anything about resolving a sheet prefix, and could not be read as a
    /// cross-sheet penalty. With the count equal the only difference left between
    /// <see cref="SharedSameSheet"/> and <see cref="SharedCrossSheet"/> is the prefix.
    /// </remarks>
    private const int SummedCells = 20;

    // Assigned by Setup, so the field is never observed null despite the declaration.
    private XLWorkbook _uniqueSameSheet = null!;
    private XLWorkbook _sharedSameSheet = null!;
    private XLWorkbook _sharedCrossSheet = null!;

    [GlobalSetup]
    public void Setup()
    {
        SixLaborsV1FontBootstrap.Register();

        // Distinct formula text per row: one AST each, every reference resolved once. The window
        // slides so the text differs, and Build populates past RowCount so every window is full.
        _uniqueSameSheet = Build(out var uniqueSheet);
        for (var row = 1; row <= RowCount; row++)
            uniqueSheet.Cell(row, 6).FormulaA1 = $"SUM(D{row}:D{row + SummedCells - 1})";

        // One formula text for the whole column: one AST, resolved RowCount times.
        _sharedSameSheet = Build(out var sharedSheet);
        for (var row = 1; row <= RowCount; row++)
            sharedSheet.Cell(row, 6).FormulaA1 = $"SUM($D$1:$D${SummedCells})";

        // As above, but through a sheet prefix, so the worksheet must be resolved first.
        _sharedCrossSheet = Build(out var crossSheet);
        var lookup = _sharedCrossSheet.AddWorksheet("Lookup");
        for (var row = 1; row <= SummedCells; row++)
            lookup.Cell(row, 1).Value = row;

        for (var row = 1; row <= RowCount; row++)
            crossSheet.Cell(row, 6).FormulaA1 = $"SUM(Lookup!$A$1:$A${SummedCells})";
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        _uniqueSameSheet.Dispose();
        _sharedSameSheet.Dispose();
        _sharedCrossSheet.Dispose();
    }

    [Benchmark(Baseline = true)]
    public void UniqueSameSheet() => _uniqueSameSheet.RecalculateAllFormulas();

    [Benchmark]
    public void SharedSameSheet() => _sharedSameSheet.RecalculateAllFormulas();

    [Benchmark]
    public void SharedCrossSheet() => _sharedCrossSheet.RecalculateAllFormulas();

    /// <summary>
    /// A workbook with the two operand columns populated and no formulas yet; the caller adds
    /// the shape it wants to measure.
    /// </summary>
    private static XLWorkbook Build(out IXLWorksheet sheet)
    {
        var workbook = new XLWorkbook();
        sheet = workbook.AddWorksheet("Data");

        // Populated past RowCount so the sliding window in UniqueSameSheet never runs off the
        // populated area - a window of blanks would be cheaper to sum than one of numbers, which
        // would reintroduce the confound this fixture exists to remove.
        for (var row = 1; row <= RowCount + SummedCells; row++)
        {
            sheet.Cell(row, 4).Value = row;
            sheet.Cell(row, 5).Value = row * 2;
        }

        return workbook;
    }
}
