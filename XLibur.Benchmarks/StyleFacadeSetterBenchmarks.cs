using System;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;

namespace XLibur.Benchmarks;

/// <summary>
/// The per-setter cost of the style facades (<c>XLFont</c>, <c>XLFill</c>, <c>XLAlignment</c>,
/// <c>XLNumberFormat</c>, <c>XLProtection</c>, <c>XLBorder</c>), on both of their paths.
///
/// Run with:
/// dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- --filter '*StyleFacadeSetter*'
/// </summary>
/// <remarks>
/// Written for #621 finding 2. Every facade setter has two paths: the cell fast path, which hands
/// the new component key straight to the style, and the <c>Modify</c> path that ranges, rows,
/// columns and worksheets take. <see cref="BulkStyleBenchmarks"/> styles one large range, so its
/// cost is the per-cell write and the setter itself is lost in it. Here each write touches a few
/// cells, so what is left is mostly the setter.
/// <para>
/// No font engine is registered here: <c>Program</c> registers one before any benchmark runs, and
/// nothing in these writes measures text.
/// </para>
/// </remarks>
[MemoryDiagnoser]
public class StyleFacadeSetterBenchmarks
{
    private const int Rows = 20_000;

    private XLWorkbook _workbook = null!;
    private IXLWorksheet _worksheet = null!;

    [IterationSetup]
    public void IterationSetup()
    {
        _workbook = new XLWorkbook();
        _worksheet = _workbook.AddWorksheet("Data");
    }

    /// <summary>
    /// Disposes the workbook and collects it outside the measured region, as
    /// <see cref="CellStylingBenchmarks"/> does.
    /// </summary>
    [IterationCleanup]
    public void IterationCleanup()
    {
        _workbook.Dispose();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true);
    }

    /// <summary>One property from each facade, on each cell: the cell fast path.</summary>
    [Benchmark(Baseline = true)]
    public void CellSetters()
    {
        for (var r = 1; r <= Rows; r++)
            SetEachFacade(_worksheet.Cell(r, 1).Style);
    }

    /// <summary>
    /// The same properties on a two-cell range per row: the <c>Modify</c> path, where each setter
    /// builds a closure and the facade used to intern its component a second time.
    /// </summary>
    [Benchmark]
    public void SmallRangeSetters()
    {
        for (var r = 1; r <= Rows; r++)
            SetEachFacade(_worksheet.Range(r, 1, r, 2).Style);
    }

    private static void SetEachFacade(IXLStyle style)
    {
        style.Font.Bold = true;
        style.Font.FontSize = 12;
        style.Fill.PatternColor = XLColor.Red;
        style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        style.NumberFormat.NumberFormatId = 14;
        style.Protection.Hidden = true;
        style.Border.DiagonalUp = true;
    }
}
