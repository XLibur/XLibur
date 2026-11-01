using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.PinataLabs;

namespace XLibur.Benchmarks;

/// <summary>
/// Spec 04 (#841). Reading a few dirty formulas out of many. Each formula in column C reads a
/// formula in column B, which reads a plain value in column A, so the closure of one read is two
/// formulas whatever the size of the workbook.
/// </summary>
/// <remarks>
/// <see cref="ReadOnDemand"/> calculates only those closures. <see cref="ReadAfterFullPass"/> is
/// what a read of a dirty formula that reads another dirty formula cost before: a recalculation of
/// every formula, and with it the dependency tree and the calculation chain. The workbook is built
/// again for each invocation, because a read leaves the formulas it calculated clean.
/// </remarks>
[MemoryDiagnoser]
[SimpleJob(invocationCount: 1, iterationCount: 5, warmupCount: 1)]
public class DemandDrivenReadBenchmarks
{
    private const int Reads = 100;

    private const int Formulas = 100_000;

    private XLWorkbook _workbook = null!;
    private IXLWorksheet _sheet = null!;

    [IterationSetup]
    public void Setup()
    {
        PinataLabsFontBootstrap.Register();

        _workbook = new XLWorkbook();
        _sheet = _workbook.AddWorksheet("Data");
        for (var row = 1; row <= Formulas; row++)
        {
            _sheet.Cell(row, 1).Value = row;
            _sheet.Cell(row, 2).FormulaA1 = $"A{row}+1";
            _sheet.Cell(row, 3).FormulaA1 = $"B{row}*2";
        }
    }

    [IterationCleanup]
    public void Cleanup() => _workbook.Dispose();

    [Benchmark]
    public double ReadOnDemand() => ReadSpread();

    [Benchmark(Baseline = true)]
    public double ReadAfterFullPass()
    {
        _workbook.RecalculateAllFormulas();
        return ReadSpread();
    }

    private double ReadSpread()
    {
        var sum = 0d;
        var step = Formulas / Reads;
        for (var i = 0; i < Reads; i++)
            sum += _sheet.GetValue<double>(1 + i * step, 3);

        return sum;
    }
}
