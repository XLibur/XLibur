using System;
using System.Linq;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Jobs;
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
[BenchmarkCategory("CI")]
[Config(typeof(DemandDrivenReadConfig))]
public class DemandDrivenReadBenchmarks
{
    /// <summary>
    /// <c>[SimpleJob(invocationCount: 1, iterationCount: 5, warmupCount: 1)]</c>, unless the command
    /// line chooses a job. A job from an attribute is run in addition to the one from <c>--job</c>, so
    /// the CI run (<c>--job short --inProcess</c>) would build and run this class out of process as
    /// well, and record two results for each benchmark.
    /// </summary>
    private sealed class DemandDrivenReadConfig : ManualConfig
    {
        public DemandDrivenReadConfig()
        {
            var chosenOnCommandLine = Environment.GetCommandLineArgs()
                .Any(a => a.StartsWith("--job", StringComparison.OrdinalIgnoreCase) || a.Equals("-j", StringComparison.Ordinal));
            if (!chosenOnCommandLine)
                AddJob(Job.Default.WithInvocationCount(1).WithIterationCount(5).WithWarmupCount(1));
        }
    }

    private const int Reads = 100;

    // 20K rows (60K cells) keeps the CI run lean (spec 57); the on-demand/full-pass gap is the same shape.
    private const int Formulas = 20_000;

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
