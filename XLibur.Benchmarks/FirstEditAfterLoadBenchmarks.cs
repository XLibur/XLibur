using System;
using System.IO;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// The cost of the first edit after a load, in a workbook whose formulas all have cached values.
/// </summary>
/// <remarks>
/// <para>
/// A load leaves a formula with a cached value clean. Only the dependency tree knows which formulas
/// an edit reaches, so since #504 the first edit after a load builds it, which parses every formula
/// in the workbook. Before #504 that edit built nothing and marked nothing dirty, which is why every
/// formula that read the edited cell kept the value from the file.
/// </para>
/// <para>
/// <see cref="Load"/> is the floor, and the gap from it to <see cref="LoadAndEditOneCell"/> is the
/// price of the first edit. <see cref="LoadEditOneCellAndSave"/> is the template-fill round trip,
/// where the first edit is paid once per request. Every row has its own formula text, as a shared
/// formula loads, so each is a separate parse: the expensive case.
/// </para>
/// </remarks>
[MemoryDiagnoser]
public class FirstEditAfterLoadBenchmarks
{
    /// <summary>
    /// Rows × formulas per row. The two narrow shapes are the ones #504 measured. The wide shape is
    /// the 200,000-formula row of #513, and <c>shared</c> saves each of its formula columns as one
    /// shared formula, as Excel does (see <see cref="FirstEditFixture"/>).
    /// </summary>
    [Params("1000x2", "10000x2", "25000x8", "25000x8shared")]
    public string Shape { get; set; } = "";

    private byte[] _package = null!;

    [GlobalSetup]
    public void Setup()
    {
        SixLaborsV1FontBootstrap.Register();

        const string sharedSuffix = "shared";
        var shared = Shape.EndsWith(sharedSuffix, StringComparison.Ordinal);
        var parts = (shared ? Shape[..^sharedSuffix.Length] : Shape).Split('x');
        _package = FirstEditFixture.Build(int.Parse(parts[0]), int.Parse(parts[1]), shared);
    }

    [Benchmark(Baseline = true)]
    public int Load()
    {
        using var input = new MemoryStream(_package, writable: false);
        using var workbook = new XLWorkbook(input);
        return workbook.Worksheets.Count;
    }

    [Benchmark]
    public int LoadAndEditOneCell()
    {
        using var input = new MemoryStream(_package, writable: false);
        using var workbook = new XLWorkbook(input);
        workbook.Worksheet(1).Cell(1, 1).Value = 5;
        return workbook.Worksheets.Count;
    }

    [Benchmark]
    public long LoadEditOneCellAndSave()
    {
        using var input = new MemoryStream(_package, writable: false);
        using var workbook = new XLWorkbook(input);
        workbook.Worksheet(1).Cell(1, 1).Value = 5;
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.Length;
    }
}
