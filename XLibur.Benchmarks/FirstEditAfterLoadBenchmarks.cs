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
    /// <summary>Rows of the fixture, each with two formulas.</summary>
    [Params(1_000, 10_000)]
    public int Rows { get; set; }

    private byte[] _package = null!;

    [GlobalSetup]
    public void Setup()
    {
        SixLaborsV1FontBootstrap.Register();

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Data");
        for (var row = 1; row <= Rows; row++)
        {
            sheet.Cell(row, 1).Value = row;
            sheet.Cell(row, 2).Value = row * 2;
            sheet.Cell(row, 3).FormulaA1 = $"A{row}*B{row}+1";
            sheet.Cell(row, 4).FormulaA1 = $"IF(C{row}>100,SUM(A{row}:C{row}),C{row}/2)";
        }

        // Calculated, so that each formula is saved with a cached value and loads clean.
        workbook.RecalculateAllFormulas();

        using var buffer = new MemoryStream();
        workbook.SaveAs(buffer);
        _package = buffer.ToArray();
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
