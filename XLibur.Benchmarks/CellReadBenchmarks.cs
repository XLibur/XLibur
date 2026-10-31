using System;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.PinataLabs;

namespace XLibur.Benchmarks;

/// <summary>
/// Attributes the per-call cost of reading one cell by coordinate, layer by layer, so the 166 B
/// that <c>XLiburReadBenchmarks.LoadAndReadAllCells</c> pays per <c>ws.Cell(r,c).GetValue&lt;string&gt;()</c>
/// is split into the wrapper, the value fetch and the conversion.
///
/// Run with:
/// dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- --filter '*CellReadBenchmarks*'
/// </summary>
/// <remarks>
/// Spec 19 area 1, task 1.5 (#842). The fixture has the same column mix as <c>XLiburReadBenchmarks</c>
/// (3 text, 5 numeric, 3 date, 3 text, 1 formula).
/// <para>
/// The rungs are <b>not additive</b>: they do different work over different cell sets, so read each as
/// its own figure. Compare like with like only:
/// <c>Cell_GetValueString</c> with <c>Seam_GetValueString</c> (all 15 columns),
/// <c>Cell_GetValueDouble</c> with <c>Seam_GetValueDouble</c> (the 5 numeric columns), and
/// <c>Cell_GetValueStringTextOnly</c> with <c>Seam_GetValueStringTextOnly</c> (the 6 text columns).
/// <c>Cell_Discarded</c> and <c>Cell_Value</c> price the wrapper and the raw value fetch.
/// Divide by the calls per operation (<c>Rows x Cols</c>, or <c>Rows x 5</c>, or <c>Rows x 6</c>).
/// </para>
/// </remarks>
[MemoryDiagnoser]
public class CellReadBenchmarks
{
    private const int Rows = 20_000;
    private const int Cols = 15;

    public const int CallsPerOp = Rows * Cols;

    private XLWorkbook _workbook = null!;
    private XLWorksheet _sheet = null!;

    // The public member, reached the way a caller reaches it: through the interface.
    private IXLWorksheet _worksheet = null!;

    [GlobalSetup]
    public void Setup()
    {
        PinataLabsFontBootstrap.Register();
        _workbook = new XLWorkbook();
        _sheet = (XLWorksheet)_workbook.AddWorksheet("Data");
        _worksheet = _sheet;

#pragma warning disable S2245 // Deterministic seed for reproducible benchmarks
        var random = new Random(42);
#pragma warning restore S2245
        var baseDate = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        string[] regions = { "North", "South", "East", "West", "Central" };

        for (var i = 0; i < Rows; i++)
        {
            var row = i + 1;
            var seed = random.Next(10000);
            _sheet.Cell(row, 1).Value = $"Item {i}-{seed}";
            _sheet.Cell(row, 2).Value = $"Cat-{i % 12}";
            _sheet.Cell(row, 3).Value = regions[i % regions.Length];
            _sheet.Cell(row, 4).Value = Math.Round(random.NextDouble() * 10000, 2);
            _sheet.Cell(row, 5).Value = random.Next(1, 5000);
            _sheet.Cell(row, 6).Value = Math.Round(random.NextDouble(), 4);
            _sheet.Cell(row, 7).Value = Math.Round(random.NextDouble() * 1000, 2);
            _sheet.Cell(row, 8).Value = random.Next(0, 100);
            _sheet.Cell(row, 9).Value = baseDate.AddDays(random.Next(0, 2000));
            _sheet.Cell(row, 10).Value = baseDate.AddDays(random.Next(0, 2000));
            _sheet.Cell(row, 11).Value = baseDate.AddDays(random.Next(0, 2000));
            _sheet.Cell(row, 12).Value = regions[i % regions.Length];
            _sheet.Cell(row, 13).Value = $"Note for row {row} with seed {seed}";
            _sheet.Cell(row, 14).Value = $"CODE-{seed:D5}";
            _sheet.Cell(row, 15).FormulaA1 = $"SUM(D{row}:H{row})";
        }

        // Settle every formula so the read is a pure read, as in a loaded workbook.
        _workbook.RecalculateAllFormulas();
    }

    [GlobalCleanup]
    public void Cleanup() => _workbook.Dispose();

    /// <summary>The floor of the wrapper: <c>ws.Cell(r,c)</c> discarded.</summary>
    [Benchmark(Baseline = true)]
    public int Cell_Discarded()
    {
        var n = 0;
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
                n += _sheet.Cell(r, c).SheetPoint.Row;
        }

        return n;
    }

    /// <summary>+ the value fetch: <c>ws.Cell(r,c).Value</c>, no conversion.</summary>
    [Benchmark]
    public int Cell_Value()
    {
        var n = 0;
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
                n += (int)_sheet.Cell(r, c).Value.Type;
        }

        return n;
    }

    /// <summary>The benchmark that <c>LoadAndReadAllCells</c> runs: <c>GetValue&lt;string&gt;()</c> on every cell.</summary>
    [Benchmark]
    public int Cell_GetValueString()
    {
        var n = 0;
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
                n += ((IXLCell)_sheet.Cell(r, c)).GetValue<string>().Length;
        }

        return n;
    }

    /// <summary>The wrapper-free seam, same work as <see cref="Cell_GetValueString"/>.</summary>
    [Benchmark]
    public int Seam_GetValueString()
    {
        var n = 0;
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= Cols; c++)
                n += _worksheet.GetValue<string>(r, c).Length;
        }

        return n;
    }

    /// <summary>The wrapper-free seam, same work as <see cref="Cell_GetValueDouble"/>.</summary>
    [Benchmark]
    public double Seam_GetValueDouble()
    {
        double n = 0;
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 4; c <= 8; c++)
                n += _worksheet.GetValue<double>(r, c);
        }

        return n;
    }

    /// <summary><c>GetValue&lt;double&gt;()</c> on the five numeric columns only: no string is produced.</summary>
    [Benchmark]
    public double Cell_GetValueDouble()
    {
        double n = 0;
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 4; c <= 8; c++)
                n += ((IXLCell)_sheet.Cell(r, c)).GetValue<double>();
        }

        return n;
    }

    /// <summary>The public seam over the same six text columns as <see cref="Cell_GetValueStringTextOnly"/>.</summary>
    [Benchmark]
    public int Seam_GetValueStringTextOnly()
    {
        var n = 0;
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= 3; c++)
                n += _worksheet.GetValue<string>(r, c).Length;
            for (var c = 12; c <= 14; c++)
                n += _worksheet.GetValue<string>(r, c).Length;
        }

        return n;
    }

    /// <summary><c>GetValue&lt;string&gt;()</c> on the text columns only (1-3, 12-14): the string already exists.</summary>
    [Benchmark]
    public int Cell_GetValueStringTextOnly()
    {
        var n = 0;
        for (var r = 1; r <= Rows; r++)
        {
            for (var c = 1; c <= 3; c++)
                n += ((IXLCell)_sheet.Cell(r, c)).GetValue<string>().Length;
            for (var c = 12; c <= 14; c++)
                n += ((IXLCell)_sheet.Cell(r, c)).GetValue<string>().Length;
        }

        return n;
    }
}
