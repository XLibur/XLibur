using System;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;

namespace XLibur.Benchmarks;

[MemoryDiagnoser]
public class XLiburReadBenchmarks
{
    private const int RowCount = 250_000;
    private const int ColCount = 15;

    private byte[] _fileBytes = null!;

    [GlobalSetup]
    public void Setup()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Data");

#pragma warning disable S2245 // Deterministic seed for reproducible benchmarks
        var random = new Random(42);
#pragma warning restore S2245
        var baseDate = new DateTime(2020, 1, 1);
        string[] regions = { "North", "South", "East", "West", "Central" };
        string[] statuses = { "Active", "Pending", "Closed", "Review", "Draft" };

        for (var i = 0; i < RowCount; i++)
        {
            var row = i + 1;
            var seed = random.Next(10000);

            // Cols 1-3: Strings
            ws.Cell(row, 1).Value = $"Item {i}-{seed}";
            ws.Cell(row, 2).Value = $"Cat-{i % 12}";
            ws.Cell(row, 3).Value = regions[i % regions.Length];

            // Cols 4-8: Numbers
            ws.Cell(row, 4).Value = Math.Round(random.NextDouble() * 10000, 2);
            ws.Cell(row, 5).Value = random.Next(1, 5000);
            ws.Cell(row, 6).Value = Math.Round(random.NextDouble(), 4);
            ws.Cell(row, 7).Value = Math.Round(random.NextDouble() * 1000, 2);
            ws.Cell(row, 8).Value = random.Next(0, 100);

            // Cols 9-11: Dates
            ws.Cell(row, 9).Value = baseDate.AddDays(random.Next(0, 2000));
            ws.Cell(row, 10).Value = baseDate.AddDays(random.Next(0, 2000));
            ws.Cell(row, 11).Value = baseDate.AddDays(random.Next(0, 2000));

            // Cols 12-14: More strings
            ws.Cell(row, 12).Value = statuses[i % statuses.Length];
            ws.Cell(row, 13).Value = $"Note for row {row} with seed {seed}";
            ws.Cell(row, 14).Value = $"CODE-{seed:D5}";

            // Col 15: SUM formula referencing cols 4-8
            ws.Cell(row, 15).FormulaA1 = $"SUM(D{row}:H{row})";
        }

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        _fileBytes = ms.ToArray();
    }

    [Benchmark]
    public void LoadWorkbook()
    {
        using var workbook = new XLWorkbook(new MemoryStream(_fileBytes));
    }

    [Benchmark]
    public void LoadAndReadAllCells()
    {
        using var workbook = new XLWorkbook(new MemoryStream(_fileBytes));
        var ws = workbook.Worksheets.First();

        for (var row = 1; row <= RowCount; row++)
        {
            for (var col = 1; col <= ColCount; col++)
            {
                _ = ws.Cell(row, col).GetValue<string>();
            }
        }
    }
}
