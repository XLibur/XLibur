using System;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using OfficeOpenXml;

namespace XLibur.Benchmarks;

[MemoryDiagnoser]
public class EpPlusReadBenchmarks
{
    private const int RowCount = 250_000;
    private const int ColCount = 15;

    private byte[] _fileBytes = null!;

    [GlobalSetup]
    public void Setup()
    {
        ExcelPackage.License.SetNonCommercialOrganization("XLibur Benchmarks");

        using var package = new ExcelPackage();
        var ws = package.Workbook.Worksheets.Add("Data");

#pragma warning disable S2245 // Deterministic seed for reproducible benchmarks
        var random = new Random(42);
#pragma warning restore S2245
        var baseDate = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        string[] regions = { "North", "South", "East", "West", "Central" };
        string[] statuses = { "Active", "Pending", "Closed", "Review", "Draft" };

        for (var i = 0; i < RowCount; i++)
        {
            var row = i + 1;
            var seed = random.Next(10000);

            // Cols 1-3: Strings
            ws.Cells[row, 1].Value = $"Item {i}-{seed}";
            ws.Cells[row, 2].Value = $"Cat-{i % 12}";
            ws.Cells[row, 3].Value = regions[i % regions.Length];

            // Cols 4-8: Numbers
            ws.Cells[row, 4].Value = Math.Round(random.NextDouble() * 10000, 2);
            ws.Cells[row, 5].Value = random.Next(1, 5000);
            ws.Cells[row, 6].Value = Math.Round(random.NextDouble(), 4);
            ws.Cells[row, 7].Value = Math.Round(random.NextDouble() * 1000, 2);
            ws.Cells[row, 8].Value = random.Next(0, 100);

            // Cols 9-11: Dates
            ws.Cells[row, 9].Value = baseDate.AddDays(random.Next(0, 2000));
            ws.Cells[row, 10].Value = baseDate.AddDays(random.Next(0, 2000));
            ws.Cells[row, 11].Value = baseDate.AddDays(random.Next(0, 2000));

            // Cols 12-14: More strings
            ws.Cells[row, 12].Value = statuses[i % statuses.Length];
            ws.Cells[row, 13].Value = $"Note for row {row} with seed {seed}";
            ws.Cells[row, 14].Value = $"CODE-{seed:D5}";

            // Col 15: SUM formula referencing cols 4-8
            ws.Cells[row, 15].Formula = $"SUM(D{row}:H{row})";
        }

        using var ms = new MemoryStream();
        package.SaveAs(ms);
        _fileBytes = ms.ToArray();
    }

    [Benchmark]
    public void LoadWorkbook()
    {
        using var package = new ExcelPackage(new MemoryStream(_fileBytes));
    }

    [Benchmark]
    public void LoadAndReadAllCells()
    {
        using var package = new ExcelPackage(new MemoryStream(_fileBytes));
        var ws = package.Workbook.Worksheets.First();

        for (var row = 1; row <= RowCount; row++)
        {
            for (var col = 1; col <= ColCount; col++)
            {
                _ = ws.Cells[row, col].GetValue<string>();
            }
        }
    }
}
