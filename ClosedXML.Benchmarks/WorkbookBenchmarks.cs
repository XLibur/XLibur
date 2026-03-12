using System;
using System.IO;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Diagnostics.dotTrace;
using ClosedXML.Excel;

namespace ClosedXML.Benchmarks;

[MemoryDiagnoser]
[DotTraceDiagnoser]
public class WorkbookBenchmarks
{
    private const int RowCount = 50_000;

    private string[] _strings = null!;
    private double[] _numbers = null!;
    private DateTime[] _dates = null!;

    [GlobalSetup]
    public void Setup()
    {
        _strings = new string[RowCount];
        _numbers = new double[RowCount];
        _dates = new DateTime[RowCount];

        var random = new Random(42);
        var baseDate = new DateTime(2020, 1, 1);

        for (var i = 0; i < RowCount; i++)
        {
            _strings[i] = $"Item {i} - {random.Next(1000):D4}";
            _numbers[i] = Math.Round(random.NextDouble() * 10000, 2);
            _dates[i] = baseDate.AddDays(random.Next(0, 1500));
        }
    }

    [Benchmark]
    public void CreateAndSave()
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Data");

        worksheet.Cell(1, 1).Value = "Name";
        worksheet.Cell(1, 2).Value = "Amount";
        worksheet.Cell(1, 3).Value = "Date";

        for (var i = 0; i < RowCount; i++)
        {
            var row = i + 2;
            worksheet.Cell(row, 1).Value = _strings[i];
            worksheet.Cell(row, 2).Value = _numbers[i];
            worksheet.Cell(row, 3).Value = _dates[i];
        }

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
    }
}
