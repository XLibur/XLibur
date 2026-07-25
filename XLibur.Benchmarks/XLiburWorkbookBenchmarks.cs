using System;
using System.IO;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Diagnostics.dotTrace;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

[MemoryDiagnoser]
//[DotMemoryDiagnoser]
[DotTraceDiagnoser]
[Config(typeof(JoinSummaryConfig))]
public class XLiburWorkbookBenchmarks
{
    private const int RowCount = 50_000;

    private BenchmarkData _data = null!;
    private string[] _strings = null!;
    private double[] _numbers = null!;
    private DateTime[] _dates = null!;

    [GlobalSetup]
    public void Setup()
    {
        SixLaborsV1FontBootstrap.Register();
        _data = BenchmarkData.Create(RowCount);
        _strings = _data.Strings;
        _numbers = _data.Numbers;
        _dates = _data.Dates;
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

        var sumRow = RowCount + 2;
        worksheet.Cell(sumRow, 1).Value = "Total";
        worksheet.Cell(sumRow, 2).FormulaA1 = $"SUM(B2:B{RowCount + 1})";

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
    }

    [Benchmark]
    public void CreateFormattedAndSave()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Formatted");

        FormattedSheetBuilder.WriteHeaders(ws);

        for (var i = 0; i < RowCount; i++)
        {
            var row = i + 2;
            var idx = i % _strings.Length;

            FormattedSheetBuilder.WriteRowData(ws, _data, row, i, idx);

            if (i % 2 == 0)
                FormattedSheetBuilder.ApplyRowFormatting(ws, row, i);
        }

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
    }
}
