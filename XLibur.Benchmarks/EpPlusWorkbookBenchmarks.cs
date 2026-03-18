using System;
using System.Drawing;
using System.IO;
using BenchmarkDotNet.Attributes;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace XLibur.Benchmarks;

[MemoryDiagnoser]
public class EpPlusWorkbookBenchmarks
{
    private const int RowCount = 50_000;

    private string[] _strings = null!;
    private double[] _numbers = null!;
    private DateTime[] _dates = null!;

    [GlobalSetup]
    public void Setup()
    {
        ExcelPackage.License.SetNonCommercialOrganization("XLibur Benchmarks");

        _strings = new string[RowCount];
        _numbers = new double[RowCount];
        _dates = new DateTime[RowCount];

#pragma warning disable S2245 // Deterministic seed for reproducible benchmarks
        var random = new Random(42);
#pragma warning restore S2245
        var baseDate = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);

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
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("Data");

        worksheet.Cells[1, 1].Value = "Name";
        worksheet.Cells[1, 2].Value = "Amount";
        worksheet.Cells[1, 3].Value = "Date";

        for (var i = 0; i < RowCount; i++)
        {
            var row = i + 2;
            worksheet.Cells[row, 1].Value = _strings[i];
            worksheet.Cells[row, 2].Value = _numbers[i];
            worksheet.Cells[row, 3].Value = _dates[i];
        }

        using var stream = new MemoryStream();
        package.SaveAs(stream);
    }

    [Benchmark]
    public void CreateFormattedAndSave()
    {
        using var package = new ExcelPackage();
        var ws = package.Workbook.Worksheets.Add("Formatted");

        WriteHeaders(ws);

        for (var i = 0; i < RowCount; i++)
        {
            var row = i + 2;
            var idx = i % _strings.Length;

            PopulateRow(ws, row, i, idx);

            if (i % 2 == 0)
                FormatEvenRow(ws, row, i);
        }

        using var stream = new MemoryStream();
        package.SaveAs(stream);
    }

    private static void WriteHeaders(ExcelWorksheet ws)
    {
        var headers = new[]
            { "Name", "Amount", "Date", "Quantity", "Price", "Total", "Status", "Category", "Region", "Notes" };
        for (var c = 1; c <= 10; c++)
        {
            var cell = ws.Cells[1, c];
            cell.Value = headers[c - 1];
            cell.Style.Font.Bold = true;
            cell.Style.Font.Color.SetColor(Color.White);
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Double;
            cell.Style.Border.Bottom.Color.SetColor(Color.Black);
        }
    }

    private void PopulateRow(ExcelWorksheet ws, int row, int i, int idx)
    {
        ws.Cells[row, 1].Value = _strings[idx];
        ws.Cells[row, 2].Value = _numbers[idx];
        ws.Cells[row, 3].Value = _dates[idx];
        ws.Cells[row, 4].Value = (i % 500) + 1;
        ws.Cells[row, 5].Value = _numbers[idx] * 0.1;
        ws.Cells[row, 6].Value = _numbers[idx] * ((i % 500) + 1) * 0.1;
        ws.Cells[row, 7].Value = GetStatus(i);
        ws.Cells[row, 8].Value = $"Cat-{(i % 12) + 1}";
        ws.Cells[row, 9].Value = GetRegion(i);
        ws.Cells[row, 10].Value = $"Note for row {row}";
    }

    private static string GetStatus(int i) => (i % 3) switch
    {
        0 => "Active",
        1 => "Pending",
        _ => "Closed"
    };

    private static string GetRegion(int i) => (i % 5) switch
    {
        0 => "North",
        1 => "South",
        2 => "East",
        3 => "West",
        _ => "Central"
    };

    private static void FormatEvenRow(ExcelWorksheet ws, int row, int i)
    {
        FormatNameColumn(ws.Cells[row, 1]);
        FormatAmountColumn(ws.Cells[row, 2]);
        FormatDateColumn(ws.Cells[row, 3]);
        FormatQuantityColumn(ws.Cells[row, 4]);
        FormatPriceColumn(ws.Cells[row, 5]);
        FormatTotalColumn(ws.Cells[row, 6]);
        FormatStatusColumn(ws.Cells[row, 7], i);
        FormatCategoryColumn(ws.Cells[row, 8]);
        FormatRegionColumn(ws.Cells[row, 9]);
        FormatNotesColumn(ws.Cells[row, 10]);
    }

    private static void FormatNameColumn(ExcelRange cell)
    {
        cell.Style.Font.Bold = true;
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
    }

    private static void FormatAmountColumn(ExcelRange cell)
    {
        cell.Style.Numberformat.Format = "#,##0.00";
        cell.Style.Font.Color.SetColor(Color.DarkRed);
        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Gray);
    }

    private static void FormatDateColumn(ExcelRange cell)
    {
        cell.Style.Numberformat.Format = "d-mmm-yy";
        cell.Style.Font.Italic = true;
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
    }

    private static void FormatQuantityColumn(ExcelRange cell)
    {
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        cell.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.DarkGoldenrod);
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
    }

    private static void FormatPriceColumn(ExcelRange cell)
    {
        cell.Style.Numberformat.Format = "$ #,##0.00";
        cell.Style.Font.Name = "Consolas";
        cell.Style.Font.Size = 10;
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(Color.LightCoral);
    }

    private static void FormatTotalColumn(ExcelRange cell)
    {
        cell.Style.Numberformat.Format = "_($* #,##0.00_)";
        cell.Style.Font.Bold = true;
        cell.Style.Border.Bottom.Style = ExcelBorderStyle.Dashed;
        cell.Style.Border.Bottom.Color.SetColor(Color.Navy);
    }

    private static void FormatStatusColumn(ExcelRange cell, int i)
    {
        cell.Style.Font.Strike = i % 3 == 2;
        cell.Style.Font.Color.SetColor((i % 3) switch
        {
            0 => Color.Green,
            1 => Color.Orange,
            _ => Color.Red
        });
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(Color.Lavender);
    }

    private static void FormatCategoryColumn(ExcelRange cell)
    {
        cell.Style.Font.UnderLine = true;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        cell.Style.Border.Right.Style = ExcelBorderStyle.Dotted;
        cell.Style.Border.Right.Color.SetColor(Color.Purple);
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
    }

    private static void FormatRegionColumn(ExcelRange cell)
    {
        cell.Style.Font.Name = "Georgia";
        cell.Style.Font.Size = 12;
        cell.Style.Font.Color.SetColor(Color.Teal);
        cell.Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Teal);
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(Color.Wheat);
    }

    private static void FormatNotesColumn(ExcelRange cell)
    {
        cell.Style.WrapText = true;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
        cell.Style.Border.BorderAround(ExcelBorderStyle.Double, Color.Chocolate);
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(Color.LightSalmon);
        cell.Style.Font.Color.SetColor(Color.DarkSlateGray);
    }
}
