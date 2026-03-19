using System;
using System.IO;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Diagnostics.dotTrace;
using XLibur.Excel;

namespace XLibur.Benchmarks;

[MemoryDiagnoser]
//[DotMemoryDiagnoser]
[DotTraceDiagnoser]
[Config(typeof(JoinSummaryConfig))]
public class XLiburWorkbookBenchmarks
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

        WriteHeaders(ws);

        for (var i = 0; i < RowCount; i++)
        {
            var row = i + 2;
            var idx = i % _strings.Length;

            WriteRowData(ws, row, i, idx);

            if (i % 2 == 0)
                ApplyRowFormatting(ws, row, i);
        }

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
    }

    private static void WriteHeaders(IXLWorksheet ws)
    {
        var headers = new[]
            { "Name", "Amount", "Date", "Quantity", "Price", "Total", "Status", "Category", "Region", "Notes" };
        for (var c = 1; c <= 10; c++)
        {
            var hdr = ws.Cell(1, c);
            hdr.Value = headers[c - 1];
            hdr.Style.Font.Bold = true;
            hdr.Style.Font.FontColor = XLColor.White;
            hdr.Style.Fill.BackgroundColor = XLColor.DarkBlue;
            hdr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            hdr.Style.Border.BottomBorder = XLBorderStyleValues.Double;
            hdr.Style.Border.BottomBorderColor = XLColor.Black;
        }
    }

    private void WriteRowData(IXLWorksheet ws, int row, int i, int idx)
    {
        ws.Cell(row, 1).Value = _strings[idx];
        ws.Cell(row, 2).Value = _numbers[idx];
        ws.Cell(row, 3).Value = _dates[idx];
        ws.Cell(row, 4).Value = (i % 500) + 1;
        ws.Cell(row, 5).Value = _numbers[idx] * 0.1;
        ws.Cell(row, 6).Value = _numbers[idx] * ((i % 500) + 1) * 0.1;
        var status = (i % 3) switch { 0 => "Active", 1 => "Pending", _ => "Closed" };
        ws.Cell(row, 7).Value = status;
        ws.Cell(row, 8).Value = $"Cat-{(i % 12) + 1}";
        var region = (i % 5) switch { 0 => "North", 1 => "South", 2 => "East", 3 => "West", _ => "Central" };
        ws.Cell(row, 9).Value = region;
        ws.Cell(row, 10).Value = $"Note for row {row}";
    }

    private static void ApplyRowFormatting(IXLWorksheet ws, int row, int i)
    {
        ApplyCellStyle(ws, row, 1, s =>
        {
            s.Font.Bold = true;
            s.Fill.BackgroundColor = XLColor.LightBlue;
        });

        ApplyCellStyle(ws, row, 2, s =>
        {
            s.NumberFormat.Format = "#,##0.00";
            s.Font.FontColor = XLColor.DarkRed;
            s.Border.OutsideBorder = XLBorderStyleValues.Thin;
            s.Border.OutsideBorderColor = XLColor.Gray;
        });

        ApplyCellStyle(ws, row, 3, s =>
        {
            s.NumberFormat.NumberFormatId = 15;
            s.Font.Italic = true;
            s.Fill.BackgroundColor = XLColor.LightGreen;
        });

        ApplyCellStyle(ws, row, 4, s =>
        {
            s.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            s.Border.OutsideBorder = XLBorderStyleValues.Medium;
            s.Border.OutsideBorderColor = XLColor.DarkGoldenrod;
            s.Fill.BackgroundColor = XLColor.LightYellow;
        });

        ApplyCellStyle(ws, row, 5, s =>
        {
            s.NumberFormat.Format = "$ #,##0.00";
            s.Font.FontName = "Consolas";
            s.Font.FontSize = 10;
            s.Fill.BackgroundColor = XLColor.LightCoral;
        });

        ApplyCellStyle(ws, row, 6, s =>
        {
            s.NumberFormat.Format = "_($* #,##0.00_)";
            s.Font.Bold = true;
            s.Border.BottomBorder = XLBorderStyleValues.Dashed;
            s.Border.BottomBorderColor = XLColor.Navy;
        });

        ApplyCellStyle(ws, row, 7, s =>
        {
            s.Font.Strikethrough = i % 3 == 2;
            s.Font.FontColor = (i % 3) switch { 0 => XLColor.Green, 1 => XLColor.Orange, _ => XLColor.Red };
            s.Fill.BackgroundColor = XLColor.Lavender;
        });

        ApplyCellStyle(ws, row, 8, s =>
        {
            s.Font.Underline = XLFontUnderlineValues.Single;
            s.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            s.Border.RightBorder = XLBorderStyleValues.Dotted;
            s.Border.RightBorderColor = XLColor.Purple;
            s.Fill.BackgroundColor = XLColor.LightGray;
        });

        ApplyCellStyle(ws, row, 9, s =>
        {
            s.Font.FontName = "Georgia";
            s.Font.FontSize = 12;
            s.Font.FontColor = XLColor.Teal;
            s.Border.OutsideBorder = XLBorderStyleValues.Thick;
            s.Border.OutsideBorderColor = XLColor.Teal;
            s.Fill.BackgroundColor = XLColor.Wheat;
        });

        ApplyCellStyle(ws, row, 10, s =>
        {
            s.Alignment.WrapText = true;
            s.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            s.Border.OutsideBorder = XLBorderStyleValues.Double;
            s.Border.OutsideBorderColor = XLColor.Chocolate;
            s.Fill.BackgroundColor = XLColor.LightSalmon;
            s.Font.FontColor = XLColor.DarkSlateGray;
        });
    }

    private static void ApplyCellStyle(IXLWorksheet ws, int row, int col, Action<IXLStyle> apply)
    {
        apply(ws.Cell(row, col).Style);
    }
}
