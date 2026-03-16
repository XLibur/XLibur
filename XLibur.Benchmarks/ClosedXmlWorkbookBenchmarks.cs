using System;
using System.IO;
using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;

namespace XLibur.Benchmarks;

[MemoryDiagnoser]
public class ClosedXmlWorkbookBenchmarks
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

#pragma warning disable S2245 // Using pseudorandom number generator - deterministic seed is intentional for benchmarks
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

            PopulateRow(ws, row, i, idx);

            if (i % 2 == 0)
                FormatEvenRow(ws, row, i);
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

    private void PopulateRow(IXLWorksheet ws, int row, int i, int idx)
    {
        ws.Cell(row, 1).Value = _strings[idx];
        ws.Cell(row, 2).Value = _numbers[idx];
        ws.Cell(row, 3).Value = _dates[idx];
        ws.Cell(row, 4).Value = (i % 500) + 1;
        ws.Cell(row, 5).Value = _numbers[idx] * 0.1;
        ws.Cell(row, 6).Value = _numbers[idx] * ((i % 500) + 1) * 0.1;
        ws.Cell(row, 7).Value = GetStatus(i);
        ws.Cell(row, 8).Value = $"Cat-{(i % 12) + 1}";
        ws.Cell(row, 9).Value = GetRegion(i);
        ws.Cell(row, 10).Value = $"Note for row {row}";
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

    private static void FormatEvenRow(IXLWorksheet ws, int row, int i)
    {
        FormatNameColumn(ws.Cell(row, 1));
        FormatAmountColumn(ws.Cell(row, 2));
        FormatDateColumn(ws.Cell(row, 3));
        FormatQuantityColumn(ws.Cell(row, 4));
        FormatPriceColumn(ws.Cell(row, 5));
        FormatTotalColumn(ws.Cell(row, 6));
        FormatStatusColumn(ws.Cell(row, 7), i);
        FormatCategoryColumn(ws.Cell(row, 8));
        FormatRegionColumn(ws.Cell(row, 9));
        FormatNotesColumn(ws.Cell(row, 10));
    }

    private static void FormatNameColumn(IXLCell cell)
    {
        cell.Style.Font.Bold = true;
        cell.Style.Fill.BackgroundColor = XLColor.LightBlue;
    }

    private static void FormatAmountColumn(IXLCell cell)
    {
        cell.Style.NumberFormat.Format = "#,##0.00";
        cell.Style.Font.FontColor = XLColor.DarkRed;
        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        cell.Style.Border.OutsideBorderColor = XLColor.Gray;
    }

    private static void FormatDateColumn(IXLCell cell)
    {
        cell.Style.NumberFormat.NumberFormatId = 15;
        cell.Style.Font.Italic = true;
        cell.Style.Fill.BackgroundColor = XLColor.LightGreen;
    }

    private static void FormatQuantityColumn(IXLCell cell)
    {
        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
        cell.Style.Border.OutsideBorderColor = XLColor.DarkGoldenrod;
        cell.Style.Fill.BackgroundColor = XLColor.LightYellow;
    }

    private static void FormatPriceColumn(IXLCell cell)
    {
        cell.Style.NumberFormat.Format = "$ #,##0.00";
        cell.Style.Font.FontName = "Consolas";
        cell.Style.Font.FontSize = 10;
        cell.Style.Fill.BackgroundColor = XLColor.LightCoral;
    }

    private static void FormatTotalColumn(IXLCell cell)
    {
        cell.Style.NumberFormat.Format = "_($* #,##0.00_)";
        cell.Style.Font.Bold = true;
        cell.Style.Border.BottomBorder = XLBorderStyleValues.Dashed;
        cell.Style.Border.BottomBorderColor = XLColor.Navy;
    }

    private static void FormatStatusColumn(IXLCell cell, int i)
    {
        cell.Style.Font.Strikethrough = i % 3 == 2;
        cell.Style.Font.FontColor = (i % 3) switch
        {
            0 => XLColor.Green,
            1 => XLColor.Orange,
            _ => XLColor.Red
        };
        cell.Style.Fill.BackgroundColor = XLColor.Lavender;
    }

    private static void FormatCategoryColumn(IXLCell cell)
    {
        cell.Style.Font.Underline = XLFontUnderlineValues.Single;
        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        cell.Style.Border.RightBorder = XLBorderStyleValues.Dotted;
        cell.Style.Border.RightBorderColor = XLColor.Purple;
        cell.Style.Fill.BackgroundColor = XLColor.LightGray;
    }

    private static void FormatRegionColumn(IXLCell cell)
    {
        cell.Style.Font.FontName = "Georgia";
        cell.Style.Font.FontSize = 12;
        cell.Style.Font.FontColor = XLColor.Teal;
        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
        cell.Style.Border.OutsideBorderColor = XLColor.Teal;
        cell.Style.Fill.BackgroundColor = XLColor.Wheat;
    }

    private static void FormatNotesColumn(IXLCell cell)
    {
        cell.Style.Alignment.WrapText = true;
        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Double;
        cell.Style.Border.OutsideBorderColor = XLColor.Chocolate;
        cell.Style.Fill.BackgroundColor = XLColor.LightSalmon;
        cell.Style.Font.FontColor = XLColor.DarkSlateGray;
    }
}
