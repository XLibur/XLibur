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

    [Benchmark]
    public void CreateFormattedAndSave()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Formatted");

        // Headers
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

        for (var i = 0; i < RowCount; i++)
        {
            var row = i + 2;
            var idx = i % _strings.Length;

            // Write data to all 10 columns
            ws.Cell(row, 1).Value = _strings[idx];
            ws.Cell(row, 2).Value = _numbers[idx];
            ws.Cell(row, 3).Value = _dates[idx];
            ws.Cell(row, 4).Value = (i % 500) + 1;
            ws.Cell(row, 5).Value = _numbers[idx] * 0.1;
            ws.Cell(row, 6).Value = _numbers[idx] * ((i % 500) + 1) * 0.1;
            ws.Cell(row, 7).Value = i % 3 == 0 ? "Active" : i % 3 == 1 ? "Pending" : "Closed";
            ws.Cell(row, 8).Value = $"Cat-{(i % 12) + 1}";
            ws.Cell(row, 9).Value = i % 5 == 0 ? "North" :
                i % 5 == 1 ? "South" :
                i % 5 == 2 ? "East" :
                i % 5 == 3 ? "West" : "Central";
            ws.Cell(row, 10).Value = $"Note for row {row}";

            // Apply formatting to every second row
            if (i % 2 != 0)
                continue;

            // Col 1: Bold font + light blue background
            var c1 = ws.Cell(row, 1);
            c1.Style.Font.Bold = true;
            c1.Style.Fill.BackgroundColor = XLColor.LightBlue;

            // Col 2: Number format + red font + thin border
            var c2 = ws.Cell(row, 2);
            c2.Style.NumberFormat.Format = "#,##0.00";
            c2.Style.Font.FontColor = XLColor.DarkRed;
            c2.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            c2.Style.Border.OutsideBorderColor = XLColor.Gray;

            // Col 3: Date format + italic + light green fill
            var c3 = ws.Cell(row, 3);
            c3.Style.NumberFormat.NumberFormatId = 15;
            c3.Style.Font.Italic = true;
            c3.Style.Fill.BackgroundColor = XLColor.LightGreen;

            // Col 4: Center aligned + medium border + yellow fill
            var c4 = ws.Cell(row, 4);
            c4.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            c4.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            c4.Style.Border.OutsideBorderColor = XLColor.DarkGoldenrod;
            c4.Style.Fill.BackgroundColor = XLColor.LightYellow;

            // Col 5: Currency format + font change + coral fill
            var c5 = ws.Cell(row, 5);
            c5.Style.NumberFormat.Format = "$ #,##0.00";
            c5.Style.Font.FontName = "Consolas";
            c5.Style.Font.FontSize = 10;
            c5.Style.Fill.BackgroundColor = XLColor.LightCoral;

            // Col 6: Accounting format + bold + dashed border
            var c6 = ws.Cell(row, 6);
            c6.Style.NumberFormat.Format = "_($* #,##0.00_)";
            c6.Style.Font.Bold = true;
            c6.Style.Border.BottomBorder = XLBorderStyleValues.Dashed;
            c6.Style.Border.BottomBorderColor = XLColor.Navy;

            // Col 7: Strikethrough + font color based on value + lavender fill
            var c7 = ws.Cell(row, 7);
            c7.Style.Font.Strikethrough = i % 3 == 2;
            c7.Style.Font.FontColor = i % 3 == 0 ? XLColor.Green : i % 3 == 1 ? XLColor.Orange : XLColor.Red;
            c7.Style.Fill.BackgroundColor = XLColor.Lavender;

            // Col 8: Underline + right aligned + dotted border + light gray fill
            var c8 = ws.Cell(row, 8);
            c8.Style.Font.Underline = XLFontUnderlineValues.Single;
            c8.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            c8.Style.Border.RightBorder = XLBorderStyleValues.Dotted;
            c8.Style.Border.RightBorderColor = XLColor.Purple;
            c8.Style.Fill.BackgroundColor = XLColor.LightGray;

            // Col 9: Different font + large size + thick border + wheat fill
            var c9 = ws.Cell(row, 9);
            c9.Style.Font.FontName = "Georgia";
            c9.Style.Font.FontSize = 12;
            c9.Style.Font.FontColor = XLColor.Teal;
            c9.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            c9.Style.Border.OutsideBorderColor = XLColor.Teal;
            c9.Style.Fill.BackgroundColor = XLColor.Wheat;

            // Col 10: Wrap text + vertical top + double border + light salmon fill
            var c10 = ws.Cell(row, 10);
            c10.Style.Alignment.WrapText = true;
            c10.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            c10.Style.Border.OutsideBorder = XLBorderStyleValues.Double;
            c10.Style.Border.OutsideBorderColor = XLColor.Chocolate;
            c10.Style.Fill.BackgroundColor = XLColor.LightSalmon;
            c10.Style.Font.FontColor = XLColor.DarkSlateGray;
        }

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
    }
}
