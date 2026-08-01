using System;
using System.Diagnostics;
using System.IO;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// Lightweight allocation accounting for the save path, split into a "create" phase and a
/// "save" phase. BenchmarkDotNet reports a single total for
/// <see cref="XLiburWorkbookBenchmarks.CreateFormattedAndSave"/>, which hides whether an
/// optimisation landed on the cell-population side or the serialisation side.
///
/// Run with: dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile alloc
/// </summary>
public static class SaveAllocationProfile
{
    private const int RowCount = 50_000;

    public static void Run()
    {
        SixLaborsV1FontBootstrap.Register();

        var data = BenchmarkData.Create(RowCount);

        // Warm up JIT / static caches so the measured run is not polluted by one-time costs.
        RunPlain(data, out _, out _);
        RunFormatted(data, out _, out _);

        ForceGC();
        var plainWatch = Stopwatch.StartNew();
        RunPlain(data, out var plainCreate, out var plainSave);
        plainWatch.Stop();

        ForceGC();
        var formattedWatch = Stopwatch.StartNew();
        RunFormatted(data, out var formattedCreate, out var formattedSave);
        formattedWatch.Stop();

        Console.WriteLine();
        Console.WriteLine("Allocated bytes (single iteration, 50,000 rows)");
        Console.WriteLine("| Scenario               | Create      | Save        | Total       | Elapsed  |");
        Console.WriteLine("|------------------------|-------------|-------------|-------------|----------|");
        WriteRow("CreateAndSave", plainCreate, plainSave, plainWatch.ElapsedMilliseconds);
        WriteRow("CreateFormattedAndSave", formattedCreate, formattedSave, formattedWatch.ElapsedMilliseconds);
    }

    private static void WriteRow(string name, long create, long save, long elapsedMs)
    {
        Console.WriteLine(
            $"| {name,-22} | {Mb(create),11} | {Mb(save),11} | {Mb(create + save),11} | {elapsedMs + " ms",8} |");
    }

    private static string Mb(long bytes) => $"{bytes / 1024.0 / 1024.0:F1} MB";

    private static void RunPlain(BenchmarkData data, out long createBytes, out long saveBytes)
    {
        var start = GC.GetTotalAllocatedBytes(precise: true);

        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Data");

        worksheet.Cell(1, 1).Value = "Name";
        worksheet.Cell(1, 2).Value = "Amount";
        worksheet.Cell(1, 3).Value = "Date";

        for (var i = 0; i < RowCount; i++)
        {
            var row = i + 2;
            worksheet.Cell(row, 1).Value = data.Strings[i];
            worksheet.Cell(row, 2).Value = data.Numbers[i];
            worksheet.Cell(row, 3).Value = data.Dates[i];
        }

        var sumRow = RowCount + 2;
        worksheet.Cell(sumRow, 1).Value = "Total";
        worksheet.Cell(sumRow, 2).FormulaA1 = $"SUM(B2:B{RowCount + 1})";

        var afterCreate = GC.GetTotalAllocatedBytes(precise: true);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);

        var afterSave = GC.GetTotalAllocatedBytes(precise: true);

        createBytes = afterCreate - start;
        saveBytes = afterSave - afterCreate;
    }

    private static void RunFormatted(BenchmarkData data, out long createBytes, out long saveBytes)
    {
        var start = GC.GetTotalAllocatedBytes(precise: true);

        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Formatted");

        FormattedSheetBuilder.WriteHeaders(ws);

        for (var i = 0; i < RowCount; i++)
        {
            var row = i + 2;
            FormattedSheetBuilder.WriteRowData(ws, data, row, i, i % RowCount);

            if (i % 2 == 0)
                FormattedSheetBuilder.ApplyRowFormatting(ws, row, i);
        }

        var afterCreate = GC.GetTotalAllocatedBytes(precise: true);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);

        var afterSave = GC.GetTotalAllocatedBytes(precise: true);

        createBytes = afterCreate - start;
        saveBytes = afterSave - afterCreate;
    }

    // ReSharper disable once InconsistentNaming
    private static void ForceGC()
    {
#pragma warning disable S1215 // Intentionally forcing GC to isolate phases
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        GC.WaitForPendingFinalizers();
        GC.Collect(2, GCCollectionMode.Forced, true, true);
#pragma warning restore S1215
    }
}

/// <summary>
/// Deterministic input data shared by the workbook benchmarks and the allocation profile,
/// so both measure the same workload.
/// </summary>
public sealed class BenchmarkData
{
    public required string[] Strings { get; init; }

    public required double[] Numbers { get; init; }

    public required DateTime[] Dates { get; init; }

    public static BenchmarkData Create(int rowCount)
    {
        var strings = new string[rowCount];
        var numbers = new double[rowCount];
        var dates = new DateTime[rowCount];

#pragma warning disable S2245 // Deterministic seed for reproducible benchmarks
        var random = new Random(42);
#pragma warning restore S2245
        var baseDate = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);

        for (var i = 0; i < rowCount; i++)
        {
            strings[i] = $"Item {i} - {random.Next(1000):D4}";
            numbers[i] = Math.Round(random.NextDouble() * 10000, 2);
            dates[i] = baseDate.AddDays(random.Next(0, 1500));
        }

        return new BenchmarkData { Strings = strings, Numbers = numbers, Dates = dates };
    }
}

/// <summary>
/// The formatted-sheet workload, shared by <see cref="XLiburWorkbookBenchmarks"/> and
/// <see cref="SaveAllocationProfile"/> so the two never drift apart.
/// </summary>
public static class FormattedSheetBuilder
{
    private static readonly string[] Headers =
        ["Name", "Amount", "Date", "Quantity", "Price", "Total", "Status", "Category", "Region", "Notes"];

    public static void WriteHeaders(IXLWorksheet ws)
    {
        for (var c = 1; c <= 10; c++)
        {
            var hdr = ws.Cell(1, c);
            hdr.Value = Headers[c - 1];
            hdr.Style.Font.Bold = true;
            hdr.Style.Font.FontColor = XLColor.White;
            hdr.Style.Fill.BackgroundColor = XLColor.DarkBlue;
            hdr.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            hdr.Style.Border.BottomBorder = XLBorderStyleValues.Double;
            hdr.Style.Border.BottomBorderColor = XLColor.Black;
        }
    }

    public static void WriteRowData(IXLWorksheet ws, BenchmarkData data, int row, int i, int idx)
    {
        ws.Cell(row, 1).Value = data.Strings[idx];
        ws.Cell(row, 2).Value = data.Numbers[idx];
        ws.Cell(row, 3).Value = data.Dates[idx];
        ws.Cell(row, 4).Value = (i % 500) + 1;
        ws.Cell(row, 5).Value = data.Numbers[idx] * 0.1;
        ws.Cell(row, 6).Value = data.Numbers[idx] * ((i % 500) + 1) * 0.1;
        var status = (i % 3) switch { 0 => "Active", 1 => "Pending", _ => "Closed" };
        ws.Cell(row, 7).Value = status;
        ws.Cell(row, 8).Value = $"Cat-{(i % 12) + 1}";
        var region = (i % 5) switch { 0 => "North", 1 => "South", 2 => "East", 3 => "West", _ => "Central" };
        ws.Cell(row, 9).Value = region;
        ws.Cell(row, 10).Value = $"Note for row {row}";
    }

    public static void ApplyRowFormatting(IXLWorksheet ws, int row, int i)
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
