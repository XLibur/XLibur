using System;
using System.IO;
using System.IO.Compression;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.PinataLabs;

namespace XLibur.Benchmarks;

/// <summary>
/// Open, amend two document properties, save, on a workbook with one sheet whose part is larger
/// than the 16 MB a load buffers (#723).
/// </summary>
/// <remarks>
/// Nothing in the edit touches the sheet, so a save that keeps the unchanged sheet's part (#702,
/// tier 1) does not serialise or deflate its cells.
/// </remarks>
[MemoryDiagnoser]
public class LargeSheetRoundTripBenchmarks
{
    private byte[] _bytes = null!;

    [GlobalSetup]
    public void Setup()
    {
        PinataLabsFontBootstrap.Register();
        _bytes = LargeSheetFixture.Build();
        Console.WriteLine($"// Sheet part: {LargeSheetFixture.SheetPartLength(_bytes):N0} bytes inflated, package {_bytes.Length:N0} bytes");
    }

    [Benchmark]
    public long OpenAmendPropertiesAndSave()
    {
        using var wb = new XLWorkbook(new MemoryStream(_bytes));

        wb.Properties.Title = PropertiesFixture.Title;
        wb.Properties.Category = PropertiesFixture.Category;

        using var output = new MemoryStream();
        wb.SaveAs(output);
        return output.Length;
    }

    /// <summary>
    /// The same round trip with one cell of the large sheet changed, so the save writes its cells.
    /// </summary>
    [Benchmark]
    public long OpenEditCellAndSave()
    {
        using var wb = new XLWorkbook(new MemoryStream(_bytes));

        wb.Worksheet(1).Cell(2, 2).Value = 1;

        using var output = new MemoryStream();
        wb.SaveAs(output);
        return output.Length;
    }

    /// <summary>Parse only, to split the round trip between load and save.</summary>
    [Benchmark]
    public int Open()
    {
        using var wb = new XLWorkbook(new MemoryStream(_bytes));
        return wb.Worksheets.Count;
    }
}

/// <summary>
/// A header row over 70,000 rows of eight numeric columns, from a fixed seed: a sheet part of
/// about 22 MB inflated.
/// </summary>
public static class LargeSheetFixture
{
    private const int RowCount = 70_000;
    private const int ColCount = 8;

    public static byte[] Build()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Numbers");

        for (var c = 1; c <= ColCount; c++)
            ws.Cell(1, c).Value = $"Column_{c}";

#pragma warning disable S2245 // Deterministic seed is intentional for reproducible benchmarks
        var random = new Random(42);
#pragma warning restore S2245

        for (var i = 0; i < RowCount; i++)
        {
            var row = 2 + i;
            ws.Cell(row, 1).Value = Math.Round(random.NextDouble() * 10000, 2);
            ws.Cell(row, 2).Value = random.Next(1, 5000);
            ws.Cell(row, 3).Value = Math.Round(random.NextDouble(), 4);
            ws.Cell(row, 4).Value = Math.Round(random.NextDouble() * 1000, 2);
            ws.Cell(row, 5).Value = random.Next(0, 100);
            ws.Cell(row, 6).Value = Math.Round(random.NextDouble() * -5000, 3);
            ws.Cell(row, 7).Value = random.Next(-1000, 1000);
            ws.Cell(row, 8).Value = Math.Round(random.NextDouble() * 100, 4);
        }

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    /// <summary>The inflated length of the first worksheet part.</summary>
    public static long SheetPartLength(byte[] package)
    {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        return archive.GetEntry("xl/worksheets/sheet1.xml")!.Length;
    }
}
