using System;
using System.IO;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// The smallest useful round trip: open a small workbook, change two document properties, save.
/// </summary>
/// <remarks>
/// Mirrors the <c>OpenAmendPropertiesAndSave</c> scenario of the cross-library XLBench suite, on
/// the same fixture. Nothing in the edit touches a worksheet, so the whole cost is the fixed price
/// of loading an 8,000-cell sheet into the model and serialising it back out.
/// </remarks>
[MemoryDiagnoser]
public class PropertiesRoundTripBenchmarks
{
    private byte[] _propertiesBytes = null!;

    [GlobalSetup]
    public void Setup()
    {
        SixLaborsV1FontBootstrap.Register();
        _propertiesBytes = PropertiesFixture.Load();
    }

    [Benchmark]
    public long OpenAmendPropertiesAndSave()
    {
        using var wb = new XLWorkbook(new MemoryStream(_propertiesBytes));

        wb.Properties.Title = PropertiesFixture.Title;
        wb.Properties.Category = PropertiesFixture.Category;

        using var output = new MemoryStream();
        wb.SaveAs(output);
        return output.Length;
    }

    /// <summary>Parse only, to split the round trip between load and save.</summary>
    [Benchmark]
    public int Open()
    {
        using var wb = new XLWorkbook(new MemoryStream(_propertiesBytes));
        return wb.Worksheets.Count;
    }
}

/// <summary>
/// The XLBench <c>PropertiesData</c> workbook: a header row over 1,000 rows of eight numeric
/// columns, from a fixed seed.
/// </summary>
public static class PropertiesFixture
{
    public const string Title = "XLBench properties round trip";
    public const string Category = "Benchmark";

    private const int RowCount = 1_000;
    private const int ColCount = 8;

    /// <summary>
    /// The fixture to measure: the file named by <c>XLIBUR_PROPERTIES_FIXTURE</c> when set, otherwise
    /// one built here.
    /// </summary>
    /// <remarks>
    /// XLBench builds its copy with ClosedXML, whose output differs from XLibur's in small ways, and
    /// XLibur cannot reference ClosedXML to build the same bytes. To measure exactly what XLBench
    /// measures, save that workbook to a file and point the variable at it.
    /// </remarks>
    public static byte[] Load() =>
        Environment.GetEnvironmentVariable("XLIBUR_PROPERTIES_FIXTURE") is { Length: > 0 } path
            ? File.ReadAllBytes(path)
            : Build();

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
}
