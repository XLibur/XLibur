using System.IO;
using System.IO.Compression;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.PinataLabs;

namespace XLibur.Benchmarks;

/// <summary>
/// The round trips of <see cref="PropertiesRoundTripBenchmarks"/> at each
/// <see cref="SaveOptions.CompressionLevel"/>, to price the level on a loaded workbook (#688).
/// </summary>
/// <remarks>
/// Changing the document properties leaves the sheet's part as the file had it, so only the small
/// parts the save always writes are compressed. Editing a cell makes the save write the sheet too,
/// which is where the level costs or saves time.
/// </remarks>
[MemoryDiagnoser]
public class RoundTripCompressionBenchmarks
{
    private byte[] _propertiesBytes = null!;

    [Params(CompressionLevel.Optimal, CompressionLevel.Fastest, CompressionLevel.NoCompression)]
    public CompressionLevel Level { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        PinataLabsFontBootstrap.Register();
        _propertiesBytes = PropertiesFixture.Load();
    }

    [Benchmark]
    public long OpenAmendPropertiesAndSave()
    {
        using var wb = new XLWorkbook(new MemoryStream(_propertiesBytes));

        wb.Properties.Title = PropertiesFixture.Title;
        wb.Properties.Category = PropertiesFixture.Category;

        using var output = new MemoryStream();
        wb.SaveAs(output, new SaveOptions { CompressionLevel = Level });
        return output.Length;
    }

    [Benchmark]
    public long OpenEditCellAndSave()
    {
        using var wb = new XLWorkbook(new MemoryStream(_propertiesBytes));

        wb.Worksheet(1).Cell(2, 1).Value = 1.5;

        using var output = new MemoryStream();
        wb.SaveAs(output, new SaveOptions { CompressionLevel = Level });
        return output.Length;
    }
}
