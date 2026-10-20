using System.IO;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// The load of a sheet of pictures, against the number of pictures (issue #804).
/// </summary>
/// <remarks>
/// The fixture is <see cref="KeptAnchorScalingProfile"/>'s: pictures spread down a sheet of 12,500
/// sized rows, each anchored over three columns and four rows, every one named. Linear in the
/// pictures, doubling them doubles the time; a load in which each picture checks its id and name
/// against every picture already loaded takes more than twice as long.
/// </remarks>
[MemoryDiagnoser]
public class PictureLoadBenchmarks
{
    private const int Rows = 12_500;

    private byte[] _source = null!;

    [Params(1_000, 2_000, 4_000)]
    public int Pictures { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        SixLaborsV1FontBootstrap.Register();
        _source = KeptAnchorScalingProfile.Fixture(Rows, Pictures);
    }

    [Benchmark]
    public int Load()
    {
        using var workbook = new XLWorkbook(new MemoryStream(_source, writable: false));
        return workbook.Worksheet(1).Pictures.Count;
    }
}
