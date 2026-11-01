using System.IO;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.PinataLabs;

namespace XLibur.Benchmarks;

/// <summary>
/// The save of a loaded sheet of pictures, against the number of pictures (issue #797).
/// </summary>
/// <remarks>
/// The fixture is <see cref="KeptAnchorScalingProfile"/>'s: pictures spread down a sheet of 12,500
/// sized rows, each anchored over three columns and four rows. <c>twoCell</c> regenerates every
/// anchor, so it takes a drawing id per picture; <c>oneCell</c> keeps every anchor and only edits it.
/// Both find each picture's anchor in the loaded drawing. Linear in the pictures, doubling them
/// doubles the time; the save that read the whole drawing for each picture took four times as long.
/// A save changes the workbook it saves, so each iteration loads a fresh copy outside the timer.
/// </remarks>
[MemoryDiagnoser]
[InvocationCount(1)]
public class PictureSaveBenchmarks
{
    private const int Rows = 12_500;

    private byte[] _source = null!;
    private XLWorkbook _workbook = null!;

    [Params(1_000, 2_000, 4_000)]
    public int Pictures { get; set; }

    [Params("twoCell", "oneCell")]
    public string EditAs { get; set; } = null!;

    [GlobalSetup]
    public void Setup()
    {
        PinataLabsFontBootstrap.Register();
        _source = KeptAnchorScalingProfile.WithEditAs(KeptAnchorScalingProfile.Fixture(Rows, Pictures), EditAs);
    }

    [IterationSetup]
    public void Load() => _workbook = new XLWorkbook(new MemoryStream(_source, writable: false));

    [IterationCleanup]
    public void Close() => _workbook.Dispose();

    [Benchmark]
    public long Save()
    {
        using var saved = new MemoryStream();
        _workbook.SaveAs(saved);
        return saved.Length;
    }
}
