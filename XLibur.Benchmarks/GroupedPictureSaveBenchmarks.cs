using System;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.PinataLabs;

namespace XLibur.Benchmarks;

/// <summary>
/// The save of a loaded sheet whose pictures are all in one group shape, against the number of
/// pictures (issue #804).
/// </summary>
/// <remarks>
/// A save writes each grouped picture back in place, so it finds the picture's <c>xdr:pic</c> in
/// the drawing by its id. Linear in the pictures, doubling them doubles the time; the save that
/// read the whole drawing for each grouped picture took four times as long. A save changes the
/// workbook it saves, so each iteration loads a fresh copy outside the timer.
/// </remarks>
[MemoryDiagnoser]
[InvocationCount(1)]
public class GroupedPictureSaveBenchmarks
{
    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEklEQVR4nGP4z8CAFTEMLQkAI7c/wTpVXRAAAAAASUVORK5CYII=");

    private byte[] _source = null!;
    private XLWorkbook _workbook = null!;

    // The in-process toolchain gives up on a benchmark that takes too long, which the quadratic save
    // of 4,000 grouped pictures did.
    [Params(500, 1_000, 2_000)]
    public int Pictures { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        PinataLabsFontBootstrap.Register();
        _source = Fixture(Pictures);
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

    /// <summary>A sheet of <paramref name="pictures"/> free-floating pictures, saved, then grouped and saved again.</summary>
    private static byte[] Fixture(int pictures)
    {
        using var free = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            for (var i = 0; i < pictures; i++)
                ws.AddPicture(new MemoryStream(Png), $"Picture{i}").MoveTo(i % 50 * 20, i / 50 * 20);

            wb.SaveAs(free);
        }

        free.Position = 0;
        using var grouped = new MemoryStream();
        using (var wb = new XLWorkbook(free))
        {
            var ws = wb.Worksheet(1);
            ws.Pictures.Group(ws.Pictures.ToArray());
            wb.SaveAs(grouped);
        }

        return grouped.ToArray();
    }
}
