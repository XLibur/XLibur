using System;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// The cost of styling cells one at a time, split by call shape, against the same writes unstyled.
///
/// Run with:
/// dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- --filter '*CellStyling*'
/// </summary>
/// <remarks>
/// Spec 18 task 2. <c>profile template</c> puts per-cell number formatting at ~2.3x the allocation of
/// formatting the column once, and the repo's own <c>profile alloc</c> shows the formatted 50K-row
/// create allocating 183 MB against 25 MB unformatted. Neither says *what* is being allocated, and
/// the answer decides the fix.
/// <para>
/// The variants exist to separate three candidates that a single number cannot:
/// </para>
/// <list type="bullet">
/// <item><see cref="ValueOnly"/> is the floor — cell lookup plus the value-slice write.</item>
/// <item><see cref="StyleFacadePerCell"/> is the natural shape, and pays for the
/// <c>XLStyle</c>/<c>XLNumberFormat</c> façade chain on each cell.</item>
/// <item><see cref="StyleAssignedPerCell"/> writes the same style through a pre-built
/// <see cref="IXLStyle"/>, so it pays the style-slice write without building any façade. The gap to
/// <see cref="StyleFacadePerCell"/> is what the façades cost.</item>
/// <item><see cref="TwoPropertiesPerCell"/> sets two properties on one cell. If the façade cost is
/// per-cell it barely moves from <see cref="StyleFacadePerCell"/>; if it is per-property it roughly
/// doubles.</item>
/// <item><see cref="StyleOnColumn"/> is the fast path callers are told to use, for scale.</item>
/// </list>
/// </remarks>
[MemoryDiagnoser]
public class CellStylingBenchmarks
{
    private const int Rows = 20_000;
    private const string DateFormat = "mmm/yyyy";

    private XLWorkbook _workbook = null!;
    private IXLWorksheet _worksheet = null!;
    private IXLStyle _preBuiltStyle = null!;
    private DateTime[] _dates = null!;

    [GlobalSetup]
    public void GlobalSetup()
    {
        SixLaborsV1FontBootstrap.Register();

        _dates = new DateTime[Rows];
        var start = new DateTime(2026, 7, 1, 0, 0, 0, DateTimeKind.Unspecified);
        for (var i = 0; i < Rows; i++)
            _dates[i] = start.AddMonths(i % 24);
    }

    [IterationSetup]
    public void IterationSetup()
    {
        _workbook = new XLWorkbook();
        _worksheet = _workbook.AddWorksheet("Data");

        // Built once, off the measured path: a style carrying the number format under test.
        var donor = _workbook.AddWorksheet("Donor").Cell(1, 1);
        donor.Style.NumberFormat.Format = DateFormat;
        _preBuiltStyle = donor.Style;
    }

    [IterationCleanup]
    public void IterationCleanup() => _workbook.Dispose();

    [Benchmark(Baseline = true)]
    public void ValueOnly()
    {
        for (var r = 0; r < Rows; r++)
            _worksheet.Cell(r + 1, 1).Value = _dates[r];
    }

    [Benchmark]
    public void StyleFacadePerCell()
    {
        for (var r = 0; r < Rows; r++)
        {
            var cell = _worksheet.Cell(r + 1, 1);
            cell.Value = _dates[r];
            cell.Style.NumberFormat.Format = DateFormat;
        }
    }

    [Benchmark]
    public void StyleAssignedPerCell()
    {
        for (var r = 0; r < Rows; r++)
        {
            var cell = _worksheet.Cell(r + 1, 1);
            cell.Value = _dates[r];
            cell.Style = _preBuiltStyle;
        }
    }

    [Benchmark]
    public void TwoPropertiesPerCell()
    {
        for (var r = 0; r < Rows; r++)
        {
            var cell = _worksheet.Cell(r + 1, 1);
            cell.Value = _dates[r];
            cell.Style.NumberFormat.Format = DateFormat;
            cell.Style.Font.Bold = true;
        }
    }

    /// <summary>Builds the <c>XLStyle</c> façade per cell and nothing else.</summary>
    [Benchmark]
    public void FacadeStyleOnly()
    {
        for (var r = 0; r < Rows; r++)
        {
            var cell = _worksheet.Cell(r + 1, 1);
            cell.Value = _dates[r];
            GC.KeepAlive(cell.Style);
        }
    }

    /// <summary>
    /// Builds the whole façade chain per cell, but sets nothing. The gap to
    /// <see cref="FacadeStyleOnly"/> is the <c>XLNumberFormat</c> façade; the gap from here to
    /// <see cref="StyleFacadePerCell"/> is the setter and the style write it triggers.
    /// </summary>
    [Benchmark]
    public void FacadeChainOnly()
    {
        for (var r = 0; r < Rows; r++)
        {
            var cell = _worksheet.Cell(r + 1, 1);
            cell.Value = _dates[r];
            GC.KeepAlive(cell.Style.NumberFormat);
        }
    }

    [Benchmark]
    public void StyleOnColumn()
    {
        _worksheet.Column(1).Style.NumberFormat.Format = DateFormat;
        for (var r = 0; r < Rows; r++)
            _worksheet.Cell(r + 1, 1).Value = _dates[r];
    }
}
