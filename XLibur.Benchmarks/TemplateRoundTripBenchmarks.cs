using System;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// The fixed cost of touching an existing workbook: open a structurally realistic template,
/// change little or nothing, save it again.
/// </summary>
/// <remarks>
/// This is the statistically trustworthy counterpart to <see cref="TemplateRoundTripProfile"/>.
/// The profile decomposes the round trip into many stages quickly, which is what makes it useful
/// while iterating, but its fastest-of-N timings move by tens of percent between runs of
/// identical code on a loaded machine — the same order as the effects worth chasing. Use the
/// profile to find where the time goes and these benchmarks to claim that it moved.
/// <para>
/// The workload is deliberately small. Every benchmark here is dominated by cost that scales with
/// the workbook's structure rather than its data, which is the cost a request-per-export service
/// pays on every single request no matter how little it writes.
/// </para>
/// </remarks>
[MemoryDiagnoser]
public class TemplateRoundTripBenchmarks
{
    private const int SheetCount = 10;
    private const int DefinedNames = 20;
    private const int Validations = 26;
    private const int LookupValues = 1_000;

    /// <summary>
    /// Rows in the <see cref="LoadRowHeavy"/> fixture. Large enough that per-row costs in the
    /// loader dominate the per-workbook ones, small enough to benchmark in reasonable time —
    /// <c>XLiburReadBenchmarks</c> covers the 250,000-row end.
    /// </summary>
    private const int HeavyRows = 20_000;

    private byte[] _template = null!;
    private byte[] _rowHeavy = null!;
    private string[] _lookupValues = null!;

    [GlobalSetup]
    public void Setup()
    {
        SixLaborsV1FontBootstrap.Register();
        _template = TemplateFixture.Build(SheetCount, DefinedNames, Validations, dataRows: 0);
        _rowHeavy = TemplateFixture.Build(sheetCount: 1, definedNames: 1, validations: 0, dataRows: HeavyRows);
        _lookupValues = Enumerable.Range(1, LookupValues).Select(i => $"Lookup value {i}").ToArray();
    }

    /// <summary>
    /// Open a workbook that already holds many rows and save it again without touching anything.
    /// </summary>
    /// <remarks>
    /// The realistic template-export shape, and the only benchmark here whose *stored* part is
    /// row-heavy — <see cref="OpenAndSaveUnchanged"/> saves a fixture whose sheets are nearly
    /// empty on disk, so it cannot see any save cost that scales with the rows already in the
    /// part. Compare against <see cref="LoadRowHeavy"/> to separate the save half from the parse.
    /// </remarks>
    [Benchmark]
    public void OpenAndSaveRowHeavyUnchanged()
    {
        using var input = new MemoryStream(_rowHeavy, writable: false);
        using var workbook = new XLWorkbook(input);
        using var output = new MemoryStream();
        workbook.SaveAs(output);
    }

    /// <summary>
    /// Parse a sheet whose cost is dominated by rows rather than structure.
    /// </summary>
    /// <remarks>
    /// Guards the cost of getting the SDK reader past <c>&lt;sheetData&gt;</c> in the loader's
    /// first pass. That is per-row work which does not show up in the other benchmarks here, all
    /// of which use structurally rich but nearly empty sheets.
    /// </remarks>
    [Benchmark]
    public int LoadRowHeavy()
    {
        using var input = new MemoryStream(_rowHeavy, writable: false);
        using var workbook = new XLWorkbook(input);
        return workbook.Worksheets.Count;
    }

    /// <summary>Parse only — the floor cost of reading the template.</summary>
    [Benchmark]
    public int Open()
    {
        using var input = new MemoryStream(_template, writable: false);
        using var workbook = new XLWorkbook(input);
        return workbook.Worksheets.Count;
    }

    /// <summary>
    /// Parse and re-serialise with no modification at all. Everything this costs is overhead:
    /// the caller has changed nothing.
    /// </summary>
    [Benchmark]
    public void OpenAndSaveUnchanged()
    {
        using var input = new MemoryStream(_template, writable: false);
        using var workbook = new XLWorkbook(input);
        using var output = new MemoryStream();
        workbook.SaveAs(output);
    }

    /// <summary>
    /// The realistic unit of work: refresh one lookup column and repoint the defined name over
    /// it. The edit itself is trivial, so the gap to <see cref="OpenAndSaveUnchanged"/> is what
    /// the actual work costs and everything below it is the round-trip floor.
    /// </summary>
    [Benchmark]
    public void RefreshLookupColumn()
    {
        using var input = new MemoryStream(_template, writable: false);
        using var workbook = new XLWorkbook(input);
        var sheet = workbook.Worksheet(TemplateFixture.FirstLookupSheet);

        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? TemplateFixture.HeaderRow;
        if (lastRow > TemplateFixture.HeaderRow)
            sheet.Range(TemplateFixture.HeaderRow + 1, 1, lastRow, 1).Clear(XLClearOptions.Contents);

        for (var i = 0; i < _lookupValues.Length; i++)
            sheet.Cell(TemplateFixture.HeaderRow + 1 + i, 1).Value = _lookupValues[i];

        if (workbook.DefinedNames.TryGetValue(TemplateFixture.LookupRangeName, out var definedName))
        {
            definedName.SetRefersTo(
                sheet.Range(TemplateFixture.HeaderRow + 1, 1, TemplateFixture.HeaderRow + _lookupValues.Length, 1));
        }

        using var output = new MemoryStream();
        workbook.SaveAs(output);
    }
}
