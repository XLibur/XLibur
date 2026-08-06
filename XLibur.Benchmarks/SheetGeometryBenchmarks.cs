using System;
using System.IO;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// The same number of string cells written and saved in two shapes — one tall narrow column versus
/// a short wide block — crossed with whether the strings repeat.
///
/// Run with:
/// dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- --filter '*SheetGeometry*'
/// </summary>
/// <remarks>
/// Spec 18 task 3. Refreshing a 10,000-value lookup column costs ~5.4 µs per cell where the wide
/// grid path runs ~1.3 µs per cell, and the finding was recorded with two candidate causes and no
/// evidence for either: shared-string pressure, because the lookup values are all distinct; or the
/// sheet's geometry, because a single column pays whatever a row costs once per cell instead of
/// once per twenty.
/// <para>
/// The two are crossed deliberately. Comparing only tall-versus-wide would confound geometry with
/// string reuse, and comparing only unique-versus-repeated would confound it the other way; the
/// 2×2 separates them. Cell count is held constant across all four, so a difference is a
/// difference in cost *per cell*.
/// </para>
/// </remarks>
[MemoryDiagnoser]
public class SheetGeometryBenchmarks
{
    private const int Cells = 20_000;
    private const int WideColumns = 20;

    /// <summary>Distinct values in the repeated variants — enough to be a realistic lookup, few enough to share.</summary>
    private const int RepeatedDistinctValues = 10;

    private string[] _unique = null!;
    private string[] _repeated = null!;

    [GlobalSetup]
    public void GlobalSetup()
    {
        SixLaborsV1FontBootstrap.Register();

        _unique = new string[Cells];
        for (var i = 0; i < Cells; i++)
            _unique[i] = $"Lookup value {i}";

        _repeated = new string[Cells];
        for (var i = 0; i < Cells; i++)
            _repeated[i] = $"Lookup value {i % RepeatedDistinctValues}";
    }

    [Benchmark(Baseline = true)]
    public void TallNarrow_Unique() => WriteAndSave(_unique, columns: 1);

    [Benchmark]
    public void TallNarrow_Repeated() => WriteAndSave(_repeated, columns: 1);

    [Benchmark]
    public void ShortWide_Unique() => WriteAndSave(_unique, WideColumns);

    [Benchmark]
    public void ShortWide_Repeated() => WriteAndSave(_repeated, WideColumns);

    /// <summary>
    /// The same writes with no save, to split the geometry penalty between the in-memory row
    /// storage and the serialiser. A cost that survives here is the slice layout; one that only
    /// appears in the saving variants is the <c>&lt;row&gt;</c> element per row.
    /// </summary>
    [Benchmark]
    public void TallNarrow_Unique_WriteOnly() => Write(_unique, columns: 1);

    [Benchmark]
    public void ShortWide_Unique_WriteOnly() => Write(_unique, WideColumns);

    private static void Write(string[] values, int columns)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Data");

        for (var i = 0; i < values.Length; i++)
            sheet.Cell(i / columns + 1, i % columns + 1).Value = values[i];
    }

    private static void WriteAndSave(string[] values, int columns)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Data");

        for (var i = 0; i < values.Length; i++)
            sheet.Cell(i / columns + 1, i % columns + 1).Value = values[i];

        using var output = new MemoryStream();
        workbook.SaveAs(output);
    }
}
