using System.IO;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// The cost of renaming and deleting a sheet of the template round-trip workbook (spec 55 task 8).
/// </summary>
/// <remarks>
/// Both edits change the workbook, so a benchmark cannot repeat one on the same copy. Each iteration
/// opens <see cref="Books"/> fresh copies of the template outside the timer, and the benchmark edits
/// one sheet in each. A single edit takes microseconds, which is why a batch is timed and reported
/// per edit through <c>OperationsPerInvoke</c> rather than timed one invocation at a time.
/// <para>
/// The input is the one <see cref="TemplateRoundTripBenchmarks"/> measures: ten sheets, twenty
/// workbook-scoped defined names over the lookup sheets, and twenty-six list validations. It holds no
/// cell formulas. The edited sheet is <see cref="TemplateFixture.FirstLookupSheet"/>, the sheet the
/// most defined names point at.
/// </para>
/// </remarks>
[MemoryDiagnoser]
[InvocationCount(1)]
public class SheetLifecycleBenchmarks
{
    // The shape of TemplateRoundTripBenchmarks' _template, whose constants are private to it.
    private const int SheetCount = 10;
    private const int DefinedNames = 20;
    private const int Validations = 26;

    private const int Books = 32;

    private readonly XLWorkbook[] _books = new XLWorkbook[Books];
    private byte[] _template = null!;

    [GlobalSetup]
    public void Setup()
    {
        SixLaborsV1FontBootstrap.Register();
        _template = TemplateFixture.Build(SheetCount, DefinedNames, Validations, dataRows: 0);
    }

    [IterationSetup]
    public void OpenBooks()
    {
        for (var i = 0; i < Books; i++)
            _books[i] = new XLWorkbook(new MemoryStream(_template, writable: false));
    }

    [IterationCleanup]
    public void DisposeBooks()
    {
        foreach (var book in _books)
            book.Dispose();
    }

    /// <summary><c>IXLWorksheet.Name =</c>, once per copy.</summary>
    [Benchmark(OperationsPerInvoke = Books)]
    public void Rename()
    {
        foreach (var book in _books)
            book.Worksheet(TemplateFixture.FirstLookupSheet).Name = "Renamed";
    }

    /// <summary><c>IXLWorksheet.Delete()</c>, once per copy.</summary>
    [Benchmark(OperationsPerInvoke = Books)]
    public void Delete()
    {
        foreach (var book in _books)
            book.Worksheet(TemplateFixture.FirstLookupSheet).Delete();
    }
}
