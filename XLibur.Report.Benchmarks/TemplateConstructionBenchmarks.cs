using System.Collections.Generic;
using System.IO;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Report;

namespace XLibur.Report.Benchmarks;

/// <summary>
/// What it costs to <em>start</em> a report, as opposed to generate one.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="ReportGenerateBenchmarks"/> measures a report big enough that its setup disappears
/// into the noise. This one measures the setup, because the service generating a hundred small
/// reports an hour pays it a hundred times and the row work barely at all.
/// </para>
/// <para>
/// Issue #276 is why it exists. Constructing an <see cref="XLTemplate"/> used to import the whole
/// ~400-function Excel library into the expression engine one function at a time — about 30 KB
/// apiece inside Scriban, so some 12 MB per template, discarded when the template was disposed.
/// That was more than nine tenths of everything a small report allocated, and twenty times the cost
/// of opening the template workbook it accompanied. The library is now imported once per culture
/// and shared, which is what <see cref="ConstructTemplate"/> should show against
/// <see cref="OpenWorkbook"/>: the template constructor ought to be a rounding error next to
/// opening the file, and it was not.
/// </para>
/// </remarks>
[MemoryDiagnoser]
public class TemplateConstructionBenchmarks
{
    private const string SheetName = "Report";

    private string _path = string.Empty;
    private XLWorkbook? _workbook;
    private List<ReportRow> _rows = new();

    [GlobalSetup]
    public void GlobalSetup()
    {
        _rows = ReportData.Rows(25);

        _path = Path.Combine(Path.GetTempPath(), "XLibur.Report.Benchmarks.Construction.xlsx");
        using (var template = Template())
        {
            template.SaveAs(_path);
        }

        // Kept open, so ConstructTemplate measures the constructor and nothing else.
        _workbook = new XLWorkbook(_path);
    }

    [GlobalCleanup]
    public void GlobalCleanup()
    {
        _workbook?.Dispose();
        _workbook = null;

        if (File.Exists(_path))
        {
            File.Delete(_path);
        }
    }

    /// <summary>The baseline the constructor should be measured against: reading the .xlsx.</summary>
    [Benchmark(Baseline = true)]
    public int OpenWorkbook()
    {
        using var workbook = new XLWorkbook(_path);

        return workbook.Worksheets.Count;
    }

    /// <summary>The constructor alone, over a workbook that is already open.</summary>
    [Benchmark]
    public bool ConstructTemplate()
    {
        using var template = new XLTemplate(_workbook!);

        return template.IsGenerated;
    }

    /// <summary>
    /// The whole shape of a request that returns one small report: open, bind, generate, save.
    /// </summary>
    [Benchmark]
    public long SmallReport()
    {
        using var workbook = new XLWorkbook(_path);
        using var template = new XLTemplate(workbook);
        template.AddVariable("Rows", _rows);
        template.Generate();

        using var stream = new MemoryStream();
        template.SaveAs(stream);

        return stream.Length;
    }

    private static XLWorkbook Template()
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet(SheetName);

        sheet.Cell("A1").Value = "Region";
        sheet.Cell("B1").Value = "Product";
        sheet.Cell("C1").Value = "Total";

        sheet.Cell("A2").Value = "{{ item.Region }}";
        sheet.Cell("B2").Value = "{{ item.Product }}";
        sheet.Cell("C2").Value = "{{ item.Total }}";
        sheet.Cell("C3").Value = "<<Sum>>";

        workbook.DefinedNames.Add("Rows", sheet.Range("A2:C3"));

        return workbook;
    }
}
