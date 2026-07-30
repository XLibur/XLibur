using System.Collections.Generic;
using BenchmarkDotNet.Attributes;
using XLibur.Excel;
using XLibur.Report;

namespace XLibur.Report.Benchmarks;

/// <summary>
/// Generating a report of 10 columns and a varying number of rows, against writing the same grid by
/// hand.
/// </summary>
/// <remarks>
/// <para>
/// The question spec 12 asks is whether generation scales <em>linearly</em> in the row count. It is
/// worth asking because the engine this one is a port of renders onto a hidden buffer sheet and splices
/// the result back, and because the obvious way to write range expansion — insert one row, copy one
/// row, repeat — is quadratic in the number of rows through the range repository and the formula
/// shifter. This engine inserts every row it needs in one call and copies the template block once per
/// item; the row counts below are chosen to show whether that holds up rather than to produce a single
/// impressive number.
/// </para>
/// <para>
/// <see cref="HandWritten"/> is the baseline: the same cells written by an ordinary loop, with no
/// template, no expressions and no tags. The ratio to it is the price of authoring a report instead of
/// programming one, which is the number worth knowing before choosing between them.
/// </para>
/// <para>
/// The template is rebuilt per iteration because generation consumes it. That is not measured — it is a
/// dozen cells — but it is why this benchmark uses an iteration setup rather than a global one.
/// </para>
/// </remarks>
[MemoryDiagnoser]
public class ReportGenerateBenchmarks
{
    private const string SheetName = "Report";

    private List<ReportRow> _rows = new();
    private XLWorkbook? _template;

    /// <summary>
    /// Row counts an order of magnitude apart, so super-linear growth is visible as a ratio rather
    /// than having to be inferred from one point.
    /// </summary>
    [Params(10_000, 50_000, 100_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void GlobalSetup() => _rows = ReportData.Rows(RowCount);

    [IterationCleanup]
    public void IterationCleanup()
    {
        _template?.Dispose();
        _template = null;
    }

    /// <summary>The floor: the same grid, written directly, with nothing to interpret.</summary>
    [Benchmark(Baseline = true)]
    public int HandWritten()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet(SheetName);

        Headings(sheet);

        var row = 2;
        foreach (var item in _rows)
        {
            sheet.Cell(row, 1).Value = item.Region;
            sheet.Cell(row, 2).Value = item.Category;
            sheet.Cell(row, 3).Value = item.Product;
            sheet.Cell(row, 4).Value = item.Reference;
            sheet.Cell(row, 5).Value = item.Quantity;
            sheet.Cell(row, 6).Value = item.UnitPrice;
            sheet.Cell(row, 7).Value = item.Discount;
            sheet.Cell(row, 8).Value = item.SoldOn;
            sheet.Cell(row, 9).Value = item.IsExport;
            sheet.Cell(row, 10).Value = item.Total;
            row++;
        }

        return row;
    }

    /// <summary>Plain expansion: ten expressions per row, no tags.</summary>
    [Benchmark]
    public int Plain()
    {
        _template = Template(options: null);

        return Generate();
    }

    /// <summary>
    /// The same, with a total under each column — the shape nearly every real report has, and the one
    /// that makes the options row survive.
    /// </summary>
    [Benchmark]
    public int Totalled()
    {
        _template = Template(options: sheet =>
        {
            sheet.Cell("E3").Value = "<<Sum>>";
            sheet.Cell("F3").Value = "<<Sum>>";
            sheet.Cell("J3").Value = "<<Sum>>";
        });

        return Generate();
    }

    /// <summary>
    /// One group level with subtotals, which is the workload the spec names. Five regions, so the group
    /// count is fixed and what varies is the number of rows in each.
    /// </summary>
    [Benchmark]
    public int Grouped()
    {
        _template = Template(options: sheet =>
        {
            sheet.Cell("A3").Value = "<<Group>>";
            sheet.Cell("E3").Value = "<<Sum>>";
            sheet.Cell("J3").Value = "<<Sum>>";
        });

        return Generate();
    }

    /// <summary>
    /// Grouped and sorted: the ordering the engine does for grouping, plus a sort inside each group.
    /// Two orderings over 100,000 rows is the most work any single template asks of the engine before
    /// a row is written.
    /// </summary>
    [Benchmark]
    public int GroupedAndSorted()
    {
        _template = Template(options: sheet =>
        {
            sheet.Cell("A3").Value = "<<Group>>";
            sheet.Cell("C3").Value = "<<Sort>>";
            sheet.Cell("E3").Value = "<<Sum>>";
            sheet.Cell("J3").Value = "<<Sum>>";
        });

        return Generate();
    }

    private int Generate()
    {
        using var template = new XLTemplate(_template!);
        template.AddVariable("Rows", _rows);
        template.Generate();

        return _template!.Worksheet(SheetName).LastRowUsed()!.RowNumber();
    }

    /// <summary>
    /// A template with headings in row 1, the repeated row in row 2 and the options row in row 3.
    /// </summary>
    private static XLWorkbook Template(System.Action<IXLWorksheet>? options)
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet(SheetName);

        Headings(sheet);

        sheet.Cell("A2").Value = "{{ item.Region }}";
        sheet.Cell("B2").Value = "{{ item.Category }}";
        sheet.Cell("C2").Value = "{{ item.Product }}";
        sheet.Cell("D2").Value = "{{ item.Reference }}";
        sheet.Cell("E2").Value = "{{ item.Quantity }}";
        sheet.Cell("F2").Value = "{{ item.UnitPrice }}";
        sheet.Cell("G2").Value = "{{ item.Discount }}";
        sheet.Cell("H2").Value = "{{ item.SoldOn }}";
        sheet.Cell("I2").Value = "{{ item.IsExport }}";
        sheet.Cell("J2").Value = "{{ item.Total }}";

        options?.Invoke(sheet);

        workbook.DefinedNames.Add("Rows", sheet.Range("A2:J3"));

        return workbook;
    }

    private static void Headings(IXLWorksheet sheet)
    {
        var headings = new[]
        {
            "Region", "Category", "Product", "Reference", "Quantity",
            "Unit price", "Discount", "Sold on", "Export", "Total",
        };

        for (var i = 0; i < headings.Length; i++)
        {
            sheet.Cell(1, i + 1).Value = headings[i];
        }
    }
}
