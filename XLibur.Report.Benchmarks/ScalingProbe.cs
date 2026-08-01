using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using XLibur.Excel;
using XLibur.Report;

namespace XLibur.Report.Benchmarks;

/// <summary>
/// Times generation at rising row counts and reports the cost <em>per row</em>, so that whether
/// generation scales linearly is read off a column rather than inferred.
/// </summary>
/// <remarks>
/// <para>
/// This is the instrument spec 12's acceptance criterion 8 actually calls for. BenchmarkDotNet answers
/// "how long does one generation take, precisely" — which needs many iterations of a workload that
/// already takes seconds, and a full run of
/// <see cref="ReportGenerateBenchmarks"/> costs the best part of an hour. The question here is a
/// different one: does the time per row stay flat as the row count grows by two orders of magnitude?
/// One timed run per size answers that, and the answer is not sensitive to the noise that makes a
/// single BenchmarkDotNet iteration untrustworthy.
/// </para>
/// <para>
/// Both instruments stay. Use this one to know the shape of the curve; use the benchmarks to compare
/// two implementations at one size.
/// </para>
/// </remarks>
public static class ScalingProbe
{
    private const string SheetName = "Report";

    private static readonly int[] RowCounts = { 1_000, 5_000, 10_000, 25_000, 50_000, 100_000 };

    public static void Run(string[] args)
    {
        var counts = ParseCounts(args) ?? RowCounts;

        Console.WriteLine("Report generation, 10 columns. Per-row cost is what matters: flat means linear.");
        Console.WriteLine();

        foreach (var (name, options) in Shapes())
        {
            Console.WriteLine(name);
            Console.WriteLine(new string('-', name.Length));
            Console.WriteLine("      rows        total        per row      relative to 1K");

            double? baseline = null;

            foreach (var count in counts)
            {
                var rows = ReportData.Rows(count);

                // Warmed once at the smallest size so the first timing is not paying for JIT and for
                // Scriban's first parse of each expression.
                if (baseline is null)
                {
                    Generate(rows.Take(Math.Min(200, count)).ToList(), options);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                var stopwatch = Stopwatch.StartNew();
                var lastRow = Generate(rows, options);
                stopwatch.Stop();

                var perRow = stopwatch.Elapsed.TotalMilliseconds / count;
                baseline ??= perRow;

                Console.WriteLine(
                    $"  {count,8:N0}   {stopwatch.Elapsed.TotalSeconds,8:N2} s   {perRow * 1000,8:N1} us"
                    + $"   {perRow / baseline.Value,8:N2}x      (last row {lastRow:N0})");
            }

            Console.WriteLine();
        }

        Console.WriteLine("A per-row cost that rises with the row count is the super-linear behaviour");
        Console.WriteLine("criterion 8 exists to catch. Flat, or falling as fixed costs amortise, is linear.");
    }

    /// <summary>The template shapes worth timing, cheapest first.</summary>
    private static IEnumerable<(string Name, Action<IXLWorksheet>? Options)> Shapes()
    {
        yield return ("Plain - ten expressions per row, no tags", null);

        yield return ("Totalled - three SUBTOTAL columns", sheet =>
        {
            sheet.Cell("E3").Value = "<<Sum>>";
            sheet.Cell("F3").Value = "<<Sum>>";
            sheet.Cell("J3").Value = "<<Sum>>";
        }
        );

        yield return ("Grouped - one group level, five groups, two totals", sheet =>
        {
            sheet.Cell("A3").Value = "<<Group>>";
            sheet.Cell("E3").Value = "<<Sum>>";
            sheet.Cell("J3").Value = "<<Sum>>";
        }
        );

        yield return ("GroupedAndSorted - grouping plus a sort inside each group", sheet =>
        {
            sheet.Cell("A3").Value = "<<Group>>";
            sheet.Cell("C3").Value = "<<Sort>>";
            sheet.Cell("E3").Value = "<<Sum>>";
            sheet.Cell("J3").Value = "<<Sum>>";
        }
        );
    }

    private static int Generate(List<ReportRow> rows, Action<IXLWorksheet>? options)
    {
        using var workbook = Template(options);

        using (var template = new XLTemplate(workbook))
        {
            template.AddVariable("Rows", rows);
            template.Generate();
        }

        return workbook.Worksheet(SheetName).LastRowUsed()!.RowNumber();
    }

    private static XLWorkbook Template(Action<IXLWorksheet>? options)
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet(SheetName);

        var headings = new[]
        {
            "Region", "Category", "Product", "Reference", "Quantity",
            "Unit price", "Discount", "Sold on", "Export", "Total",
        };

        for (var i = 0; i < headings.Length; i++)
        {
            sheet.Cell(1, i + 1).Value = headings[i];
        }

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

    private static int[]? ParseCounts(string[] args)
    {
        var counts = args
            .Skip(1)
            .Select(argument => int.TryParse(argument, out var value) ? value : 0)
            .Where(value => value > 0)
            .ToArray();

        return counts.Length > 0 ? counts : null;
    }
}
