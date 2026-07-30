using System;
using System.Diagnostics;
using System.Linq;
using XLibur.Excel;

namespace XLibur.Report.Benchmarks;

/// <summary>
/// Splits expansion into the two things it does per item — copy the template block, then evaluate the
/// copy — and reports how each behaves as the sheet grows underneath it.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="ScalingProbe"/> showed generation to be super-linear in the row count: the cost per row
/// roughly doubles between 10,000 rows and 50,000. This probe exists to say <em>which half</em>, and it
/// deliberately uses no report code at all — just a worksheet, a bulk row insert, and the same
/// <c>CopyTo</c> and <c>CellsUsed</c> calls the expander makes. If the curve reproduces here, it is a
/// core-library scaling property and not something the report engine can fix by rearranging itself.
/// </para>
/// <para>
/// Read the buckets, not the total: a per-operation cost that climbs from bucket to bucket is the
/// quadratic term, and its slope says how much of the run it accounts for.
/// </para>
/// </remarks>
public static class ExpansionPhaseProbe
{
    public static void Run(string[] args)
    {
        var rows = args.Skip(1).Select(a => int.TryParse(a, out var v) ? v : 0).FirstOrDefault(v => v > 0);
        rows = rows > 0 ? rows : 50_000;
        var buckets = 10;

        Console.WriteLine($"Expansion phases over {rows:N0} rows of 10 columns, in {buckets} buckets.");
        Console.WriteLine("Per-operation cost climbing bucket to bucket is the quadratic term.");
        Console.WriteLine();

        Warm();

        Measure("CopyTo - copying the template block once per item", rows, buckets, Work.CopyTo);
        Measure("CellsUsed - what evaluation enumerates, per block", rows, buckets, Work.CellsUsed);
        Measure("Range - creating a range object at a new address, and nothing else", rows, buckets, Work.NewRange);
        Measure("Range - creating one at the SAME address over and over", rows, buckets, Work.SameRange);
        Measure("Clear - what CopyTo does to its target before writing", rows, buckets, Work.Clear);
        Measure("CopyFrom - one cell, so without the range-level work", rows, buckets, Work.CopyCell);
        Measure("Clear(Contents) - the same, minus everything Clear does only for All", rows, buckets, Work.ClearContents);
        Measure("Clear(DataValidation) - isolating the create-then-delete of a validation", rows, buckets, Work.ClearDv);
    }

    /// <summary>What one bucket's inner loop does.</summary>
    /// <remarks>
    /// The last two are the discriminating pair. Both create a range object per iteration and do nothing
    /// with it; they differ only in whether the address is one the worksheet has seen before. If the
    /// new-address case climbs and the repeated-address case stays flat, the cost is in <em>accumulating
    /// range objects</em> rather than in copying — which is a property of the range repository, and makes
    /// every caller that creates ranges in a loop over a large sheet quadratic, not just CopyTo.
    /// </remarks>
    private enum Work
    {
        CopyTo,
        CellsUsed,
        NewRange,
        SameRange,
        Clear,
        CopyCell,
        ClearContents,
        ClearDv,
    }

    /// <summary>One small run first, so the first bucket is not paying for JIT.</summary>
    private static void Warm()
    {
        foreach (var work in Enum.GetValues<Work>())
        {
            Measure(null, 500, 1, work);
        }
    }

    private static void Measure(string? title, int rows, int buckets, Work work)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Report");

        // The template block: one row of ten cells, as the expander's is.
        for (var column = 1; column <= 10; column++)
        {
            sheet.Cell(1, column).Value = "value " + column;
        }

        var template = sheet.Range(1, 1, 1, 10);

        // One bulk insert, as the expander does — so what follows measures the per-item work alone.
        sheet.Row(1).InsertRowsBelow(rows - 1);

        if (title is not null)
        {
            Console.WriteLine(title);
            Console.WriteLine(new string('-', title.Length));
            Console.WriteLine("     rows so far      bucket        per operation      vs first bucket");
        }

        var perBucket = Math.Max(1, rows / buckets);
        double? first = null;

        for (var bucket = 0; bucket < buckets; bucket++)
        {
            var from = (bucket * perBucket) + 1;
            var to = Math.Min(rows, from + perBucket - 1);

            var stopwatch = Stopwatch.StartNew();

            for (var row = from; row <= to; row++)
            {
                switch (work)
                {
                    case Work.CopyTo:
                        template.CopyTo(sheet.Cell(row, 1));
                        break;

                    case Work.CellsUsed:
                        // Materialised, as the evaluator does: it writes to the cells it enumerates.
                        _ = sheet.Range(row, 1, row, 10).CellsUsed(XLCellsUsedOptions.Contents).ToList();
                        break;

                    case Work.NewRange:
                        _ = sheet.Range(row, 1, row, 10);
                        break;

                    case Work.SameRange:
                        _ = sheet.Range(1, 1, 1, 10);
                        break;

                    case Work.Clear:
                        sheet.Range(row, 1, row, 10).Clear();
                        break;

                    case Work.CopyCell:
                        sheet.Cell(row, 1).CopyFrom(sheet.Cell(1, 1));
                        break;

                    case Work.ClearContents:
                        sheet.Range(row, 1, row, 10).Clear(XLClearOptions.Contents);
                        break;

                    case Work.ClearDv:
                        sheet.Range(row, 1, row, 10).Clear(XLClearOptions.DataValidation);
                        break;
                }
            }

            stopwatch.Stop();

            if (title is null)
            {
                continue;
            }

            var per = stopwatch.Elapsed.TotalMilliseconds / (to - from + 1) * 1000;
            first ??= per;

            Console.WriteLine(
                $"  {to,12:N0}   {stopwatch.Elapsed.TotalSeconds,8:N2} s   {per,12:N1} us"
                + $"   {per / first.Value,8:N2}x");
        }

        if (title is not null)
        {
            Console.WriteLine();
        }
    }
}
