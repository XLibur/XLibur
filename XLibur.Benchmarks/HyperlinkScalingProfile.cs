using System;
using System.Diagnostics;
using XLibur.Excel;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Benchmarks;

/// <summary>
/// Whether <c>XLHyperlinks</c> scales linearly in the number of hyperlinks on a sheet.
///
/// Run with: <c>dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile hyperlinks</c>
/// </summary>
/// <remarks>
/// Spec 19 area 5 task 5.6. <c>XLHyperlinks</c> keys a <c>Dictionary&lt;Area, XLHyperlink&gt;</c> and
/// every hyperlink it stores is a single cell, so it had exactly the shape that
/// <c>Area.GetHashCode</c> collapsed: <c>first ^ last</c> is zero when the corners are equal, and all
/// of them landed in one bucket. The dependency tree's version of this made building it quadratic
/// (task 5.4); this is the same defect on a path nothing benchmarks.
/// <para>
/// Three operations, because they cost different things. Adding goes through
/// <c>XLCell.SetHyperlink</c>, which is one lookup plus a remove and an add. Looking up by address is
/// the single dictionary probe. Deleting is the one to watch for a different reason: it resolves the
/// hyperlink back to its area with a LINQ scan of the whole dictionary, which is O(N) whatever the
/// hash does, so it should stay quadratic even after the fix.
/// </para>
/// <para>
/// Doubling the count doubles a linear cost and quadruples a quadratic one. Times are single passes
/// and only the growth ratio is being read from them.
/// </para>
/// </remarks>
public static class HyperlinkScalingProfile
{
    private static readonly int[] Counts = [2_500, 5_000, 10_000, 20_000];

    public static void Run()
    {
        SixLaborsV1FontBootstrap.Register();

        // Warm the paths so the smallest count does not absorb JIT.
        Add(1_000);

        Console.WriteLine();
        Console.WriteLine("Hyperlink operations against the number of hyperlinks on one sheet.");
        Console.WriteLine("Doubling the count: ~2x is linear, ~4x is quadratic.");
        Console.WriteLine();
        Console.WriteLine("  | hyperlinks | add ms | vs prev | lookup ms | vs prev | delete ms | vs prev |");
        Console.WriteLine("  |------------|--------|---------|-----------|---------|-----------|---------|");

        double prevAdd = 0, prevLookup = 0, prevDelete = 0;

        foreach (var count in Counts)
        {
            var (addMs, workbook) = Add(count);
            var lookupMs = Lookup(workbook, count);
            var deleteMs = Delete(workbook, count);
            workbook.Dispose();

            Console.WriteLine(
                $"  | {count,10:N0} | {addMs,6:F1} | {Ratio(addMs, prevAdd),7} | " +
                $"{lookupMs,9:F1} | {Ratio(lookupMs, prevLookup),7} | " +
                $"{deleteMs,9:F1} | {Ratio(deleteMs, prevDelete),7} |");

            prevAdd = addMs;
            prevLookup = lookupMs;
            prevDelete = deleteMs;
        }

        Console.WriteLine();
        Console.WriteLine("  Delete resolves a hyperlink to its area by scanning the dictionary, so it is O(N) per");
        Console.WriteLine("  call regardless of how the keys hash. It is listed to keep the two apart.");
    }

    private static string Ratio(double current, double previous) =>
        previous > 0 ? $"{current / previous,6:F2}x" : "      -";

    private static (double Ms, XLWorkbook Workbook) Add(int count)
    {
        var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Links");

        var sw = Stopwatch.StartNew();
        for (var row = 1; row <= count; row++)
            ws.Cell(row, 1).SetHyperlink(new XLHyperlink($"https://example.com/{row}"));
        sw.Stop();

        return (sw.Elapsed.TotalMilliseconds, workbook);
    }

    private static double Lookup(XLWorkbook workbook, int count)
    {
        var ws = workbook.Worksheet(1);

        var sw = Stopwatch.StartNew();
        var found = 0;
        for (var row = 1; row <= count; row++)
        {
            if (ws.Cell(row, 1).HasHyperlink)
                found++;
        }

        sw.Stop();

        if (found != count)
            throw new InvalidOperationException($"Expected {count} hyperlinks, found {found}.");

        return sw.Elapsed.TotalMilliseconds;
    }

    private static double Delete(XLWorkbook workbook, int count)
    {
        var ws = workbook.Worksheet(1);

        var sw = Stopwatch.StartNew();
        for (var row = 1; row <= count; row++)
            ws.Cell(row, 1).SetHyperlink(null);
        sw.Stop();

        return sw.Elapsed.TotalMilliseconds;
    }
}
