using System;
using System.Diagnostics;
using System.IO;
using XLibur.Excel;

namespace XLibur.Benchmarks;

/// <summary>
/// Loops <see cref="PropertiesRoundTripBenchmarks.OpenAmendPropertiesAndSave"/> so an external
/// profiler can be attached to it, and prints a median time and exact allocation per round trip.
/// </summary>
/// <remarks>
/// Run with: dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile properties [seconds]
/// <para>Use BenchmarkDotNet for any claim that time moved; the bytes here are exact.</para>
/// </remarks>
public static class PropertiesRoundTripProfile
{
    public static void Run(string[] args)
    {
        var seconds = args.Length > 2 && int.TryParse(args[2], out var s) ? s : 5;
        var bytes = PropertiesFixture.Load();

        for (var i = 0; i < 50; i++)
            RoundTrip(bytes);

        var before = GC.GetTotalAllocatedBytes(precise: true);
        const int allocPasses = 20;
        for (var i = 0; i < allocPasses; i++)
            RoundTrip(bytes);
        var perPass = (GC.GetTotalAllocatedBytes(precise: true) - before) / allocPasses;

        var pauseBefore = GC.GetTotalPauseDuration();
        var gcBefore = GC.CollectionCount(0);
        var elapsed = Stopwatch.StartNew();
        var iterations = 0;
        while (elapsed.Elapsed.TotalSeconds < seconds)
        {
            RoundTrip(bytes);
            iterations++;
        }

        Console.WriteLine($"{iterations:N0} iterations in {elapsed.Elapsed.TotalSeconds:F1}s, "
            + $"{elapsed.Elapsed.TotalMilliseconds / iterations:F2} ms mean, {perPass / 1024.0:F1} KB per round trip, "
            + $"GC pause {(GC.GetTotalPauseDuration() - pauseBefore).TotalMilliseconds / iterations:F3} ms/iter, "
            + $"{GC.CollectionCount(0) - gcBefore} gen0");
    }

    private static long RoundTrip(byte[] bytes)
    {
        using var wb = new XLWorkbook(new MemoryStream(bytes));
        wb.Properties.Title = PropertiesFixture.Title;
        wb.Properties.Category = PropertiesFixture.Category;
        using var output = new MemoryStream();
        wb.SaveAs(output);
        return output.Length;
    }
}
