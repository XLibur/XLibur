using System.Text.Json;

namespace PerfReport;

/// <summary>What the tool takes from one BenchmarkDotNet JSON report.</summary>
public sealed class BdnRun
{
    public List<BenchmarkEntry> Benchmarks { get; } = [];

    public int LogicalCores { get; set; }
}

/// <summary>
/// Reads the file BenchmarkDotNet writes with <c>--exporters json</c>. The brief and full exporters
/// share the properties used here; everything else is ignored.
/// </summary>
public static class BdnJsonReader
{
    /// <param name="json">Contents of a <c>*-report*.json</c> file.</param>
    /// <param name="sourceHashOf">Hash of the source file for a benchmark class name.</param>
    public static BdnRun Read(string json, Func<string, string> sourceHashOf)
    {
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("The BenchmarkDotNet report is not a JSON object.");

        var run = new BdnRun { LogicalCores = ReadLogicalCores(root) };

        if (!root.TryGetProperty("Benchmarks", out var benchmarks) || benchmarks.ValueKind != JsonValueKind.Array)
            throw new InvalidDataException("The BenchmarkDotNet report has no 'Benchmarks' array.");

        foreach (var b in benchmarks.EnumerateArray())
            run.Benchmarks.Add(ReadBenchmark(b, sourceHashOf));

        return run;
    }

    private static int ReadLogicalCores(JsonElement root) =>
        root.TryGetProperty("HostEnvironmentInfo", out var host)
        && host.TryGetProperty("LogicalCoreCount", out var cores)
        && cores.TryGetInt32(out var count)
            ? count
            : 0;

    private static BenchmarkEntry ReadBenchmark(JsonElement b, Func<string, string> sourceHashOf)
    {
        if (b.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException("An entry of 'Benchmarks' is not an object.");

        var type = RequiredString(b, "Type");
        var method = RequiredString(b, "Method");
        var parameters = b.TryGetProperty("Parameters", out var p) && p.ValueKind == JsonValueKind.String
            ? p.GetString() ?? ""
            : "";
        var key = parameters.Length == 0 ? $"{type}.{method}" : $"{type}.{method}({parameters})";

        var mean = RequiredMean(b, key);
        if (!b.TryGetProperty("Memory", out var memory) || memory.ValueKind != JsonValueKind.Object)
            throw new InvalidDataException($"Benchmark '{key}' has no Memory section. Run with MemoryDiagnoser.");

        var ops = Number(memory, "TotalOperations");
        return new BenchmarkEntry
        {
            Key = key,
            Method = method,
            Parameters = parameters,
            MeanNs = mean,
            AllocatedBytes = (long)Number(memory, "BytesAllocatedPerOperation"),
            Gen0 = PerThousand(Number(memory, "Gen0Collections"), ops),
            Gen1 = PerThousand(Number(memory, "Gen1Collections"), ops),
            Gen2 = PerThousand(Number(memory, "Gen2Collections"), ops),
            SourceHash = sourceHashOf(type),
        };
    }

    private static double RequiredMean(JsonElement benchmark, string key)
    {
        if (benchmark.TryGetProperty("Statistics", out var stats) && stats.ValueKind == JsonValueKind.Object
            && stats.TryGetProperty("Mean", out var mean) && mean.ValueKind == JsonValueKind.Number)
        {
            return mean.GetDouble();
        }

        throw new InvalidDataException($"Benchmark '{key}' has no numeric Statistics.Mean (did it fail?).");
    }

    private static string RequiredString(JsonElement element, string name)
    {
        if (element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String)
            return value.GetString()!;
        throw new InvalidDataException($"A benchmark in the report has no '{name}'.");
    }

    private static double Number(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.Number
            ? value.GetDouble()
            : 0;

    private static double PerThousand(double collections, double operations) =>
        operations > 0 ? collections / operations * 1000.0 : 0;
}
