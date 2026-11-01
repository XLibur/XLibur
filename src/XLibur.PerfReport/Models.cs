using System.Text.Json.Serialization;

namespace PerfReport;

/// <summary>One benchmark measured in one run.</summary>
public sealed class BenchmarkEntry
{
    /// <summary>Stable identity: type, method and parameters. Matches an entry across releases.</summary>
    public string Key { get; set; } = "";

    public string Method { get; set; } = "";

    public string Parameters { get; set; } = "";

    public double MeanNs { get; set; }

    public long AllocatedBytes { get; set; }

    /// <summary>Gen0, Gen1 and Gen2 collections per 1,000 operations, as BenchmarkDotNet reports them.</summary>
    public double Gen0 { get; set; }

    public double Gen1 { get; set; }

    public double Gen2 { get; set; }

    /// <summary>
    /// Hash of the benchmark source file. Together with <see cref="Method"/> and <see cref="Parameters"/>
    /// this is the fingerprint: when it changes, the benchmark measures something else and is re-baselined.
    /// </summary>
    public string SourceHash { get; set; } = "";

    [JsonIgnore]
    public string Fingerprint => $"{Method}|{Parameters}|{SourceHash}";
}

/// <summary>One run of the benchmark set, stored as one JSON file.</summary>
public sealed class ReleaseRecord
{
    public const int CurrentSchema = 1;

    public int Schema { get; set; } = CurrentSchema;

    public string Version { get; set; } = "";

    public string Commit { get; set; } = "";

    /// <summary>A run started by hand. Stored, but never a baseline, never flagged, never in the table.</summary>
    public bool Manual { get; set; }

    public string RecordedAt { get; set; } = "";

    public string RunnerImage { get; set; } = "";

    public int Cores { get; set; }

    public long PeakWorkingSetBytes { get; set; }

    public List<BenchmarkEntry> Benchmarks { get; set; } = [];

    /// <summary>A tag such as <c>1.2.0-beta.1</c>. Stored, but not a baseline.</summary>
    [JsonIgnore]
    public bool PreRelease => Version.Contains('-');

    [JsonIgnore]
    public bool IsBaseline => !Manual && !PreRelease;
}

/// <summary>The comparison of one benchmark in one run against earlier releases.</summary>
public sealed class ComparisonRow
{
    public required BenchmarkEntry Entry { get; init; }

    /// <summary>No earlier release measured this benchmark; nothing to compare with.</summary>
    public bool IsNew { get; init; }

    /// <summary>The fingerprint differs from the previous release, so the history restarts here.</summary>
    public bool Rebaselined { get; init; }

    public bool AllocFlag { get; init; }

    public bool TimeFlag { get; init; }

    /// <summary>Allocation change against the previous release, in percent. Infinity when it was zero.</summary>
    public double? AllocDeltaPercent { get; init; }

    /// <summary>Mean time change against the median of the last five releases, in percent.</summary>
    public double? TimeDeltaPercent { get; init; }
}
