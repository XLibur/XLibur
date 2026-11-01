namespace PerfReport;

/// <summary>
/// Compares a run with the releases before it. Allocated bytes are deterministic, so they carry the
/// tight threshold; time on a shared runner is noisy, so it is only indicative.
/// </summary>
public static class Comparer
{
    public const double AllocThresholdPercent = 5.0;

    public const double TimeThresholdPercent = 30.0;

    public const int MedianWindow = 5;

    /// <param name="current">The run to judge.</param>
    /// <param name="history">
    /// Other runs, in any order. Manual and pre-release runs are ignored. "Earlier" means a lower version,
    /// not recorded earlier: a re-run of an old release compares with the release before it, and a hotfix
    /// tag compares with the version it patches.
    /// </param>
    public static IReadOnlyList<ComparisonRow> Compare(ReleaseRecord current, IReadOnlyList<ReleaseRecord> history)
    {
        var baselines = VersionOrder.Sort(history.Where(r => r.IsBaseline
            && (!VersionOrder.IsParseable(current.Version) || VersionOrder.IsLower(r, current))));
        return current.Benchmarks.Select(entry => CompareEntry(current, entry, baselines)).ToList();
    }

    private static ComparisonRow CompareEntry(ReleaseRecord current, BenchmarkEntry entry, List<ReleaseRecord> baselines)
    {
        // Oldest first; the newest comes last.
        var earlier = baselines
            .Select(r => r.Benchmarks.FirstOrDefault(b => b.Key == entry.Key))
            .Where(b => b is not null)
            .Select(b => b!)
            .ToList();

        if (earlier.Count == 0)
            return new ComparisonRow { Entry = entry, IsNew = true };

        var previous = earlier[^1];
        if (previous.Fingerprint != entry.Fingerprint)
            return new ComparisonRow { Entry = entry, Rebaselined = true };

        var allocDelta = PercentChange(previous.AllocatedBytes, entry.AllocatedBytes);

        // The window stops at the first release with another fingerprint: older numbers measured something else.
        var window = new List<double>();
        for (var i = earlier.Count - 1; i >= 0 && window.Count < MedianWindow; i--)
        {
            if (earlier[i].Fingerprint != entry.Fingerprint)
                break;
            window.Add(earlier[i].MeanNs);
        }

        var timeDelta = PercentChange(Median(window), entry.MeanNs);
        var flagging = !current.Manual;

        return new ComparisonRow
        {
            Entry = entry,
            AllocDeltaPercent = allocDelta,
            TimeDeltaPercent = timeDelta,
            AllocFlag = flagging && allocDelta > AllocThresholdPercent,
            TimeFlag = flagging && timeDelta > TimeThresholdPercent,
        };
    }

    private static double PercentChange(double before, double after)
    {
        if (before == 0)
            return after == 0 ? 0 : double.PositiveInfinity;
        return (after - before) / before * 100.0;
    }

    private static double Median(List<double> values)
    {
        var sorted = values.OrderBy(v => v).ToList();
        var mid = sorted.Count / 2;
        return sorted.Count % 2 == 1 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2.0;
    }
}
