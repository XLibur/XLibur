using PerfReport;

namespace PerfReport.Tests;

public class VersionOrderTests
{
    private static ReleaseRecord R(string version, long alloc, string at = "2026-01-01T00:00:00Z") => new()
    {
        Version = version,
        RecordedAt = at,
        Benchmarks = [new BenchmarkEntry { Key = "B.Run", Method = "Run", MeanNs = 1000, AllocatedBytes = alloc, SourceHash = "h" }],
    };

    [Test]
    public async Task SortsBySemverNotByRecordingTime()
    {
        var sorted = VersionOrder.Sort([
            R("1.10.0", 1, "2026-01-01T00:00:00Z"),
            R("1.2.0", 1, "2026-03-01T00:00:00Z"),
            R("1.2.1", 1, "2026-04-01T00:00:00Z"),
            R("1.3.0-beta.1", 1, "2026-02-01T00:00:00Z"),
            R("1.3.0", 1, "2026-02-15T00:00:00Z"),
        ]);

        await Assert.That(sorted.Select(r => r.Version).ToArray())
            .IsEquivalentTo(new[] { "1.2.0", "1.2.1", "1.3.0-beta.1", "1.3.0", "1.10.0" }, TUnit.Assertions.Enums.CollectionOrdering.Matching);
    }

    [Test]
    public async Task UnparseableVersionsSortAfterSemverByRecordingTime()
    {
        var sorted = VersionOrder.Sort([R("main-abc", 1, "2026-02-01T00:00:00Z"), R("1.0.0", 1), R("main-def", 1, "2026-01-01T00:00:00Z")]);

        await Assert.That(sorted.Select(r => r.Version).ToArray())
            .IsEquivalentTo(new[] { "1.0.0", "main-def", "main-abc" }, TUnit.Assertions.Enums.CollectionOrdering.Matching);
    }

    [Test]
    public async Task ReRunOfAnOlderVersionComparesWithTheReleaseBeforeIt()
    {
        var history = new[] { R("1.0.0", 1000), R("1.1.0", 1000), R("1.2.0", 5000) };

        var row = Comparer.Compare(R("1.1.0", 1100), history).Single();

        await Assert.That(row.AllocDeltaPercent!.Value).IsEqualTo(10.0).Within(0.0001);
        await Assert.That(row.AllocFlag).IsTrue();
    }

    [Test]
    public async Task HotfixTagComparesWithTheVersionItPatches()
    {
        // 1.3.0 is already recorded when the 1.2.1 hotfix is cut.
        var history = new[] { R("1.2.0", 1000), R("1.3.0", 9000) };

        var row = Comparer.Compare(R("1.2.1", 1010, "2026-06-01T00:00:00Z"), history).Single();

        await Assert.That(row.AllocDeltaPercent!.Value).IsEqualTo(1.0).Within(0.0001);
        await Assert.That(row.AllocFlag).IsFalse();
    }

    [Test]
    public async Task OldestVersionRecordedLastIsStillNew()
    {
        var row = Comparer.Compare(R("0.9.0", 1), new[] { R("1.0.0", 1000) }).Single();

        await Assert.That(row.IsNew).IsTrue();
    }
}
