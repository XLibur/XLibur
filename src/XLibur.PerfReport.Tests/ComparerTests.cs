using PerfReport;

namespace PerfReport.Tests;

public class ComparerTests
{
    private const string Key = "Bench.Run";

    private static BenchmarkEntry Entry(double meanNs, long alloc, string hash = "h1", string key = Key) =>
        new() { Key = key, Method = "Run", Parameters = "", MeanNs = meanNs, AllocatedBytes = alloc, SourceHash = hash };

    private static ReleaseRecord Record(string version, double meanNs, long alloc, string hash = "h1",
        bool manual = false, string at = "2026-01-01T00:00:00Z") => new()
    {
        Version = version,
        Manual = manual,
        RecordedAt = at,
        Benchmarks = [Entry(meanNs, alloc, hash)],
    };

    private static ComparisonRow Single(ReleaseRecord current, params ReleaseRecord[] history) =>
        Comparer.Compare(current, history).Single();

    [Test]
    public async Task FirstRunSeedsHistoryWithNoFlags()
    {
        var row = Single(Record("1.0.0", 1000, 1000));

        await Assert.That(row.IsNew).IsTrue();
        await Assert.That(row.AllocFlag).IsFalse();
        await Assert.That(row.TimeFlag).IsFalse();
    }

    [Test]
    public async Task AllocationAboveFivePercentOverPreviousReleaseIsFlagged()
    {
        var row = Single(Record("1.1.0", 1000, 1060), Record("1.0.0", 1000, 1000));

        await Assert.That(row.AllocFlag).IsTrue();
        await Assert.That(row.AllocDeltaPercent!.Value).IsEqualTo(6.0).Within(0.0001);
    }

    [Test]
    public async Task AllocationAtFivePercentIsNotFlagged()
    {
        var row = Single(Record("1.1.0", 1000, 1050), Record("1.0.0", 1000, 1000));

        await Assert.That(row.AllocFlag).IsFalse();
    }

    [Test]
    public async Task AllocationIsComparedToPreviousReleaseOnly()
    {
        var row = Single(Record("1.2.0", 1000, 1040),
            Record("1.0.0", 1000, 500), Record("1.1.0", 1000, 1000));

        await Assert.That(row.AllocFlag).IsFalse();
    }

    [Test]
    public async Task TimeAboveThirtyPercentOverMedianOfLastFiveIsFlagged()
    {
        // Median of the last five is 1000; the sixth release is older and must be ignored.
        var history = new[]
        {
            Record("0.9.0", 9000, 1), Record("1.0.0", 1000, 1), Record("1.1.0", 1000, 1),
            Record("1.2.0", 1000, 1), Record("1.3.0", 900, 1), Record("1.4.0", 1100, 1),
        };

        var flagged = Single(Record("1.5.0", 1310, 1), history);
        var notFlagged = Single(Record("1.5.0", 1290, 1), history);

        await Assert.That(flagged.TimeFlag).IsTrue();
        await Assert.That(notFlagged.TimeFlag).IsFalse();
    }

    [Test]
    public async Task ChangedFingerprintIsRebaselinedAndNotFlagged()
    {
        var row = Single(Record("1.1.0", 5000, 5000, hash: "h2"), Record("1.0.0", 1000, 1000, hash: "h1"));

        await Assert.That(row.Rebaselined).IsTrue();
        await Assert.That(row.AllocFlag).IsFalse();
        await Assert.That(row.TimeFlag).IsFalse();
    }

    [Test]
    public async Task TimeMedianIgnoresReleasesBeforeAFingerprintChange()
    {
        var row = Single(Record("1.2.0", 1000, 1, hash: "h2"),
            Record("1.0.0", 100, 1, hash: "h1"), Record("1.1.0", 1000, 1, hash: "h2"));

        await Assert.That(row.TimeFlag).IsFalse();
        await Assert.That(row.Rebaselined).IsFalse();
    }

    [Test]
    public async Task ManualAndPreReleaseRecordsAreExcludedFromBaselines()
    {
        var history = new[]
        {
            Record("1.0.0", 1000, 1000),
            Record("manual", 100, 10, manual: true),
            Record("1.1.0-beta.1", 100, 10),
        };

        var row = Single(Record("1.1.0", 1000, 1040), history);

        await Assert.That(row.AllocFlag).IsFalse();
        await Assert.That(row.TimeFlag).IsFalse();
    }

    [Test]
    public async Task ManualRunIsNeverFlagged()
    {
        var row = Single(Record("manual", 9000, 9000, manual: true), Record("1.0.0", 1000, 1000));

        await Assert.That(row.AllocFlag).IsFalse();
        await Assert.That(row.TimeFlag).IsFalse();
    }

    [Test]
    public async Task BenchmarkMissingFromHistoryIsNewNotFlagged()
    {
        var current = new ReleaseRecord
        {
            Version = "1.1.0",
            RecordedAt = "2026-02-01T00:00:00Z",
            Benchmarks = [Entry(1000, 1000, key: "Bench.Added")],
        };

        var row = Single(current, Record("1.0.0", 1000, 1000));

        await Assert.That(row.IsNew).IsTrue();
        await Assert.That(row.AllocFlag).IsFalse();
    }
}
