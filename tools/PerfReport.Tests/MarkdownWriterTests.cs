using PerfReport;

namespace PerfReport.Tests;

public class MarkdownWriterTests
{
    private static ReleaseRecord Record(string version, double meanNs, long alloc, bool manual = false) => new()
    {
        Version = version,
        Manual = manual,
        RecordedAt = $"2026-01-{Math.Abs(version.GetHashCode()) % 28 + 1:00}T00:00:00Z",
        Benchmarks =
        [
            new BenchmarkEntry
            {
                Key = "Bench.Run(N=1)", Method = "Run", Parameters = "N=1",
                MeanNs = meanNs, AllocatedBytes = alloc, SourceHash = "h",
            },
        ],
    };

    private static List<ReleaseRecord> History(params ReleaseRecord[] records) => [.. records];

    [Test]
    public async Task LinksToXlBenchAndItsChangelog()
    {
        var md = MarkdownWriter.RenderHistory(History(Record("1.0.0", 1000, 1000)));

        await Assert.That(md).Contains("https://jafin.github.io/XLBench/");
        await Assert.That(md).Contains("https://github.com/jafin/XLBench/blob/main/CHANGELOG.md");
    }

    [Test]
    public async Task ShowsOnlyTheLastFiveReleasesOfABenchmark()
    {
        var records = Enumerable.Range(1, 7)
            .Select(i => Record($"1.{i}.0", 1000, 1000))
            .Select((r, i) => { r.RecordedAt = $"2026-01-{i + 1:00}T00:00:00Z"; return r; })
            .ToArray();

        var md = MarkdownWriter.RenderHistory(History(records));

        await Assert.That(md).Contains("| 1.7.0 |");
        await Assert.That(md).Contains("| 1.3.0 |");
        await Assert.That(md).DoesNotContain("| 1.2.0 |");
    }

    [Test]
    public async Task ManualAndPreReleaseRunsAreNotInTheTable()
    {
        var md = MarkdownWriter.RenderHistory(History(
            Record("1.0.0", 1000, 1000),
            Record("manual-run", 1000, 1000, manual: true),
            Record("1.1.0-beta.1", 1000, 1000)));

        await Assert.That(md).Contains("| 1.0.0 |");
        await Assert.That(md).DoesNotContain("manual-run");
        await Assert.That(md).DoesNotContain("1.1.0-beta.1");
    }

    [Test]
    public async Task FlagsAnAllocationIncreaseAndLabelsTimeAsIndicative()
    {
        var a = Record("1.0.0", 1000, 1000);
        var b = Record("1.1.0", 1000, 2000);
        b.RecordedAt = "2026-02-01T00:00:00Z";

        var md = MarkdownWriter.RenderHistory(History(a, b));

        await Assert.That(md).Contains("alloc +100.0%");
        await Assert.That(md).Contains("indicative");
    }

    [Test]
    public async Task ShowsRebaselinedForAChangedFingerprint()
    {
        var a = Record("1.0.0", 1000, 1000);
        var b = Record("1.1.0", 1000, 9000);
        b.RecordedAt = "2026-02-01T00:00:00Z";
        b.Benchmarks[0].SourceHash = "changed";

        var md = MarkdownWriter.RenderHistory(History(a, b));

        await Assert.That(md).Contains("re-baselined");
        await Assert.That(md).DoesNotContain("alloc +");
    }

    [Test]
    public async Task SummaryShowsTheCurrentRunWithItsFlags()
    {
        var previous = Record("1.0.0", 1000, 1000);
        var current = Record("1.1.0", 1000, 1500);
        var rows = Comparer.Compare(current, History(previous));

        var md = MarkdownWriter.RenderSummary(current, rows);

        await Assert.That(md).Contains("1.1.0");
        await Assert.That(md).Contains("Bench.Run(N=1)");
        await Assert.That(md).Contains("alloc +50.0%");
    }

    [Test]
    public async Task SummaryOfAFirstRunSaysItSeededHistory()
    {
        var current = Record("1.0.0", 1000, 1000);

        var md = MarkdownWriter.RenderSummary(current, Comparer.Compare(current, History()));

        await Assert.That(md).Contains("seeded");
        await Assert.That(md).DoesNotContain("alloc +");
    }
}
