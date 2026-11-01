using PerfReport;

namespace PerfReport.Tests;

public class CommandLineTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "perfreport-cli2-" + Guid.NewGuid().ToString("N"));

    public CommandLineTests() => Directory.CreateDirectory(_root);

    public void Dispose()
    {
        Directory.Delete(_root, recursive: true);
        GC.SuppressFinalize(this);
    }

    private static int Run(out string error, params string[] args)
    {
        var stderr = new StringWriter();
        var code = Program.Run(args, new StringWriter(), stderr);
        error = stderr.ToString();
        return code;
    }

    private string WriteReport(string name, string mean, long allocated)
    {
        var path = Path.Combine(_root, name);
        File.WriteAllText(path, $$"""
            { "Benchmarks": [ { "Type": "Bench", "Method": "Run", "Statistics": { "Mean": {{mean}} },
              "Memory": { "TotalOperations": 1, "BytesAllocatedPerOperation": {{allocated}} } } ] }
            """);
        return path;
    }

    [Test]
    public async Task UnknownArgumentIsAUsageError()
    {
        var code = Run(out var error, "--nope");

        await Assert.That(code).IsEqualTo(1);
        await Assert.That(error).Contains("Unknown argument");
    }

    [Test]
    public async Task OptionWithoutValueIsAUsageError()
    {
        var code = Run(out var error, "--bdn");

        await Assert.That(code).IsEqualTo(1);
        await Assert.That(error).Contains("needs a value");
    }

    [Test]
    public async Task NonNumericCoreCountIsAUsageError()
    {
        await Assert.That(Run(out _, "--cores", "many")).IsEqualTo(1);
    }

    [Test]
    public async Task MainReturnsTheUsageErrorForNoArguments()
    {
        await Assert.That(Program.Main([])).IsEqualTo(1);
    }

    [Test]
    public async Task ReportFileCanBeNamedDirectly()
    {
        var report = WriteReport("direct-report.json", "100", 10);
        var history = Path.Combine(_root, "h");

        var code = Run(out _, "--bdn", report, "--history", history, "--version", "1.0.0", "--commit", "c");

        await Assert.That(code).IsEqualTo(0);
        await Assert.That(File.Exists(Path.Combine(history, "records", "1.0.0.json"))).IsTrue();
    }

    [Test]
    public async Task FolderWithoutAReportIsAnInputError()
    {
        var empty = Path.Combine(_root, "empty");
        Directory.CreateDirectory(empty);

        var code = Run(out var error, "--bdn", empty, "--history", Path.Combine(_root, "h"), "--version", "1.0.0", "--commit", "c");

        await Assert.That(code).IsEqualTo(1);
        await Assert.That(error).Contains("No BenchmarkDotNet report");
    }

    [Test]
    public async Task ReportWithNoBenchmarksIsAnInputError()
    {
        var path = Path.Combine(_root, "none-report.json");
        File.WriteAllText(path, """{ "Benchmarks": [] }""");

        var code = Run(out var error, "--bdn", path, "--history", Path.Combine(_root, "h"), "--version", "1.0.0", "--commit", "c");

        await Assert.That(code).IsEqualTo(1);
        await Assert.That(error).Contains("no benchmarks");
    }

    [Test]
    public async Task RenderOnlyRegeneratesThePageFromTheStoredRecords()
    {
        var history = Path.Combine(_root, "h");
        Run(out _, "--bdn", WriteReport("a-report.json", "100", 10), "--history", history, "--version", "1.0.0", "--commit", "c");
        var page = Path.Combine(history, "docs", "performance.md");
        File.Delete(page);

        var code = Run(out _, "--render-only", "--history", history);

        await Assert.That(code).IsEqualTo(0);
        await Assert.That(File.ReadAllText(page)).Contains("Bench.Run");
    }

    [Test]
    public async Task RenderOnlyStillNeedsAHistoryFolder()
    {
        await Assert.That(Run(out _, "--render-only")).IsEqualTo(1);
    }

    [Test]
    public async Task AllocationRiseFromZeroIsFlaggedAndNamedInTheIssueBody()
    {
        var history = Path.Combine(_root, "h");
        var issue = Path.Combine(_root, "issue.md");
        Run(out _, "--bdn", WriteReport("a-report.json", "100", 0), "--history", history, "--version", "1.0.0", "--commit", "c");

        Run(out _, "--bdn", WriteReport("b-report.json", "100", 5), "--history", history, "--version", "1.1.0", "--commit", "c",
            "--issue-body", issue);

        await Assert.That(File.ReadAllText(issue)).Contains("from 0 B");
    }

    [Test]
    public async Task TimeRiseIsReportedInTheSummaryAsIndicative()
    {
        var history = Path.Combine(_root, "h");
        var summary = Path.Combine(_root, "summary.md");
        Run(out _, "--bdn", WriteReport("a-report.json", "1000", 10), "--history", history, "--version", "1.0.0", "--commit", "c");

        Run(out _, "--bdn", WriteReport("b-report.json", "5000", 10), "--history", history, "--version", "1.1.0", "--commit", "c",
            "--summary", summary, "--cores", "4", "--runner-image", "img");

        var text = File.ReadAllText(summary);
        await Assert.That(text).Contains("time +400.0% (indicative)");
        await Assert.That(text).Contains("4 cores");
    }
}
