using PerfReport;

namespace PerfReport.Tests;

public class CliTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "perfreport-" + Guid.NewGuid().ToString("N"));

    public CliTests() => Directory.CreateDirectory(_root);

    public void Dispose() => Directory.Delete(_root, recursive: true);

    private string Bdn(long allocated, string name = "bdn")
    {
        var dir = Path.Combine(_root, name);
        Directory.CreateDirectory(dir);
        File.WriteAllText(Path.Combine(dir, "XLibur.Benchmarks.Bench-report-full-compressed.json"), $$"""
            { "HostEnvironmentInfo": { "LogicalCoreCount": 4 },
              "Benchmarks": [ { "Type": "Bench", "Method": "Run", "Parameters": "",
                "Statistics": { "Mean": 1000 },
                "Memory": { "Gen0Collections": 1, "Gen1Collections": 0, "Gen2Collections": 0,
                            "TotalOperations": 1, "BytesAllocatedPerOperation": {{allocated}} } } ] }
            """);
        return dir;
    }

    private string Source()
    {
        var dir = Path.Combine(_root, "src");
        Directory.CreateDirectory(dir);
        File.WriteAllText(Path.Combine(dir, "Bench.cs"), "class Bench {}\r\n");
        return dir;
    }

    private (int Code, string Out, string Err) Run(string version, long allocated, params string[] extra)
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        var args = new List<string>
        {
            "--bdn", Bdn(allocated, "bdn-" + version), "--history", Path.Combine(_root, "history"),
            "--benchmarks-dir", Source(), "--version", version, "--commit", "abc123",
            "--runner-image", "ubuntu-24.04", "--peak-working-set-bytes", "1000",
            "--summary", Path.Combine(_root, "summary.md"),
            "--issue-body", Path.Combine(_root, "issue.md"),
            "--github-output", Path.Combine(_root, "output.txt"),
            "--recorded-at", $"2026-01-{version.Replace(".", "").PadLeft(2, '0')[..2]}T00:00:00Z",
        };
        args.AddRange(extra);
        return (Program.Run([.. args], stdout, stderr), stdout.ToString(), stderr.ToString());
    }

    private string HistoryPath(params string[] parts) => Path.Combine([_root, "history", .. parts]);

    [Test]
    public async Task FirstRunAppendsASchemaOneRecordAndWritesTheMarkdown()
    {
        var result = Run("1.0.0", 1000);

        await Assert.That(result.Code).IsEqualTo(0);
        var record = File.ReadAllText(HistoryPath("records", "1.0.0.json"));
        await Assert.That(record).Contains("\"schema\": 1");
        await Assert.That(record).Contains("\"commit\": \"abc123\"");
        await Assert.That(record).Contains("\"cores\": 4");
        await Assert.That(record).Contains("\"peakWorkingSetBytes\": 1000");
        await Assert.That(File.ReadAllText(HistoryPath("docs", "performance.md"))).Contains("Bench.Run");
        await Assert.That(File.ReadAllText(Path.Combine(_root, "output.txt"))).Contains("alloc_flags=0");
    }

    [Test]
    public async Task AllocationIncreaseWritesAnIssueBodyAndCountsTheFlag()
    {
        Run("1.0.0", 1000);

        var result = Run("1.1.0", 1200);

        await Assert.That(result.Code).IsEqualTo(0);
        await Assert.That(File.ReadAllText(Path.Combine(_root, "output.txt"))).Contains("alloc_flags=1");
        await Assert.That(File.ReadAllText(Path.Combine(_root, "issue.md"))).Contains("Bench.Run");
        await Assert.That(File.ReadAllText(Path.Combine(_root, "summary.md"))).Contains("alloc +20.0%");
    }

    [Test]
    public async Task ManualRunIsStoredWithTheManualFlagAndNeverFlagged()
    {
        Run("1.0.0", 1000);

        var result = Run("1.1.0", 5000, "--manual");

        await Assert.That(result.Code).IsEqualTo(0);
        var manual = Directory.GetFiles(HistoryPath("records"), "manual-*.json").Single();
        await Assert.That(File.ReadAllText(manual)).Contains("\"manual\": true");
        await Assert.That(File.ReadAllText(Path.Combine(_root, "output.txt"))).DoesNotContain("alloc_flags=1");
    }

    [Test]
    public async Task UnknownSchemaInHistoryIsRefused()
    {
        Directory.CreateDirectory(HistoryPath("records"));
        File.WriteAllText(HistoryPath("records", "9.0.0.json"), """{ "schema": 2, "version": "9.0.0" }""");

        var result = Run("1.0.0", 1000);

        await Assert.That(result.Code).IsEqualTo(2);
        await Assert.That(result.Err).Contains("schema");
        await Assert.That(File.Exists(HistoryPath("records", "1.0.0.json"))).IsFalse();
    }

    [Test]
    public async Task SourceHashIgnoresLineEndings()
    {
        var lf = Path.Combine(_root, "lf");
        var crlf = Path.Combine(_root, "crlf");
        Directory.CreateDirectory(lf);
        Directory.CreateDirectory(crlf);
        File.WriteAllText(Path.Combine(lf, "Bench.cs"), "a\nb\n");
        File.WriteAllText(Path.Combine(crlf, "Bench.cs"), "a\r\nb\r\n");

        await Assert.That(SourceHasher.HashFor(lf, "Bench")).IsEqualTo(SourceHasher.HashFor(crlf, "Bench"));
        await Assert.That(SourceHasher.HashFor(lf, "Missing")).IsEqualTo("unknown");
    }

    [Test]
    public async Task MissingRequiredArgumentIsAUsageError()
    {
        var result = Program.Run([], new StringWriter(), new StringWriter());

        await Assert.That(result).IsEqualTo(1);
    }
}
