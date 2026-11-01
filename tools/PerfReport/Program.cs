using System.Globalization;
using System.Text;

namespace PerfReport;

/// <summary>
/// Appends one benchmark run to the perf-history checkout, compares it with earlier releases and
/// regenerates <c>docs/performance.md</c>. Exit codes: 0 done, 1 bad arguments or input, 2 unknown schema.
/// </summary>
public static class Program
{
    private const string Usage =
        "Usage: PerfReport --bdn <dir|file> --history <dir> --version <v> --commit <sha> [--runner-image <s>]\n" +
        "       [--cores <n>] [--peak-working-set-bytes <n>] [--benchmarks-dir <dir>] [--manual]\n" +
        "       [--summary <file>] [--issue-body <file>] [--github-output <file>] [--recorded-at <iso>]";

    public static int Main(string[] args) => Run(args, Console.Out, Console.Error);

    public static int Run(string[] args, TextWriter output, TextWriter error)
    {
        try
        {
            var options = Options.Parse(args);
            return Execute(options, output);
        }
        catch (UnsupportedSchemaException e)
        {
            error.WriteLine("Refusing to compare: " + e.Message + " (unknown schema)");
            return 2;
        }
        catch (Exception e) when (e is ArgumentException or InvalidDataException or IOException or System.Text.Json.JsonException)
        {
            error.WriteLine(e.Message);
            error.WriteLine(Usage);
            return 1;
        }
    }

    private static int Execute(Options o, TextWriter output)
    {
        var history = HistoryStore.Load(o.History);

        var run = ReadBdn(o);
        var current = new ReleaseRecord
        {
            Version = o.Version.TrimStart('v'),
            Commit = o.Commit,
            Manual = o.Manual,
            RecordedAt = o.RecordedAt ?? DateTime.UtcNow.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture),
            RunnerImage = o.RunnerImage,
            Cores = o.Cores ?? run.LogicalCores,
            PeakWorkingSetBytes = o.PeakWorkingSetBytes,
            Benchmarks = run.Benchmarks,
        };

        // Compare against what was stored before this run, so a re-run of a version does not baseline itself.
        var earlier = history.Where(r => current.Manual || r.Version != current.Version).ToList();
        var rows = Comparer.Compare(current, earlier);

        HistoryStore.Save(o.History, current);
        var all = HistoryStore.Load(o.History);

        var page = Path.Combine(o.History, "docs", "performance.md");
        Directory.CreateDirectory(Path.GetDirectoryName(page)!);
        File.WriteAllText(page, MarkdownWriter.RenderHistory(all));

        var summary = MarkdownWriter.RenderSummary(current, rows);
        if (o.Summary is not null)
            File.AppendAllText(o.Summary, summary);
        output.Write(summary);

        var allocFlags = rows.Where(r => r.AllocFlag).ToList();
        if (o.IssueBody is not null && allocFlags.Count > 0)
            File.WriteAllText(o.IssueBody, IssueBody(current, allocFlags));

        if (o.GithubOutput is not null)
        {
            File.AppendAllText(o.GithubOutput,
                $"alloc_flags={allocFlags.Count}\ntime_flags={rows.Count(r => r.TimeFlag)}\n");
        }

        return 0;
    }

    private static BdnRun ReadBdn(Options o)
    {
        var files = Directory.Exists(o.Bdn)
            ? Directory.GetFiles(o.Bdn, "*-report*.json", SearchOption.AllDirectories)
            : [o.Bdn];
        if (files.Length == 0)
            throw new InvalidDataException($"No BenchmarkDotNet report (*-report*.json) found in {o.Bdn}.");

        string HashOf(string type) => o.BenchmarksDir is null ? SourceHasher.Unknown : SourceHasher.HashFor(o.BenchmarksDir, type);

        var merged = new BdnRun();
        foreach (var file in files.OrderBy(f => f, StringComparer.Ordinal))
        {
            var run = BdnJsonReader.Read(File.ReadAllText(file), HashOf);
            merged.LogicalCores = Math.Max(merged.LogicalCores, run.LogicalCores);
            // The brief and the full exporter both write a file; keep one entry per benchmark.
            foreach (var entry in run.Benchmarks)
            {
                if (merged.Benchmarks.All(b => b.Key != entry.Key))
                    merged.Benchmarks.Add(entry);
            }
        }

        if (merged.Benchmarks.Count == 0)
            throw new InvalidDataException("The BenchmarkDotNet report holds no benchmarks.");
        return merged;
    }

    private static string IssueBody(ReleaseRecord current, List<ComparisonRow> flagged)
    {
        var sb = new StringBuilder();
        sb.Append(CultureInfo.InvariantCulture, $"Allocations rose by more than {Comparer.AllocThresholdPercent:0}% in {current.Version} (commit `{current.Commit}`) against the previous release.\n\n");
        sb.Append("| Benchmark | Allocated | Change |\n|---|---:|---:|\n");
        foreach (var row in flagged)
        {
            var delta = double.IsPositiveInfinity(row.AllocDeltaPercent!.Value)
                ? "from 0 B"
                : string.Create(CultureInfo.InvariantCulture, $"{row.AllocDeltaPercent:+0.0}%");
            sb.Append(CultureInfo.InvariantCulture,
                $"| {row.Entry.Key} | {MarkdownWriter.FormatBytes(row.Entry.AllocatedBytes)} | {delta} |\n");
        }

        sb.Append("\nFull history: `docs/performance.md` on the `perf-history` branch. ");
        sb.Append("Allocated bytes are deterministic, so a flag here is a real change, not runner noise.\n");
        return sb.ToString();
    }

    private sealed class Options
    {
        public string Bdn = "";
        public string History = "";
        public string Version = "";
        public string Commit = "";
        public string RunnerImage = "";
        public int? Cores;
        public long PeakWorkingSetBytes;
        public string? BenchmarksDir;
        public bool Manual;
        public string? Summary;
        public string? IssueBody;
        public string? GithubOutput;
        public string? RecordedAt;

        public static Options Parse(string[] args)
        {
            var o = new Options();
            for (var i = 0; i < args.Length; i++)
            {
                string Next() => i + 1 < args.Length ? args[++i] : throw new ArgumentException($"{args[i]} needs a value.");
                switch (args[i])
                {
                    case "--bdn": o.Bdn = Next(); break;
                    case "--history": o.History = Next(); break;
                    case "--version": o.Version = Next(); break;
                    case "--commit": o.Commit = Next(); break;
                    case "--runner-image": o.RunnerImage = Next(); break;
                    case "--cores": o.Cores = int.Parse(Next(), CultureInfo.InvariantCulture); break;
                    case "--peak-working-set-bytes": o.PeakWorkingSetBytes = long.Parse(Next(), CultureInfo.InvariantCulture); break;
                    case "--benchmarks-dir": o.BenchmarksDir = Next(); break;
                    case "--manual": o.Manual = true; break;
                    case "--summary": o.Summary = Next(); break;
                    case "--issue-body": o.IssueBody = Next(); break;
                    case "--github-output": o.GithubOutput = Next(); break;
                    case "--recorded-at": o.RecordedAt = Next(); break;
                    default: throw new ArgumentException($"Unknown argument '{args[i]}'.");
                }
            }

            foreach (var (name, value) in new[] { ("--bdn", o.Bdn), ("--history", o.History), ("--version", o.Version), ("--commit", o.Commit) })
            {
                if (value.Length == 0)
                    throw new ArgumentException($"{name} is required.");
            }

            return o;
        }
    }
}
