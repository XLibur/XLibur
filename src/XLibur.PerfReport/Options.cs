using System.Globalization;

namespace PerfReport;

/// <summary>The command line of <c>PerfReport</c>.</summary>
internal sealed class Options
{
    public string Bdn { get; private set; } = "";
    public string History { get; private set; } = "";
    public string Version { get; private set; } = "";
    public string Commit { get; private set; } = "";
    public string RunnerImage { get; private set; } = "";
    public int? Cores { get; private set; }
    public long PeakWorkingSetBytes { get; private set; }
    public string? BenchmarksDir { get; private set; }
    public bool Manual { get; private set; }

    /// <summary>Only regenerate <c>docs/performance.md</c> from the stored records; no run is added.</summary>
    public bool RenderOnly { get; private set; }

    public string? Summary { get; private set; }
    public string? IssueBodyPath { get; private set; }
    public string? GithubOutput { get; private set; }
    public string? RecordedAt { get; private set; }

    public static Options Parse(string[] args)
    {
        var o = new Options();
        var cursor = new Cursor(args);
        while (cursor.TryRead(out var name))
            o.Apply(name, cursor);

        o.Validate();
        return o;
    }

    private void Apply(string name, Cursor cursor)
    {
        switch (name)
        {
            case "--bdn": Bdn = cursor.Value(name); break;
            case "--history": History = cursor.Value(name); break;
            case "--version": Version = cursor.Value(name); break;
            case "--commit": Commit = cursor.Value(name); break;
            case "--runner-image": RunnerImage = cursor.Value(name); break;
            case "--cores": Cores = int.Parse(cursor.Value(name), CultureInfo.InvariantCulture); break;
            case "--peak-working-set-bytes": PeakWorkingSetBytes = long.Parse(cursor.Value(name), CultureInfo.InvariantCulture); break;
            case "--benchmarks-dir": BenchmarksDir = cursor.Value(name); break;
            case "--manual": Manual = true; break;
            case "--render-only": RenderOnly = true; break;
            case "--summary": Summary = cursor.Value(name); break;
            case "--issue-body": IssueBodyPath = cursor.Value(name); break;
            case "--github-output": GithubOutput = cursor.Value(name); break;
            case "--recorded-at": RecordedAt = cursor.Value(name); break;
            default: throw new ArgumentException($"Unknown argument '{name}'.");
        }
    }

    private void Validate()
    {
        var required = RenderOnly
            ? new[] { ("--history", History) }
            : [("--bdn", Bdn), ("--history", History), ("--version", Version), ("--commit", Commit)];
        foreach (var (name, value) in required)
        {
            if (value.Length == 0)
                throw new ArgumentException($"{name} is required.");
        }
    }

    private sealed class Cursor(string[] args)
    {
        private int _next;

        public bool TryRead(out string name)
        {
            name = _next < args.Length ? args[_next++] : "";
            return name.Length > 0;
        }

        public string Value(string option) =>
            _next < args.Length ? args[_next++] : throw new ArgumentException($"{option} needs a value.");
    }
}
