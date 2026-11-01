using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace PerfReport;

/// <summary>
/// Hashes what a benchmark class measures, so an edit to the benchmark re-baselines it. The hash covers the
/// class's own file and the helper files it names (fixtures and data builders), with comments and
/// whitespace removed so that a comment edit or a reformat does not re-baseline.
/// </summary>
/// <remarks>
/// This is approximate. A helper counts when the class file mentions a type the helper declares, and then
/// the helper's own helpers count too. A file that holds a benchmark class is never a helper of another
/// class. Code in the library under test is not hashed, as a change there is what the benchmark exists to
/// measure. A helper reached only through a member the class never names, or through reflection, is missed.
/// Each file is read once per instance.
/// </remarks>
public sealed partial class SourceHasher
{
    public const string Unknown = "unknown";

    [GeneratedRegex(@"//[^\r\n]*|/\*.*?\*/", RegexOptions.Singleline | RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 2000)]
    private static partial Regex Comments();

    [GeneratedRegex(@"\s+", RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 2000)]
    private static partial Regex Whitespace();

    [GeneratedRegex(@"\b(?:class|struct|record|enum|interface)\s+([A-Za-z_]\w*)", RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 2000)]
    private static partial Regex Declaration();

    [GeneratedRegex(@"[A-Za-z_]\w*", RegexOptions.CultureInvariant, matchTimeoutMilliseconds: 2000)]
    private static partial Regex Identifier();

    private sealed record SourceFile(string Name, string Text, HashSet<string> Declared, HashSet<string> Mentioned, bool HasBenchmarks);

    private readonly Lazy<Dictionary<string, SourceFile>> _files;
    private readonly Dictionary<string, string> _hashes = new(StringComparer.Ordinal);

    public SourceHasher(string dir)
    {
        _files = new Lazy<Dictionary<string, SourceFile>>(() => Load(dir));
    }

    /// <summary>The hash for a benchmark class, or <see cref="Unknown"/> when its file is not in the folder.</summary>
    public string HashFor(string typeName)
    {
        if (_hashes.TryGetValue(typeName, out var cached))
            return cached;

        var hash = Compute(typeName);
        _hashes[typeName] = hash;
        return hash;
    }

    private string Compute(string typeName)
    {
        var files = _files.Value;
        if (!files.TryGetValue(typeName, out var own))
            return Unknown;

        var included = new SortedDictionary<string, SourceFile>(StringComparer.Ordinal) { [own.Name] = own };
        var queue = new Queue<SourceFile>();
        queue.Enqueue(own);
        while (queue.Count > 0)
        {
            var file = queue.Dequeue();
            foreach (var candidate in files.Values)
            {
                if (candidate.HasBenchmarks || included.ContainsKey(candidate.Name))
                    continue;
                if (candidate.Declared.Overlaps(file.Mentioned))
                {
                    included[candidate.Name] = candidate;
                    queue.Enqueue(candidate);
                }
            }
        }

        var sb = new StringBuilder();
        foreach (var file in included.Values)
            sb.Append(file.Name).Append('\n').Append(file.Text).Append('\n');
        return Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(sb.ToString())))[..16];
    }

    private static Dictionary<string, SourceFile> Load(string dir)
    {
        // Keyed by file name without extension, which is the class name for a benchmark file.
        var result = new Dictionary<string, SourceFile>(StringComparer.Ordinal);
        if (!Directory.Exists(dir))
            return result;

        foreach (var path in Directory.GetFiles(dir, "*.cs", SearchOption.TopDirectoryOnly))
        {
            var raw = File.ReadAllText(path);
            var stripped = Comments().Replace(raw, " ");
            var text = Whitespace().Replace(stripped, " ").Trim();
            var name = Path.GetFileNameWithoutExtension(path);
            result[name] = new SourceFile(
                name,
                text,
                Declaration().Matches(stripped).Select(m => m.Groups[1].Value).ToHashSet(StringComparer.Ordinal),
                Identifier().Matches(stripped).Select(m => m.Value).ToHashSet(StringComparer.Ordinal),
                stripped.Contains("[Benchmark", StringComparison.Ordinal));
        }

        return result;
    }
}
