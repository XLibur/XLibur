using System.Security.Cryptography;
using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace PerfReport;

/// <summary>
/// Hashes what a benchmark class measures, so an edit to the benchmark re-baselines it. The hash covers the
/// class's own file and the helper files it names (fixtures and data builders). A file is hashed as its C#
/// tokens: each token's kind and text, so string and number literals keep their exact value, while comments,
/// whitespace and line endings between tokens do not count.
/// </summary>
/// <remarks>
/// This is approximate. A helper counts when the class file mentions a type the helper declares, and then
/// the helper's own helpers count too. A file that holds a benchmark class is never a helper of another
/// class. Code in the library under test is not hashed, as a change there is what the benchmark exists to
/// measure. A helper reached only through a member the class never names, or through reflection, is missed.
/// Code inside a disabled <c>#if</c> branch is trivia and is not hashed. Each file is read once per instance.
/// </remarks>
public sealed class SourceHasher
{
    public const string Unknown = "unknown";

    private sealed record SourceFile(string Name, string Tokens, HashSet<string> Declared, HashSet<string> Mentioned, bool HasBenchmarks);

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
            foreach (var helper in files.Values.Where(c => IsHelperOf(c, file, included)))
            {
                included[helper.Name] = helper;
                queue.Enqueue(helper);
            }
        }

        var sb = new StringBuilder();
        foreach (var file in included.Values)
            sb.Append(file.Name).Append('\n').Append(file.Tokens).Append('\n');
        return Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(sb.ToString())))[..16];
    }

    private static bool IsHelperOf(SourceFile candidate, SourceFile user, SortedDictionary<string, SourceFile> included) =>
        !candidate.HasBenchmarks && !included.ContainsKey(candidate.Name) && candidate.Declared.Overlaps(user.Mentioned);

    private static Dictionary<string, SourceFile> Load(string dir)
    {
        // Keyed by file name without extension, which is the class name for a benchmark file.
        var result = new Dictionary<string, SourceFile>(StringComparer.Ordinal);
        if (!Directory.Exists(dir))
            return result;

        foreach (var path in Directory.GetFiles(dir, "*.cs", SearchOption.TopDirectoryOnly))
        {
            var name = Path.GetFileNameWithoutExtension(path);
            result[name] = Parse(name, File.ReadAllText(path));
        }

        return result;
    }

    private static SourceFile Parse(string name, string text)
    {
        var root = CSharpSyntaxTree.ParseText(text).GetRoot();
        var tokens = new StringBuilder();
        var mentioned = new HashSet<string>(StringComparer.Ordinal);
        foreach (var token in root.DescendantTokens())
        {
            // The kind separates `"a"` from an identifier `a`; the text keeps literals exact.
            tokens.Append((int)token.Kind()).Append(':').Append(token.Text).Append('\u001f');
            if (token.IsKind(SyntaxKind.IdentifierToken))
                mentioned.Add(token.Text);
        }

        var declared = root.DescendantNodes().OfType<BaseTypeDeclarationSyntax>()
            .Select(d => d.Identifier.Text).ToHashSet(StringComparer.Ordinal);
        var hasBenchmarks = root.DescendantNodes().OfType<AttributeSyntax>().Any(IsBenchmarkAttribute);
        return new SourceFile(name, tokens.ToString(), declared, mentioned, hasBenchmarks);
    }

    private static bool IsBenchmarkAttribute(AttributeSyntax attribute)
    {
        var name = attribute.Name.ToString();
        return name is "Benchmark" or "BenchmarkAttribute" || name.EndsWith(".Benchmark", StringComparison.Ordinal);
    }
}
