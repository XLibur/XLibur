using System.Security.Cryptography;
using System.Text;

namespace PerfReport;

/// <summary>Hashes a benchmark class's source file, so an edit to the benchmark re-baselines it.</summary>
public static class SourceHasher
{
    public const string Unknown = "unknown";

    /// <summary>Hash of <c>{dir}/{typeName}.cs</c> with line endings normalised, or <see cref="Unknown"/>.</summary>
    public static string HashFor(string dir, string typeName)
    {
        var path = Path.Combine(dir, typeName + ".cs");
        if (!File.Exists(path))
            return Unknown;

        // A CRLF checkout and an LF checkout are the same source.
        var text = File.ReadAllText(path).Replace("\r\n", "\n");
        return Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(text)))[..16];
    }
}
