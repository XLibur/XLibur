using System;
using System.IO;
using System.Reflection;

namespace XLibur.Tests.Utils;

/// <summary>
/// A committed set of XML fixtures pinning exactly what some part of the writer emits today. A
/// refactor that means to change no output is measured against it: any diff is a finding to
/// investigate, never noise to re-baseline without a written explanation.
/// </summary>
/// <remarks>
/// <para>
/// The fixtures travel as embedded resources rather than as files beside the test that reads them.
/// A release build normalises compile-time paths to <c>/_/</c>, so nothing may resolve a fixture
/// from where its source file was, and nothing copies a <c>Golden/</c> directory next to the test
/// binary.
/// </para>
/// <para>
/// Regenerating is deliberately not automatic. A fixture the corpus is missing has to fail the run,
/// or CI would quietly write and then assert against its own output, which gates nothing.
/// </para>
/// </remarks>
internal sealed class GoldenCorpus
{
    private readonly string _relativeDirectory;
    private readonly string _writeEnvironmentVariable;
    private readonly Assembly _assembly = typeof(GoldenCorpus).Assembly;

    /// <param name="relativeDirectory">
    /// Where the fixtures live under the test project, as path segments — for example
    /// <c>Excel/DrawingML/Golden</c>. Both the embedded resource name and the regeneration target
    /// are derived from it.
    /// </param>
    /// <param name="writeEnvironmentVariable">
    /// The variable a developer sets to <c>1</c> to regenerate this corpus.
    /// </param>
    internal GoldenCorpus(string relativeDirectory, string writeEnvironmentVariable)
    {
        _relativeDirectory = relativeDirectory;
        _writeEnvironmentVariable = writeEnvironmentVariable;
    }

    /// <summary>Whether this run may write fixtures back to the source tree.</summary>
    internal bool CanWrite =>
        Environment.GetEnvironmentVariable(_writeEnvironmentVariable) == "1";

    /// <summary>The instruction to put in a failure message when a fixture is missing.</summary>
    internal string RegenerationHint =>
        $"Regenerate with {_writeEnvironmentVariable}=1, then rebuild so the new file is embedded.";

    /// <summary>
    /// The committed fixture of the given name, or <c>null</c> when the corpus does not hold one.
    /// </summary>
    internal string? Read(string name)
    {
        var resource = ResourceNameFor(name);
        if (resource is null)
            return null;

        using var stream = _assembly.GetManifestResourceStream(resource)!;
        using var reader = new StreamReader(stream);
        return Normalise(reader.ReadToEnd());
    }

    /// <summary>
    /// Writes a fixture into the source tree, for a developer regenerating the corpus. The written
    /// file is only picked up by the next build, because the corpus is embedded.
    /// </summary>
    internal void Write(string name, string xml)
    {
        var directory = Path.Combine(
            SourceDirectory(), Path.Combine(_relativeDirectory.Split('/')));
        Directory.CreateDirectory(directory);
        File.WriteAllText(Path.Combine(directory, name + ".xml"), Normalise(xml));
    }

    /// <summary>
    /// Line endings are not part of what a corpus pins — the XML it holds is a single line — but
    /// <c>.gitattributes</c> checks every text file out as CRLF, so a fixture is compared for what
    /// it says rather than for how the working copy stores it.
    /// </summary>
    internal static string Normalise(string xml) => xml.Replace("\r\n", "\n");

    private string? ResourceNameFor(string name)
    {
        var dotted = _relativeDirectory.Replace('/', '.');
        var expected = $"{_assembly.GetName().Name}.{dotted}.{name}.xml";
        if (Array.IndexOf(_assembly.GetManifestResourceNames(), expected) >= 0)
            return expected;

        // Fall back to a suffix match so that moving the corpus, or a project whose root namespace
        // differs from its assembly name, degrades into a slower lookup rather than a missing
        // fixture that reads as a regression.
        var suffix = $".{dotted}.{name}.xml";
        return Array.Find(_assembly.GetManifestResourceNames(),
            resource => resource.EndsWith(suffix, StringComparison.Ordinal));
    }

    /// <summary>
    /// The <c>XLibur.Tests</c> project directory, found by walking up from the test binary. Used
    /// only when regenerating, so a build that cannot find it is never a test failure.
    /// </summary>
    private static string SourceDirectory()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "XLibur.Tests.csproj")))
                return directory.FullName;

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException(
            $"XLibur.Tests.csproj is not above {AppContext.BaseDirectory}, so there is nowhere to " +
            "write the golden fixtures. Regenerate them from a normal local build.");
    }
}
