using System;
using System.Linq;

namespace XLibur.Benchmarks;

/// <summary>
/// What the command line of this process says about the job. One definition, read from one place, for
/// <c>Program</c> and for any benchmark config that must know whether <c>--job</c> was given.
/// </summary>
internal static class BenchmarkCommandLine
{
    /// <summary>True when a job was chosen with <c>--job</c> or <c>-j</c> (exact option names).</summary>
    public static bool HasJob(string[] args) =>
        args.Any(a => a is "--job" or "-j" || a.StartsWith("--job=", StringComparison.Ordinal));

    /// <summary>True when <c>--inProcess</c> or <c>-i</c> was given.</summary>
    public static bool HasInProcess(string[] args) =>
        args.Any(a => a is "--inProcess" or "-i" || a.StartsWith("--inProcess=", StringComparison.Ordinal));

    /// <summary>The arguments of this process, without the executable path.</summary>
    public static string[] ProcessArgs() => Environment.GetCommandLineArgs().Skip(1).ToArray();
}
