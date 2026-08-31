using SharpFuzz;
using XLibur.Fonts.SixLabors.V1;

namespace XLibur.Fuzz;

/// <summary>
/// Entry point for the fuzzing harness. Two modes:
///
/// <list type="bullet">
/// <item><b>Fuzz</b> — the default. libFuzzer drives the process through SharpFuzz.</item>
/// <item><b>Replay</b> — set <c>XLIBUR_FUZZ_REPLAY</c> to a file or directory and the harness runs
/// the same target over those inputs, printing what each one did. This is triage, and it runs the
/// <em>same</em> oracle as fuzzing on purpose: a triage tool that judges differently from the
/// fuzzer is a tool that disagrees with its own findings.</item>
/// </list>
/// </summary>
internal static class Program
{
    private const string TargetEnvironmentVariable = "XLIBUR_FUZZ_TARGET";
    private const string ReplayEnvironmentVariable = "XLIBUR_FUZZ_REPLAY";

    public static int Main()
    {
        var target = Environment.GetEnvironmentVariable(TargetEnvironmentVariable) ?? FuzzTargets.Workbook;
        if (!FuzzTargets.IsKnown(target))
        {
            Console.Error.WriteLine(
                $"{TargetEnvironmentVariable} must be one of: {string.Join(", ", FuzzTargets.All)}.");
            return 2;
        }

        SixLaborsV1FontBootstrap.Register();

        var replayPath = Environment.GetEnvironmentVariable(ReplayEnvironmentVariable);
        if (!string.IsNullOrWhiteSpace(replayPath))
            return Replay(target, replayPath);

        Fuzzer.LibFuzzer.Run(data => FuzzTargets.Run(target, data));
        return 0;
    }

    /// <summary>
    /// Run the target over saved inputs and report one line per input, plus a summary grouping
    /// them by failure. The grouping is the point: six crash artifacts that are all one bug
    /// should read as one bug.
    /// </summary>
    private static int Replay(string target, string path)
    {
        var files = Directory.Exists(path)
            ? Directory.GetFiles(path).OrderBy(f => f, StringComparer.Ordinal).ToArray()
            : [path];

        if (files.Length == 0)
        {
            Console.Error.WriteLine($"No inputs found at {path}.");
            return 2;
        }

        var distinct = new Dictionary<string, List<string>>(StringComparer.Ordinal);

        foreach (var file in files)
        {
            var name = Path.GetFileName(file);
            var bytes = File.ReadAllBytes(file);

            string signature;
            try
            {
                var outcome = FuzzTargets.Run(target, bytes);
                signature = $"(no failure: {outcome})";
                Console.WriteLine($"{name}\t{bytes.Length} bytes\t{outcome}");
            }
            catch (Exception exception)
            {
                signature = Signature(exception);
                Console.WriteLine($"{name}\t{bytes.Length} bytes\t{exception.GetType().FullName}\t{exception.Message}");
                Console.WriteLine($"    {FirstFrame(exception)}");
            }

            if (!distinct.TryGetValue(signature, out var members))
                distinct[signature] = members = [];

            members.Add(name);
        }

        Console.WriteLine();
        Console.WriteLine($"{files.Length} input(s), {distinct.Count} distinct outcome(s):");
        foreach (var (signature, members) in distinct.OrderBy(p => p.Key, StringComparer.Ordinal))
            Console.WriteLine($"  {signature}  x{members.Count}");

        // A replay reports; it does not pass judgement on the run as a whole. Exit 0 unless the
        // inputs could not be read, so that a triage sweep is never mistaken for a failed build.
        return 0;
    }

    /// <summary>
    /// Group inputs by exception type and originating frame, so that identical bugs collapse and
    /// different bugs do not. Deliberately excludes the message, which often carries input-derived
    /// text and would split one bug into many.
    /// </summary>
    private static string Signature(Exception exception)
    {
        return $"{exception.GetType().FullName} at {FirstFrame(exception)}";
    }

    private static string FirstFrame(Exception exception)
    {
        var stack = exception.StackTrace;
        if (string.IsNullOrEmpty(stack))
            return "(no stack)";

        foreach (var line in stack.Split('\n'))
        {
            var trimmed = line.Trim();
            if (trimmed.Length > 0)
                return trimmed;
        }

        return "(no stack)";
    }
}
