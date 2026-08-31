using System.Runtime.CompilerServices;
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

        var replayPath = Environment.GetEnvironmentVariable(ReplayEnvironmentVariable);
        return !string.IsNullOrWhiteSpace(replayPath)
            ? Replay(target, replayPath)
            : Fuzz(target);
    }

    /// <summary>
    /// Run the target under libFuzzer.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <b>Nothing in this program may name <c>XLibur.Fonts.SixLabors.V1</c> from a method that
    /// runs before <see cref="Fuzzer.LibFuzzer"/> starts.</b> That is a hard constraint, not a
    /// style preference, and it is why registration sits inside the callback rather than at the
    /// top of <c>Main</c>.
    /// </para>
    /// <para>
    /// That assembly carries a <c>[ModuleInitializer]</c> which registers the font engine as soon
    /// as the assembly loads, and registering reads <c>LoadOptions.DefaultFontEngine</c> — code
    /// inside <c>XLibur.dll</c>, which SharpFuzz has rewritten to report coverage. The rewritten
    /// code dereferences a trace buffer that <c>Fuzzer.LibFuzzer.Run</c> is what allocates. Load
    /// the fonts assembly first and the process dies during module initialisation, before a
    /// single input is fuzzed:
    /// </para>
    /// <code>
    /// TypeInitializationException -> NullReferenceException
    ///   at XLibur.Excel.LoadOptions.get_DefaultFontEngine()
    ///   at SixLaborsV1FontBootstrap.Register()
    ///   at ModuleInit.Initialize()
    /// </code>
    /// <para>
    /// The trap is that a <em>reference</em> is enough — the branch need never execute. An earlier
    /// version put the callback registration in exactly the right place but still called
    /// <c>Register()</c> in <c>Main</c>'s replay branch, and JIT-compiling <c>Main</c> loaded the
    /// assembly anyway. Hence the split into separate methods, and hence <c>NoInlining</c>: if
    /// this body were inlined back into <c>Main</c>, the reference would return with it.
    /// </para>
    /// <para>
    /// Failure is near-silent. libFuzzer sees only an exit code and then waits for a target that
    /// is already gone, ignoring its own <c>-max_total_time</c>; one run sat for 25 minutes on a
    /// 600-second budget with no output and no corpus growth. Replay cannot catch it either,
    /// because replay runs an uninstrumented build. <c>fuzz.ps1</c>'s watchdog exists for this.
    /// </para>
    /// </remarks>
    [MethodImpl(MethodImplOptions.NoInlining)]
    private static int Fuzz(string target)
    {
        var fontEngineRegistered = false;
        Fuzzer.LibFuzzer.Run(data =>
        {
            if (!fontEngineRegistered)
            {
                SixLaborsV1FontBootstrap.Register();
                fontEngineRegistered = true;
            }

            FuzzTargets.Run(target, data);
        });

        return 0;
    }

    /// <summary>
    /// Run the target over saved inputs and report one line per input, plus a summary grouping
    /// them by failure. The grouping is the point: six crash artifacts that are all one bug
    /// should read as one bug.
    /// </summary>
    [MethodImpl(MethodImplOptions.NoInlining)]
    private static int Replay(string target, string path)
    {
        // Safe here, and only here: replay runs an uninstrumented build, so loading the fonts
        // assembly touches no rewritten XLibur code. See the remarks on Fuzz for why this call
        // cannot be hoisted into Main, even into a branch that never executes.
        SixLaborsV1FontBootstrap.Register();

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
                signature = StackSummary.Signature(exception);
                Console.WriteLine($"{name}\t{bytes.Length} bytes\t{exception.GetType().FullName}\t{exception.Message}");
                Console.WriteLine($"    {StackSummary.FirstMeaningfulFrame(exception)}");
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

}
