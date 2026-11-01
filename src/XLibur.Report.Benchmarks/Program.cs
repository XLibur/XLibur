using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;
using BenchmarkDotNet.Toolchains.InProcess.Emit;
using XLibur.Fonts.SkiaSharp;

namespace XLibur.Report.Benchmarks;

/// <summary>
/// Report-generation benchmarks.
/// </summary>
/// <remarks>
/// <code>
/// dotnet run -c Release --project XLibur.Report.Benchmarks -f net10.0 -- scaling
/// dotnet run -c Release --project XLibur.Report.Benchmarks -f net10.0 -- scaling 1000 10000
/// dotnet run -c Release --project XLibur.Report.Benchmarks -f net10.0 -- --filter "*ReportGenerate*"
/// </code>
/// <para>
/// Reach for <c>scaling</c> first. It answers the question spec 12 asks — does the cost per row stay
/// flat as the row count grows — in a couple of minutes. A full BenchmarkDotNet run over these
/// workloads takes the better part of an hour, because it iterates a workload that already takes
/// seconds; it is the right tool for comparing two implementations at one size, not for reading the
/// shape of a curve.
/// </para>
/// The report package is its own project for the same reason it is its own package: nothing in
/// <c>XLibur.Benchmarks</c> has to know that reporting exists.
/// </remarks>
public static class Program
{
    public static void Main(string[] args)
    {
        // Column and row fitting measure text, and a workbook saved without a font engine registered
        // throws rather than guessing.
        SkiaSharpFontBootstrap.Register();

        if (args.Length > 0 && args[0].Equals("scaling", System.StringComparison.OrdinalIgnoreCase))
        {
            ScalingProbe.Run(args);
            return;
        }

        if (args.Length > 0 && args[0].Equals("phases", System.StringComparison.OrdinalIgnoreCase))
        {
            ExpansionPhaseProbe.Run(args);
            return;
        }

        // InProcessEmitToolchain, matching XLibur.Benchmarks: the default CsProj toolchain breaks when
        // duplicate project files exist in the repo (a git worktree will do it), and skipping the
        // per-benchmark project regeneration is faster anyway.
        var config = DefaultConfig.Instance
            .AddJob(Job.Default.WithToolchain(InProcessEmitToolchain.Instance));

        BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, config);
    }
}
