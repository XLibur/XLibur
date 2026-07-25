using System;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;
using BenchmarkDotNet.Toolchains.InProcess.Emit;
using XLibur.Benchmarks;
using XLibur.Fonts.SixLabors.V1;

SixLaborsV1FontBootstrap.Register();

if (args.Length > 0 && args[0].Equals("profile", StringComparison.OrdinalIgnoreCase))
{
    // "profile alloc" is a fast, GC-exact allocation report for the save path, split into
    // create/save phases. The other modes attach dotMemory and target the load path.
    if (args.Length > 1 && args[1].Equals("alloc", StringComparison.OrdinalIgnoreCase))
        SaveAllocationProfile.Run();
    else
        MemoryProfile.Run(args);

    return;
}

// Use InProcessEmitToolchain by default. The default CsProj-based toolchain breaks
// when there are duplicate project files in the repo (e.g., a git worktree), and
// in-process is faster anyway since it skips the per-benchmark project regeneration.
var config = DefaultConfig.Instance
    .AddJob(Job.Default.WithToolchain(InProcessEmitToolchain.Instance));

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, config);
