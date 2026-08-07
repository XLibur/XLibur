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
    // The fast, GC-exact profiling modes are:
    //   alloc       an allocation report for the save path, split into create and save phases
    //   create      the create phase broken down per API call
    //   streaming   peak heap for the forward-only writer measured against the in-memory
    //               model, optionally taking a row count
    //   structural  the cost of repeated row inserts split into its range-shift and
    //               formula-shift halves
    //   template    the open->edit->save round trip of an existing workbook, split into parse
    //               and serialise; optionally takes a path to a real .xlsx template
    // Every other mode attaches dotMemory and targets the load path.
    if (args.Length > 1 && args[1].Equals("alloc", StringComparison.OrdinalIgnoreCase))
        SaveAllocationProfile.Run();
    else if (args.Length > 1 && args[1].Equals("create", StringComparison.OrdinalIgnoreCase))
        CreatePhaseProbe.Run();
    else if (args.Length > 1 && args[1].Equals("streaming", StringComparison.OrdinalIgnoreCase))
        StreamingMemoryProfile.Run(args);
    else if (args.Length > 1 && args[1].Equals("structural", StringComparison.OrdinalIgnoreCase))
        StructuralEditProfile.Run();
    else if (args.Length > 1 && args[1].Equals("template", StringComparison.OrdinalIgnoreCase))
        TemplateRoundTripProfile.Run(args);
    else if (args.Length > 1 && args[1].Equals("shiftercorpus", StringComparison.OrdinalIgnoreCase))
        ShifterCorpusDump.Run();
    else if (args.Length > 1 && args[1].Equals("compression", StringComparison.OrdinalIgnoreCase))
        CompressionProfile.Run();
    else if (args.Length > 1 && args[1].Equals("loadalloc", StringComparison.OrdinalIgnoreCase))
        LoadDecompositionProfile.Run();
    else if (args.Length > 1 && args[1].Equals("dirtyread", StringComparison.OrdinalIgnoreCase))
        DirtyFormulaReadProfile.Run();
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
