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
    //   bulkedit    the dependency-tree MarkDirty walk's cost on repeated single-cell value
    //               writes, with and without dependents to walk
    //   template    the open->edit->save round trip of an existing workbook, split into parse
    //               and serialise; optionally takes a path to a real .xlsx template
    // Every other mode attaches dotMemory and targets the load path.
    // Lowercased once rather than per arm: the mode names are ASCII, so this preserves the
    // case-insensitive match the eleven separate OrdinalIgnoreCase comparisons gave.
    switch (args.Length > 1 ? args[1].ToLowerInvariant() : string.Empty)
    {
        case "alloc": SaveAllocationProfile.Run(); break;
        case "create": CreatePhaseProbe.Run(); break;
        case "streaming": StreamingMemoryProfile.Run(args); break;
        case "structural": StructuralEditProfile.Run(); break;
        case "template": TemplateRoundTripProfile.Run(args); break;
        case "shiftercorpus": ShifterCorpusDump.Run(); break;
        case "compression": CompressionProfile.Run(); break;
        case "loadalloc": LoadDecompositionProfile.Run(); break;
        case "dirtyread": DirtyFormulaReadProfile.Run(); break;
        case "hyperlinks": HyperlinkScalingProfile.Run(); break;
        case "bulkedit": BulkEditDirtyWalkProfile.Run(); break;
        default: MemoryProfile.Run(args); break;
    }

    return;
}

// Use InProcessEmitToolchain by default. The default CsProj-based toolchain breaks
// when there are duplicate project files in the repo (e.g., a git worktree), and
// in-process is faster anyway since it skips the per-benchmark project regeneration.
var config = DefaultConfig.Instance
    .AddJob(Job.Default.WithToolchain(InProcessEmitToolchain.Instance));

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, config);
