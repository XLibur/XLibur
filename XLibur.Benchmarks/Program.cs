using System;
using BenchmarkDotNet.Running;
using XLibur.Benchmarks;
using XLibur.Fonts.SixLabors.V1;

SixLaborsV1FontBootstrap.Register();

if (args.Length > 0 && args[0].Equals("profile", StringComparison.OrdinalIgnoreCase))
{
    MemoryProfile.Run(args);
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
