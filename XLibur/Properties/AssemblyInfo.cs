using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("XLibur.Tests")]
[assembly: InternalsVisibleTo("XLibur.Benchmarks")]

// XLibur.Report bridges the calc engine's internal FunctionRegistry into template expressions,
// and re-points internal pivot cache sources after a template range expands.
[assembly: InternalsVisibleTo("XLibur.Report")]
