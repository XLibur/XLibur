using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("XLibur.Tests")]
[assembly: InternalsVisibleTo("XLibur.Benchmarks")]

// XLibur.Report bridges the calc engine's internal FunctionRegistry into template expressions,
// and re-points internal pivot cache sources after a template range expands. Its tests assert on
// the same internals — a pivot cache's source and record count are not on the public surface, so
// there is no other way to state what re-pointing one did.
[assembly: InternalsVisibleTo("XLibur.Report")]
[assembly: InternalsVisibleTo("XLibur.Report.Tests")]
