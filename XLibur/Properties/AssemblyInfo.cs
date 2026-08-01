using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("XLibur.Tests")]
[assembly: InternalsVisibleTo("XLibur.Benchmarks")]

// XLibur.Report itself is deliberately NOT here: it ships as its own package on its own version
// stream, and a package compiled against internals can only ever declare an exact dependency on
// the core it was built against, because internals carry no compatibility contract. It uses the
// public XLFunctionLibrary and IXLPivotCache source members instead — see spec 13.
//
// The asymmetry with its test assembly is deliberate. XLibur.Report.Tests is IsPackable=false and
// always builds against the core in this tree, so its use of internals — asserting on a pivot
// cache's RecordCount, which is not on the public surface and should not be — constrains nothing
// that ships.
[assembly: InternalsVisibleTo("XLibur.Report.Tests")]
