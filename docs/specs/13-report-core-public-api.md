# Spec 13 — A Public Core Surface for `XLibur.Report`

**Area:** Arch · API · Packaging  **Effort:** M  **Status:** Accepted
**Parallelizable?** Tasks 1 and 2 are fully independent; task 3 needs both.

> **Amended 2026-08-01.** Four corrections after a review against the merged tree:
> §B gained a source-kind discriminator (the original three members could not tell an
> unresolvable *source kind* from a named source that resolved to nothing, which silently
> changes `PivotRewriter`'s error behaviour); a third consumer file turned up that this spec
> did not list, needing the §C addition; the version numbers were renumbered from the stale
> 0.107.0 to a 0.201.0 floor; and the `XLiburPackAgainstReleasedCore` / `XLiburCoreFloor`
> build switches this spec cited never existed in the tree.

## Summary

`XLibur.Report` reads core internals through an `InternalsVisibleTo` grant. That is fine while
the two ship as one version off one tag, and it is what spec 12 assumed. It stops being fine
the moment Report versions independently, which is now the intent: Report moves on its own
cadence, is pre-1.0, and should neither inherit core's version nor drag core along when it
breaks.

This spec replaces the friend grant with a small, supported public surface — **two additions,
carried almost entirely by types that are already public** — so Report can depend on a
published core package with an honest version floor. Only two new public types are introduced:
`XLFunctionLibrary` and the `XLPivotSourceKind` enum.

The core additions and Report's refactor onto them are one branch, because both projects live
in this repository — see the work plan.

## Current state

`XLibur/Properties/AssemblyInfo.cs` grants:

```csharp
[assembly: InternalsVisibleTo("XLibur.Report")]
[assembly: InternalsVisibleTo("XLibur.Report.Tests")]
```

The grant is consumed by three files:

| File | Internals used |
|---|---|
| `XLibur.Report/Functions/ExcelFunctionBridge.cs` | `XLCalcEngine`, `XLCalcEngine.Functions`, `FunctionRegistry`, `FunctionDefinition`, `AnyValue`, `ScalarValue`, `CalcContext`, `MissingContextException` |
| `XLibur.Report/Rewriting/PivotRewriter.cs` | `XLPivotCache`, `XLPivotCache.Source`, `XLPivotSourceReference` (`UsesName`, `Area`, `Name`, `TryGetSource`), `SheetArea`, `Area` |
| `XLibur.Report/Ranges/RangeExpander.cs` | `XLConditionalFormat.SetAreas`, `XLAreaList.FromRanges` |

The third was missed when this spec was first written — grep for the *named types* it listed does
not find it, because it reaches for a method on a cast rather than for a distinctive type name.
Only removing the grant and building found it. That is the lesson for the acceptance criteria: the
compiler is the authority on what the grant is load-bearing for, not a search.

Report also *added* two core members to serve the bridge, which this spec supersedes:

- `FunctionRegistry.Names` (new `public` member on an `internal` class)
- `XLCalcEngine.Functions` (new `internal` property)

`XLibur.Report.Tests` additionally reads `XLPivotCache.RecordCount` and assigns
`XLPivotCache.Source`. That grant stays — see *The test-assembly grant stays* below.

There is no build switch that compiles Report against a released core package: both projects
live in this repository and Report holds a plain `ProjectReference`. The cost of the grant is
therefore not a build failure you can reproduce, but the **exact** version range it forces —
`XLiburDependencyVersion` is `[0.200.0]` in `XLibur.Report.props`, and that file's own comment
names this spec as the thing that lets it become an open floor.

## Why an open version floor over internals is unsafe

Internals carry no compatibility contract, so `XLibur >= 0.201.0` would be a promise core never
made. The failure is not hypothetical, because Report is an addon *to* core and consumers will
routinely reference both:

```xml
<PackageReference Include="XLibur" Version="0.215.0" />
<PackageReference Include="XLibur.Report" Version="0.201.0" />  <!-- built against 0.201.0 internals -->
```

NuGet unifies core upward to 0.215.0. Report was compiled against internals that may since have
been renamed, resharpened or deleted, and fails at runtime with `MissingMethodException` or
`TypeLoadException`. Restore is silent; the break surfaces in production.

The alternative — pinning Report to one exact core version — is lockstep wearing a different
number, and makes every core release a Report release. Hence a real public surface.

## Design

Guiding principle: **expose a purpose-built API in terms of types that are already public**
(`XLCellValue`, `XLError`, `IXLRange`, `IXLWorksheet`), rather than publishing the internals as
they stand. `AnyValue`, `ScalarValue`, `CalcContext`, `FunctionDefinition`, `SheetArea` and
`Area` are deep calc-engine and coordinate representations; publishing them freezes
implementation detail permanently for the sake of two call sites.

### A. Invoking a workbook function without a grid

New file `XLibur/Excel/CalcEngine/XLFunctionLibrary.cs`.

```csharp
namespace XLibur.Excel.CalcEngine;

/// <summary>
/// The workbook function library, callable outside a worksheet — the same functions a cell
/// formula can call, evaluated without a grid to be relative to.
/// </summary>
/// <remarks>
/// An instance holds no per-call state and is safe for concurrent use. Construct one per
/// culture and share it; constructing one per call rebuilds the whole function table.
/// </remarks>
public sealed class XLFunctionLibrary
{
    /// <param name="culture">Culture for parsing and formatting. Defaults to invariant.</param>
    public XLFunctionLibrary(CultureInfo? culture = null);

    /// <summary>Names of every available function, in Excel's own casing.</summary>
    public IReadOnlyCollection<string> Names { get; }

    /// <summary>
    /// Calls <paramref name="name"/> with <paramref name="arguments"/>.
    /// </summary>
    /// <returns>
    /// <c>false</c> if no function has that name, leaving <paramref name="result"/> default.
    /// Otherwise <c>true</c>; a call that was made but could not succeed — wrong arity, wrong
    /// argument type, a division by zero — returns <c>true</c> with an <see cref="XLError"/>
    /// result, which is how Excel itself reports these.
    /// </returns>
    /// <exception cref="XLNoWorksheetContextException">
    /// The function needs a worksheet to be relative to (<c>ROW</c>, <c>OFFSET</c>,
    /// <c>INDIRECT</c> and the like). Those belong in a real cell formula.
    /// </exception>
    public bool TryInvoke(string name, ReadOnlySpan<XLCellValue> arguments, out XLCellValue result);
}
```

Also new: `XLNoWorksheetContextException` in `XLibur.Excel.CalcEngine.Exceptions`, a public
wrapper over the internal `MissingContextException`.

Why `XLCellValue` is the right currency:

- It already models every scalar the bridge converts to and from — blank, logical, number, text,
  error — so `AnyValue` and `ScalarValue` stay internal.
- Array and reference results have no `XLCellValue` representation, so the case Report already
  rejects becomes unrepresentable rather than needing a public `AnyValue`. `TryInvoke` returns
  `XLError.IncompatibleValue` for them, matching Report's current behaviour.
- Callers keep their own domain conversions. Report's `DateTime` → OA-date mapping is Report's
  concern and stays in Report.

Implementation is a thin adapter over what already exists: construct an `XLCalcEngine`, look the
name up in its `FunctionRegistry`, arity-check against `MinParams`/`MaxParams`, build the
no-grid `CalcContext`, convert in and out, and translate `MissingContextException`. Conversion
needs nothing new — `ScalarValue` already has an implicit `XLCellValue` operator and an internal
`ToCellValue()`. The `FunctionRegistry.Names` and `XLCalcEngine.Functions` members Report added
stay `internal` and serve this class instead of serving Report directly.

**Thread-safety is a requirement, not an accident.** `ExcelFunctionBridge` deliberately shares
one engine and one adapter set across every template and thread, because importing ~400
functions per `XLTemplate` cost about 12 MB — more than nine tenths of what generating a small
report allocated. A per-call or non-shareable `XLFunctionLibrary` would silently give that back.
The sharing is sound for the same reason it is sound today: these calls have no grid, so
`CalcContext.CalcEngine` is never reached, and the engine is passed only to satisfy the
`CalcContext` constructor. Task 1 must state this in the XML docs and cover it with a
concurrent-invocation test.

### B. Re-pointing a pivot cache's source

`IXLPivotCache` is **already public** and already exposes `Refresh()` and
`SetRefreshDataOnOpen()`. The only gap is the source reference, currently
`internal IXLPivotSource XLPivotCache.Source { get; set; }`.

`PivotRewriter` makes a **three-way** decision per cache, and the surface has to preserve all
three:

1. the source is not a sheet reference at all (a connection, a consolidation, an external
   workbook, a scenario) → leave the cache exactly as the template had it, **silently**;
2. the source is a name or table that no longer resolves → leave it alone but **report a
   template error**, naming the source;
3. the source resolves → re-point it if it is a direct area, then refresh.

A nullable-only surface collapses 1 and 2 into the same "null" and makes every
connection-sourced pivot in a template emit a spurious *"source data could not be read"*. So
the discriminator is explicit:

```csharp
namespace XLibur.Excel;

/// <summary>What a pivot cache reads from. Only <see cref="Range"/> and <see cref="Name"/>
/// resolve to a worksheet; XLibur cannot read the others.</summary>
public enum XLPivotSourceKind
{
    /// <summary>A direct cell area on a sheet in this workbook.</summary>
    Range,

    /// <summary>A table or a book-scoped defined name.</summary>
    Name,

    Consolidation,
    Scenario,
    ExternalWorkbook,
    Connection,
}

public interface IXLPivotCache            // additions only, no changes to existing members
{
    /// <summary>What kind of source this cache reads from.</summary>
    XLPivotSourceKind SourceKind { get; }

    /// <summary>
    /// The range this cache reads from. Non-null only when <see cref="SourceKind"/> is
    /// <see cref="XLPivotSourceKind.Range"/> and that range's sheet still exists.
    /// </summary>
    IXLRange? SourceRange { get; }

    /// <summary>
    /// The table or defined name this cache reads from. Non-null exactly when
    /// <see cref="SourceKind"/> is <see cref="XLPivotSourceKind.Name"/> — the name is what the
    /// template recorded, whether or not it still resolves.
    /// </summary>
    string? SourceName { get; }

    /// <summary>
    /// The worksheet this cache reads from, resolved through the table or defined name when the
    /// source is one. Null when the source does not resolve — a deleted name, a missing sheet —
    /// or when <see cref="SourceKind"/> is a kind XLibur cannot read.
    /// </summary>
    IXLWorksheet? SourceWorksheet { get; }

    /// <summary>
    /// Re-points the cache at <paramref name="range"/>, making <see cref="SourceKind"/>
    /// <see cref="XLPivotSourceKind.Range"/> whatever it was before. Does not refresh.
    /// </summary>
    IXLPivotCache SetSourceRange(IXLRange range);
}
```

This maps one-to-one onto what `PivotRewriter` does today: `SourceKind` replaces the
`cache.Source is not XLPivotSourceReference` type test *and* `XLPivotSourceReference.UsesName`,
`SourceWorksheet` replaces the `SourceSheetName` helper (which calls the internal
`TryGetSource`), `SourceName` replaces `reference.Name` in the error message, and
`SourceRange`/`SetSourceRange` replace reading and assigning `XLPivotSourceReference` over
`SheetArea`. Report's expansion arithmetic then works in `IXLRange` and never needs `Area`.

Why an enum rather than a `bool HasSheetSource`: the five internal `IXLPivotSource`
implementations are a closed set that the file format itself defines, a bool would have to be
replaced the first time a caller wants to tell a connection from an external workbook, and the
enum costs one trivially-stable public type. `XLPivotSourceKind.Range` and `.Name` collapse
`XLPivotSourceReference`'s `UsesName` flag into the same enum rather than adding a second
discriminator, which is why there are six members for five implementations.

Adding members to a public interface is a source-breaking change for external implementers.
`IXLPivotCache` is not designed to be implemented outside the library — `XLPivotCache` is its
only implementation and the constructor is internal — so this is acceptable pre-1.0. Note it in
the changelog under Breaking Changes.

### C. Rewriting a conditional format's coverage

`RangeExpander` widens one template rule over every block generated from it, rather than leaving
the copy-per-block that expansion would otherwise produce (upstream ClosedXML.Report issue #216 —
three rules over three rows become nine). `IXLConditionalFormat.Ranges` is a *fresh projection* of
the rule's internal area list, so adding to what it returns does nothing, and there is no public
way to replace coverage wholesale. Hence one method, the public counterpart of the internal
`SetAreas`:

```csharp
public interface IXLConditionalFormat        // addition only
{
    /// <summary>Replaces the ranges this rule covers.</summary>
    IXLConditionalFormat SetRanges(IEnumerable<IXLRange> ranges);
}
```

It validates what the internal path did not have to: a non-empty set, all on the rule's own
worksheet. Areas are bare rectangles interpreted against that sheet, so a range from elsewhere
would silently *move* the rule rather than being rejected. `XLAreaList` and `Area` stay internal.

### What stays internal (non-goals)

`XLCalcEngine`, `FunctionRegistry`, `FunctionDefinition`, `CalcContext`, `AnyValue`,
`ScalarValue`, `MissingContextException`, `XLPivotCache`, `XLPivotSourceReference`,
`IXLPivotSource`, `SheetArea`, `Area`, `XLAreaList`, `XLConditionalFormat`. If a task finds itself
widening any of these, the design above is being bypassed — stop and revise the spec instead.

Out of scope: registering *custom* functions with the calc engine, evaluating whole formula
strings outside a workbook, and any other pivot-source kind (consolidation ranges, external
sources). None are needed by Report and each is a larger design in its own right.

### The test-assembly grant stays

Drop `InternalsVisibleTo("XLibur.Report")`. **Keep `InternalsVisibleTo("XLibur.Report.Tests")`**,
with a comment recording that the asymmetry is deliberate.

The compatibility contract that matters is the *shipped package's*. `XLibur.Report.Tests` is
`IsPackable=false`, lives in this repo, and always builds against the core in this tree, so its
use of internals — asserting on a pivot cache's `RecordCount`, which is not on the public
surface and should not be — costs nothing and constrains nothing. Publishing a record count
merely to avoid an asymmetry would be the tail wagging the dog.

## Work plan

| # | Task | Files | Depends on |
|---|---|---|---|
| 1 | `XLFunctionLibrary` + `XLNoWorksheetContextException`, with tests | `XLibur/Excel/CalcEngine/XLFunctionLibrary.cs`, `.../Exceptions/`, `XLibur.Tests/Excel/CalcEngine/` | — |
| 2 | `XLPivotSourceKind` + `IXLPivotCache.SourceKind` / `SourceRange` / `SourceName` / `SourceWorksheet` / `SetSourceRange`, with tests | `XLibur/Excel/PivotTables/XLPivotSourceKind.cs`, `IXLPivotCache.cs`, `XLPivotCache.cs`, `XLibur.Tests/Excel/PivotTables/` | — |
| 3 | `IXLConditionalFormat.SetRanges`, with tests | `XLibur/Excel/ConditionalFormats/IXLConditionalFormat.cs`, `XLConditionalFormat.cs`, `XLibur.Tests/Excel/ConditionalFormats/` | — |
| 4 | Refactor `ExcelFunctionBridge`, `PivotRewriter` and `RangeExpander` onto the new surface; revert `FunctionRegistry.Names` / `XLCalcEngine.Functions` to plain internals | `XLibur.Report/**` | 1, 2, 3 |
| 5 | Drop the `XLibur.Report` friend grant, keep the `.Tests` one, and open the version floor in `XLibur.Report.props`; add the surface-guard test | `XLibur/Properties/AssemblyInfo.cs`, `XLibur.Report.props`, `XLibur.Tests/` | 4 |
| 6 | Release core **0.201.0** carrying all of the above | — | 5 |

Tasks 1, 2 and 3 are disjoint and can run in parallel; 4 then 5 are strictly sequential.

Both projects live in this repository and Report holds a `ProjectReference`, so — unlike what
this spec first assumed — there is no cross-branch sequencing problem: task 4 simply will not
compile until task 3 has landed, and the compiler is the check. The whole sequence is one
branch.

## Acceptance criteria

1. `XLibur.Report` compiles with `XLibur`'s friend grant removed — that is, the whole solution
   builds. This is the check that the public surface is sufficient.
2. `XLibur/Properties/AssemblyInfo.cs` contains no `InternalsVisibleTo("XLibur.Report")`.
   The `XLibur.Report.Tests` grant is still present and commented.
3. Every type named under *What stays internal* is still `internal`. Assert this in a test that
   reflects over the public surface, so a later change cannot widen one by accident.
4. The `XLibur.Report` test suite passes unchanged — the refactor is behaviour-preserving.
   No Report test may be edited to accommodate the new API except where it named an internal type
   directly.
5. `XLFunctionLibrary` covers what the bridge needs: `Names` is non-empty and includes `SUM`;
   `TryInvoke("SUM", [1, 2, 3])` yields `6`; an unknown name returns `false`; wrong arity returns
   `true` with `XLError.IncompatibleValue`; `ROW` throws `XLNoWorksheetContextException`; and a
   round-trip test per scalar kind (blank, logical, number, text, error). One test invokes
   concurrently from several threads on a shared instance.
6. A pivot cache loaded from a file reports `SourceKind == Range` with a non-null `SourceRange`
   for a range source, and `SourceKind == Name` with a non-null `SourceName` and `SourceWorksheet`
   for a named-range source; a cache whose name has been deleted reports `SourceKind == Name`
   with a non-null `SourceName` and a **null** `SourceWorksheet`; `SetSourceRange` followed by
   `Refresh()` re-reads from the new range. Verify against a real workbook, per the repo's
   file-format rule.
7. `IXLConditionalFormat.SetRanges` replaces coverage, returns the format, and rejects null, an
   empty set, and a range from another worksheet. One test asserts that mutating what `Ranges`
   returns does nothing — the fact that makes this method necessary — and one round-trips the
   coverage through a save and load.
8. `XLiburDependencyVersion` in `XLibur.Report.props` is an open floor (`0.201.0`), not the
   exact `[0.200.0]` range, and the comment above it no longer defers to this spec. This is the
   check that proves the lockstep is gone.

## Risks

- **The public surface is permanent.** `TryInvoke`'s signature in particular is hard to widen
  later. Review the shape before implementing; a wrong shape here is worse than the friend grant.
- **`ReadOnlySpan<XLCellValue>` cannot be used in an `async` method or captured.** Report's bridge
  is synchronous, so this is free today; if a future caller needs otherwise, add an
  `IReadOnlyList<XLCellValue>` overload rather than changing this one.
- **`XLCellValue` round-tripping may lose fidelity** for values the calc engine represents more
  precisely than a cell can. Task 1 must include a test per scalar kind. If a genuine loss is
  found, record it here rather than reaching for `AnyValue`.
- **Interface additions break external implementers** of `IXLPivotCache` (see §B). Judged
  acceptable pre-1.0; changelog it.
- **A behaviour regression here is silent.** No Report test covers a connection, consolidation,
  external-workbook or scenario source, so if the surface cannot express case 1 of §B's three-way
  decision, nothing fails — the spurious template errors only show up in a user's report. Task 2
  must test all four unreadable kinds directly, since the Report suite will not catch them.
- **Report cannot release until core 0.201.0 is out.** This spec is a hard prerequisite for
  Report's independent version stream, not a parallel workstream.

## References

- Spec 12 — [Report templating](12-report-templating.md), *Core APIs this depends on*, which
  documents the internal call patterns being replaced. Its Summary states Report's only touch on
  core is an `InternalsVisibleTo` grant; this spec is what makes that no longer true.
- `XLibur.Report.props` — `XLiburDependencyVersion`, the exact `[0.200.0]` range this spec
  turns into an open floor, and the comment that defers to this spec for doing so.
- NuGet [dependency resolution](https://learn.microsoft.com/nuget/concepts/dependency-resolution)
  — lowest-applicable wins, and why a floor is a floor and not a pin.
