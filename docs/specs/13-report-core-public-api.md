# Spec 13 — A Public Core Surface for `XLibur.Report`

**Area:** Arch · API · Packaging  **Effort:** M  **Status:** Proposed
**Parallelizable?** Tasks 1 and 2 are fully independent; task 3 needs both.

## Summary

`XLibur.Report` reads core internals through an `InternalsVisibleTo` grant. That is fine while
the two ship as one version off one tag, and it is what spec 12 assumed. It stops being fine
the moment Report versions independently, which is now the intent: Report moves on its own
cadence, is pre-1.0, and should neither inherit core's version nor drag core along when it
breaks.

This spec replaces the friend grant with a small, supported public surface — **two additions,
expressed in types that are already public** — so Report can depend on a published core
package with an honest version floor.

It is a core-only change. Report's own refactor onto the new surface (task 3) is a separate
piece of work on the Report branch and is described here only so the surface can be judged
against its consumer.

## Current state

`XLibur/Properties/AssemblyInfo.cs` grants:

```csharp
[assembly: InternalsVisibleTo("XLibur.Report")]
[assembly: InternalsVisibleTo("XLibur.Report.Tests")]
```

Building Report against the released `XLibur 0.106.0` package instead of the project fails with
**9 errors** (`CS0122` inaccessible, `CS0246` not found) across exactly two files:

| File | Internals used |
|---|---|
| `XLibur.Report/Functions/ExcelFunctionBridge.cs` | `XLCalcEngine`, `FunctionDefinition`, `AnyValue`, `ScalarValue`, `CalcContext`, `MissingContextException` |
| `XLibur.Report/Rewriting/PivotRewriter.cs` | `XLPivotCache`, `XLPivotSourceReference`, `SheetArea`, `Area` |

The Report branch also *added* two core members to serve the bridge, which this spec supersedes:

- `FunctionRegistry.Names` (new `public` member on an `internal` class)
- `XLCalcEngine.Functions` (new `internal` property)

Reproduce the failure from the Report branch:

```
dotnet build XLibur.Report/XLibur.Report.csproj -c Release -p:XLiburPackAgainstReleasedCore=true
```

## Why an open version floor over internals is unsafe

Internals carry no compatibility contract, so `XLibur >= 0.107.0` would be a promise core never
made. The failure is not hypothetical, because Report is an addon *to* core and consumers will
routinely reference both:

```xml
<PackageReference Include="XLibur" Version="0.120.0" />
<PackageReference Include="XLibur.Report" Version="0.1.0" />  <!-- built against 0.107.0 internals -->
```

NuGet unifies core upward to 0.120.0. Report was compiled against internals that may since have
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
no-grid `CalcContext`, convert in and out, and translate `MissingContextException`. The
`FunctionRegistry.Names` and `XLCalcEngine.Functions` members the Report branch added stay
`internal` and serve this class instead of serving Report directly.

### B. Re-pointing a pivot cache's source

`IXLPivotCache` is **already public** and already exposes `Refresh()` and
`SetRefreshDataOnOpen()`. The only gap is the source reference, currently
`internal IXLPivotSource XLPivotCache.Source { get; set; }`. Add three members:

```csharp
public interface IXLPivotCache            // additions only, no changes to existing members
{
    /// <summary>
    /// The range this cache reads from, when the source is a direct range on a sheet.
    /// Null when the source is a named range or a table — use <see cref="SourceWorksheet"/>.
    /// </summary>
    IXLRange? SourceRange { get; }

    /// <summary>
    /// The worksheet this cache reads from, resolved through the named range or table when the
    /// source is one. Null if the source cannot be resolved, e.g. a deleted name.
    /// </summary>
    IXLWorksheet? SourceWorksheet { get; }

    /// <summary>Re-points the cache at <paramref name="range"/>. Does not refresh.</summary>
    IXLPivotCache SetSourceRange(IXLRange range);
}
```

These map one-to-one onto what `PivotRewriter` does today: `SourceWorksheet` replaces its
`SourceSheetName` helper (which calls the internal `TryGetSource`), and
`SourceRange`/`SetSourceRange` replace reading and assigning `XLPivotSourceReference` over
`SheetArea`. Report's expansion arithmetic then works in `IXLRange` and never needs `Area`.

Adding members to a public interface is a source-breaking change for external implementers.
`IXLPivotCache` is not designed to be implemented outside the library — `XLPivotCache` is its
only implementation and the constructor is internal — so this is acceptable pre-1.0. Note it in
the changelog under Breaking Changes.

### What stays internal (non-goals)

`XLCalcEngine`, `FunctionRegistry`, `FunctionDefinition`, `CalcContext`, `AnyValue`,
`ScalarValue`, `MissingContextException`, `XLPivotCache`, `XLPivotSourceReference`,
`IXLPivotSource`, `SheetArea`, `Area`. If a task finds itself widening any of these, the design
above is being bypassed — stop and revise the spec instead.

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
| 2 | `IXLPivotCache.SourceRange` / `SourceWorksheet` / `SetSourceRange`, with tests | `XLibur/Excel/PivotTables/IXLPivotCache.cs`, `XLPivotCache.cs`, `XLibur.Tests/Excel/PivotTables/` | — |
| 3 | Drop the `XLibur.Report` friend grant; keep the `.Tests` one | `XLibur/Properties/AssemblyInfo.cs` | 1, 2 |
| 4 | Release core **0.107.0** carrying all of the above | — | 1, 2, 3 |
| 5 | *(Report branch, separate PR)* refactor `ExcelFunctionBridge` and `PivotRewriter` onto the new surface; revert the `FunctionRegistry.Names` / `XLCalcEngine.Functions` additions to plain internals | `XLibur.Report/**` | 4 |

Tasks 1 and 2 are disjoint and can run in parallel. Task 3 is a two-line deletion that will not
compile until 1 and 2 are complete *and* task 5 has landed on the Report branch — so on core's
own branch, task 3 lands together with a temporary `ProjectReference` build check, or simply
after Report is refactored. Sequence 1‖2 → 5 → 3 → 4 if the Report branch can absorb the
refactor before core releases.

## Acceptance criteria

1. `dotnet build XLibur.Report/XLibur.Report.csproj -c Release -p:XLiburPackAgainstReleasedCore=true`
   against core ≥ 0.107.0 succeeds with **0 errors**. It currently produces 9.
2. `XLibur/Properties/AssemblyInfo.cs` contains no `InternalsVisibleTo("XLibur.Report")`.
   The `XLibur.Report.Tests` grant is still present and commented.
3. Every type named under *What stays internal* is still `internal`. Assert this in a test that
   reflects over the public surface, so a later change cannot widen one by accident.
4. The `XLibur.Report` test suite passes unchanged — the refactor is behaviour-preserving.
   No Report test may be edited to accommodate the new API except where it named an internal type
   directly.
5. `XLFunctionLibrary` covers what the bridge needs: `Names` is non-empty and includes `SUM`;
   `TryInvoke("SUM", [1, 2, 3])` yields `6`; an unknown name returns `false`; wrong arity returns
   `true` with `XLError.IncompatibleValue`; `ROW` throws `XLNoWorksheetContextException`.
6. A pivot cache loaded from a file reports a non-null `SourceRange` for a range source and a
   non-null `SourceWorksheet` for a named-range source; `SetSourceRange` followed by `Refresh()`
   re-reads from the new range. Verify against a real workbook, per the repo's file-format rule.
7. The packed `XLibur.Report.nuspec` declares `<dependency id="XLibur" version="0.107.0" />` —
   the floor, not a MinVer-computed pin. This is the check that proves the lockstep is gone.

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
- **Report cannot release until core 0.107.0 is out.** This spec is a hard prerequisite for
  Report's independent version stream, not a parallel workstream.

## References

- Spec 12 — [Report templating](12-report-templating.md), *Core APIs this depends on*, which
  documents the internal call patterns being replaced. Its Summary states Report's only touch on
  core is an `InternalsVisibleTo` grant; this spec is what makes that no longer true.
  *(That file is on the `feat/spec-12-report-templating` branch — the link resolves once it
  merges.)*
- `XLibur.Report/XLibur.Report.csproj` — `XLiburCoreFloor` (0.107.0) and the
  `XLiburPackAgainstReleasedCore` switch that flips the core reference from project to package.
- NuGet [dependency resolution](https://learn.microsoft.com/nuget/concepts/dependency-resolution)
  — lowest-applicable wins, and why a floor is a floor and not a pin.
