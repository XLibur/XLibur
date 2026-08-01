# TODO triage — SonarQube INFO issues (XLibur_XLibur)

Source: SonarQube, `impactSeverities=INFO`, `issueStatuses=OPEN,CONFIRMED`, 27 issues
(10 × `csharpsquid:S1135` TODO tags, 17 × `csharpsquid:S1133` deprecated code).

Each item below was read in context and checked against the surrounding code. Verdicts:

- **Stale** — removed from the code, no follow-up.
- **Partly stale** — comment corrected, remaining work tracked.
- **Genuine** — tracked as a task.

---

## Part 1 — TODO comments (S1135)

| # | Location | Verdict | Task |
|---|----------|---------|------|
| 6 | `Coordinates/SheetPoint.cs:47` | **Stale — removed** | — |
| 3 | `CalcEngine/DependencyTree.cs:98` | **Partly stale — reworded** | #5 |
| 1 | `CalcEngine/AstNode.cs:298` | Genuine (perf) | #6 |
| 2 | `CalcEngine/DependenciesVisitor.cs:194` | Genuine (correctness) | #4 |
| 4 | `CalcEngine/Visitors/FormulaReferences.cs:103` | Genuine (correctness) | #4 |
| 7 | `DataValidation/XLDataValidations.cs:278` | Genuine (**bug**) | #1 |
| 10 | `IO/PivotTableDefinitionPartReader.cs:127` | Genuine (data loss) | #2 |
| 11 | `IO/PivotTableDefinitionPartReader.cs:132` | Genuine (data loss) | #3 |
| 12 | `IO/PivotTableDefinitionPartWriter2.cs:79` | Genuine (trivial) | #7 |
| 17 | `PivotTables/XLPivotReference.cs:56` | Genuine (hardening) | #8 |

### Removed as stale

**#6 `SheetPoint.cs:47`** — `/// TODO: SheetId doesn't work nicely with renames, but will in
the future.` Aspirational, no actionable content, and contradicted by the `<summary>` directly
below it, which already documents the real semantics (a sheet id never changes during the
workbook lifecycle; it is *deletion*, not renaming, that invalidates a point). It was also
malformed — a bare `///` line sitting above the `<summary>` tag rather than inside it.

### Corrected

**#3 `DependencyTree.cs:98`** — was `// TODO: Implement other formulas. Don't throw on data
table or shared formulas.` The "don't throw" half is stale: `CreateFrom` skips unhandled
formula types silently, nothing throws, and `FormulaType.Shared` is marked `// Not used` in
`XLCellFormula.cs:15`. The genuine remainder is that data-table formulas are never registered.
Comment rewritten to say only that; work tracked as task #5.

### Verification notes for the genuine ones

- **#7 (data validations)** is a real defect, not a nicety. `SplitBy` can leave a rule with
  zero areas, `Consolidate` preserves it, and `DataValidationWriter` then emits
  `sqref=""` — invalid per the schema.
- **#10 / #11 (pivot)** are both round-trip data loss. The `chartFormat` *attribute* is
  handled; the `<chartFormats>` and `<filters>` *elements* are not read or written at all.
  Note #11 is not the report/page filters — those are supported via `XLPivotTable.Filters`.
- **#2 and #4** are the same underlying gap (structured references) seen from two sides, so
  they share one task.

---

## Part 2 — Deprecated code (S1133) — items 5, 8, 9, 13–16, 18–27

**None of these are stale.** All 17 are deliberate, correctly-written `[Obsolete]` shims doing
their job. S1133 flags them by design and will keep flagging them until they are deleted.

They are shipped public API (tracked in `PublicAPI.Shipped.txt`), so removal is a breaking
change and a release-planning decision — not something to action item by item. Tracked as a
single decision, **task #10**, covering three groups:

| Group | Members | Locations |
|-------|---------|-----------|
| `NamedRange`/`NamedRanges` → `DefinedName`/`DefinedNames` | 10 | `IXLDefinedNames.cs:10`, `XLDefinedNames.cs:38`, `IXLWorkbook.cs:64,199`, `IXLWorksheet.cs:306,314`, `XLWorkbook.cs:145,275`, `XLWorksheet.cs:120,597` |
| `SetDataValidation` → `GetDataValidation`/`CreateDataValidation` | 5 | `XLCell.cs:973`, `IXLRangeBase.cs:306`, `IXLRanges.cs:60`, `XLRangeBase.cs:1289`, `XLRanges.cs:319` |
| `XLFontCharSet.Hangeul` → `Hangul` | 1 | `IXLFont.cs:66` |

**Explicitly excluded from removal:** `XLColor.NoColor` (`XLColor_Static.cs:284`, item 22) was
deprecated deliberately and recently in PR #232 — see `CHANGELOG.md:182`. It needs a normal
deprecation period before it can be considered.

### Decision (owner, July 2026)

**All three groups above are removed in the next minor `v0.x` release.** The library is still
pre-1.0, so a breaking change costs less now than it will after 1.0, and these shims have all
carried their replacement in the message for some time. `XLColor.NoColor` stays, per the
exclusion above.

`IXLBaseCollection` (below) is *not* part of that removal — it is only being deprecated now, so
its own period runs from this release.

### Two findings surfaced while triaging this part

- **Task #9** — `IXLBaseCollection<TSingle, TMultiple>` (`IXLBaseCollection.cs:5`) is **dead
  public API**. It is shipped (`PublicAPI.Shipped.txt:166-180`, 15 members, including a
  `SetDataValidation()` that carries no `[Obsolete]`), but nothing implements it and nothing
  consumes it: `git grep IXLBaseCollection` outside `PublicAPI.Shipped.txt` returns only its
  own declaration, and `IXLColumns`, `IXLRows`, `IXLCells`, `IXLRangeColumns` and
  `IXLRangeRows` all derive from `IEnumerable<T>` alone. So this is a disposal question that
  belongs with task #10, not a deprecation to complete — adding the missing
  `CreateDataValidation()` counterpart would only grow dead surface.
  **Decided:** the interface is marked `[Obsolete]` as of this release and removed once it has
  served a deprecation period, rather than being removed alongside the groups above. Marking it
  produced zero build warnings, which is itself a second confirmation that nothing uses it —
  `TreatWarningsAsErrors` would have turned any `CS0618` into an error.
- **Task #11** — `XLWorksheet.cs:597` reads `"Used {nameof(DefinedName)} instead."`; should be
  "Use", matching its nine siblings.

---

## Task summary

| Task | Item(s) | Kind | Status |
|------|---------|------|--------|
| #1 | 7 | Bug — empty data validations written with empty `sqref` | **Merged** — [#290](https://github.com/XLibur/XLibur/pull/290) |
| #2 | 10 | Feature — pivot `chartFormats` round trip | **Merged** — [#294](https://github.com/XLibur/XLibur/pull/294) |
| #3 | 11 | Feature — pivot `filters` round trip | **Merged** — [#300](https://github.com/XLibur/XLibur/pull/300); `autoFilter` modelled rather than preserved as a string in [#301](https://github.com/XLibur/XLibur/issues/301) |
| #4 | 2, 4 | Correctness — structured references in the dependency tree | In review — [#297](https://github.com/XLibur/XLibur/pull/297) |
| #5 | 3 | Perf — register data-table formulas in the dependency tree | **Closed, will not do** — see below |
| #6 | 1 | Perf — cache parsed reference addresses | **Merged** — [#286](https://github.com/XLibur/XLibur/pull/286) |
| #7 | 12 | Cleanup — `CacheId` nullability | **Merged** — [#287](https://github.com/XLibur/XLibur/pull/287) |
| #8 | 17 | Hardening — validate pivot field item values | Not started |
| #9 | — | Decision — dispose of the orphan `IXLBaseCollection` interface | **Merged** — [#296](https://github.com/XLibur/XLibur/pull/296) |
| #10 | 5, 8, 9, 13–16, 18–27 | Decision — removal release for 17 obsolete members | Decided — groups 1–3 removed in the next minor `v0.x`; removal not yet done |
| #11 | — | Typo — `"Used"` → `"Use"` | **Merged** — [#292](https://github.com/XLibur/XLibur/pull/292) |

---

## Progress

The first four PRs went out as a stack and are now merged. Everything since branches from `main`
directly — the stack cost a rebase of every branch above each merge, which was not worth it for
changes that mostly touch different files.

| PR | Contents | State |
|----|----------|-------|
| [#285](https://github.com/XLibur/XLibur/pull/285) | This document; the stale TODO removed and the partly stale one reworded | Merged |
| [#286](https://github.com/XLibur/XLibur/pull/286) | Task #6 | Merged |
| [#287](https://github.com/XLibur/XLibur/pull/287) | Task #7 | Merged |
| [#288](https://github.com/XLibur/XLibur/pull/288) | This section | Merged |
| [#290](https://github.com/XLibur/XLibur/pull/290) | Task #1 | Merged |
| [#292](https://github.com/XLibur/XLibur/pull/292) | Task #11 | Merged |
| [#294](https://github.com/XLibur/XLibur/pull/294) | Task #2 (replaces #291, which was stacked on #287) | Merged |
| [#295](https://github.com/XLibur/XLibur/pull/295) | Correction to the task #9 finding | Merged |
| [#296](https://github.com/XLibur/XLibur/pull/296) | Task #9, plus the #9 and #10 decisions | Merged |
| [#297](https://github.com/XLibur/XLibur/pull/297) | Task #4, plus the task #5 correction | Open |

### Task #6 — reference resolution (#286)

`ReferenceNode.GetReference` parsed the address string on every evaluation — a string the
constructor itself generates from the parsed `ReferenceArea`, so parsing it only recovered
what the node already held. It now builds the address from the area directly and memoises the
`Reference`: unconditionally for the sheet-less form, and keyed on the resolved worksheet for
the prefixed form, so replacing a sheet cannot serve the previous address.

Measured with the new `FormulaEvaluationBenchmarks` (20K formula rows, `RecalculateAllFormulas`
per operation, net10.0, Ryzen 9 5950X):

| Shape | Before | After | Time | Allocated |
|---|---|---|---|---|
| `UniqueSameSheet` | 19.51 ms / 15.87 MB | 16.06 ms / 10.38 MB | −17.7% | −34.6% |
| `SharedSameSheet` | 16.17 ms / 16.17 MB | 13.56 ms / 10.38 MB | −16.1% | −35.8% |
| `SharedCrossSheet` | 37.64 ms / 16.18 MB | 32.99 ms / 10.38 MB | −12.4% | −35.8% |

**Found and fixed along the way, not on the triage list:** a reversed range such as
`=SUM(B2:A1)` threw `ArgumentException("Range address must be normalized")` out of the
`Reference` constructor — an unhandled exception reaching the caller rather than a `#REF!`.
The parser returns endpoints in written order and nothing normalized them. The new
construction orders each axis independently, carrying each fixed flag with its own coordinate.

### Task #7 — pivot cache id (#287)

The TODO asked whether `XLPivotCache.CacheId` needed to be nullable. It did not, because it did
not belong on the cache: it was documented as coming from the file, but the reader discards the
`cacheId` it reads, so nothing ever set it on load. Its only writer assigns a position while
rebuilding the workbook `pivotCaches` element and its only reader is the pivot table writer, so
the value belongs to one save. Moved to `SaveContext` as a cache-to-id map; the property is
gone, and so is `PivotSourceCacheId`, which was a loop counter with workbook-level scope.

### Task #1 — empty data validations (#290)

`SplitBy` can strip a rule of every area when a new validation wholly covers it, and the writer
then joined its zero ranges into `sqref=""`. Excel treats that as corruption and repairs the
workbook, dropping *every* validation on the sheet. `ClearRanges` and `RemoveRange` are public
and reach the same state without any splitting, so the fix is at both ends: the split path drops
rules it empties, and the writer skips any rule with no coverage.

The sweep runs after the split loop rather than off the coverage-changed event because `SplitBy`
passes through zero areas transiently, between removing an area and adding back its remainder.

### Task #2 — pivot chartFormats (#294)

The `chartFormats` element — which ties each PivotChart formatting record to the pivot area it
applies to — was never read or written, so manual chart formatting vanished on load/save. Added
`XLPivotChartFormat` plus a reader and writer following the existing `formats`/`conditionalFormats`
pattern. Not to be confused with the `chartFormat` *attribute*, which already worked.

### Task #4 — structured references in the dependency tree (#297)

`DependenciesVisitor` returned nothing for every structured reference, so `=SUM(Table1[Amount])`
registered no precedents at all and kept serving a stale value when the table changed. The
resolution `CalculationVisitor` already performed is now shared by both visitors through
`StructuredReferenceResolver`, which also closed the `FormulaReferences` placeholder — the same
gap from the other side.

The resolver goes through an interface rather than taking workbook/worksheet/point directly, and
that is load-bearing: `CalcContext.Worksheet` and `FormulaAddress` throw when an expression has
no anchoring cell, and the original code only touched them on the paths that need a formula
location. Passing them eagerly broke 33 tests.

### Task #9 — correction

The entry above in Part 2 has been rewritten. The original claim — that this was a half-applied
deprecation on the interface behind `IXLColumns`/`IXLRows`/`IXLCells` — was wrong: nothing
extends `IXLBaseCollection` and nothing implements it. There is no deprecation to complete, only
a disposal decision, which belongs with #10.

### Task #5 — closed, will not do

The triage entry said data-table formulas were "skipped, so their dependents fall back to a full
recalculation", and treated the original TODO's *"don't throw on data table or shared formulas"*
warning as stale. **That was wrong**, and the reasoning matters more than the outcome:

- `DependencyTree.AddFormula` derives precedents by **parsing the formula text**
  (`GetFormulaPrecedents` → `formula.GetAst(engine)` → `engine.Parse(formula.A1)`).
- A data-table formula's text is the placeholder `{TABLE(A1,}`, which is not valid formula
  syntax.
- Parsing it throws `ExpressionParseException: Error at char 1 of '{TABLE(A1,}': Unexpected
  token`. Verified directly rather than reasoned about.

So adding `FormulaType.DataTable` to that chain would throw — exactly what the original TODO
warned about. Registering them properly would need precedents built from `Input1`/`Input2` and
the table's header formulas instead of from an AST.

And the payoff would be small: XLibur does not evaluate data tables at all — there is no `TABLE`
function registered — so a data-table cell's value never changes and its dependents never need
invalidating. The only gain is dropping the full-workbook `Recalculate` that
`XLCalcEngine.TryEvaluateSingleCell` triggers whenever a data-table cell is read.

**Decision (owner): closed, will not do.** Revisit only if that recalculation shows up as a real
performance problem. The comment in `DependencyTree` now records why data tables are skipped.

### Open bug found while investigating #5

`DataTableFormulaFormat` (`XLCellFormula.cs:34`) is `"{{TABLE({0},{1}}}"`, which produces
`{TABLE(A1,}` — the closing parenthesis is missing. This is **user-visible**: `XLCell.FormulaA1`
returns `Formula?.A1`, so anyone reading the formula of a data-table cell gets the malformed
text. It does not affect files, because `SheetDataWriter` writes the `dataTable` element from the
formula's attributes rather than from this string, and nothing parses it.

Not fixed, because the correct target text is a judgement call — Excel displays
`{=TABLE(row_input,col_input)}`, so this needs more than adding the missing bracket. No test
depends on the current value.

### Still to do

- **#10 removal** — the decision is made but the code change is not: 16 members across three
  groups, plus the `PublicAPI.*.txt` and `CHANGELOG.md` updates a breaking change needs, and
  the test files that still call `NamedRange`/`NamedRanges`.
- **#3** — the largest job left. `CT_PivotFilter` carries a required full `autoFilter` subtree,
  so it needs real OpenXML modelling rather than another pass of the `formats` pattern.
- **#8** — smallest, but needs pivot-table/cache context threaded into a static loader to
  validate input, and risks rejecting files Excel accepts. Worth weighing before doing.
- **The `DataTableFormulaFormat` bug** above, if the intended display text can be settled.

Tasks #1, #2, #4, #6, #7, #9 and #11 are done; #5 is closed as will-not-do.
