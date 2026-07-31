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

### Two findings surfaced while triaging this part

- **Task #9** — `IXLBaseCollection<TSingle, TMultiple>` (`IXLBaseCollection.cs:5`) is **dead
  public API**. It is shipped (`PublicAPI.Shipped.txt:166-180`, 15 members, including a
  `SetDataValidation()` that carries no `[Obsolete]`), but nothing implements it and nothing
  consumes it: `git grep IXLBaseCollection` outside `PublicAPI.Shipped.txt` returns only its
  own declaration, and `IXLColumns`, `IXLRows`, `IXLCells`, `IXLRangeColumns` and
  `IXLRangeRows` all derive from `IEnumerable<T>` alone. So this is a disposal question that
  belongs with task #10, not a deprecation to complete — adding the missing
  `CreateDataValidation()` counterpart would only grow dead surface.
- **Task #11** — `XLWorksheet.cs:597` reads `"Used {nameof(DefinedName)} instead."`; should be
  "Use", matching its nine siblings.

---

## Task summary

| Task | Item(s) | Kind | Status |
|------|---------|------|--------|
| #1 | 7 | Bug — empty data validations written with empty `sqref` | In review — [#290](https://github.com/XLibur/XLibur/pull/290) |
| #2 | 10 | Feature — pivot `chartFormats` round trip | In review — [#294](https://github.com/XLibur/XLibur/pull/294) |
| #3 | 11 | Feature — pivot `filters` round trip | Not started |
| #4 | 2, 4 | Correctness — structured references in the dependency tree | Not started |
| #5 | 3 | Perf — register data-table formulas in the dependency tree | Not started |
| #6 | 1 | Perf — cache parsed reference addresses | **Merged** — [#286](https://github.com/XLibur/XLibur/pull/286) |
| #7 | 12 | Cleanup — `CacheId` nullability | **Merged** — [#287](https://github.com/XLibur/XLibur/pull/287) |
| #8 | 17 | Hardening — validate pivot field item values | Not started |
| #9 | — | Decision — dispose of the orphan `IXLBaseCollection` interface | **Needs an owner decision** |
| #10 | 5, 8, 9, 13–16, 18–27 | Decision — removal release for 17 obsolete members | **Needs an owner decision** |
| #11 | — | Typo — `"Used"` → `"Use"` | In review — [#292](https://github.com/XLibur/XLibur/pull/292) |

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
| [#290](https://github.com/XLibur/XLibur/pull/290) | Task #1 | Open |
| [#292](https://github.com/XLibur/XLibur/pull/292) | Task #11 | Open |
| [#294](https://github.com/XLibur/XLibur/pull/294) | Task #2 (replaces #291, which was stacked on #287) | Open |

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

### Task #9 — correction

The entry above in Part 2 has been rewritten. The original claim — that this was a half-applied
deprecation on the interface behind `IXLColumns`/`IXLRows`/`IXLCells` — was wrong: nothing
extends `IXLBaseCollection` and nothing implements it. There is no deprecation to complete, only
a disposal decision, which belongs with #10.

### Suggested order for what remains

#3 and #8 are the remaining pivot items and touch the same reader/writer pair, so they conflict
least taken together — though #3 is much the largest job left, since `CT_PivotFilter` carries a
required full `autoFilter` subtree. #4 and #5 are both calc-engine dependency-tree work and pair
naturally. #9 and #10 are decisions rather than code, and #9 does not gate #10 in the way the
earlier note claimed — they are two parts of the same disposal question.
