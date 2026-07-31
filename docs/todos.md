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

- **Task #9** — the `SetDataValidation` deprecation is only half-applied.
  `IXLBaseCollection<TSingle, TMultiple>.SetDataValidation()` (`IXLBaseCollection.cs:11`, the
  interface behind `IXLColumns`/`IXLRows`/`IXLCells`) is shipped public API
  (`PublicAPI.Shipped.txt:176`) but carries no `[Obsolete]` and has no `CreateDataValidation()`
  counterpart. Sonar missed it precisely because the attribute is absent.
- **Task #11** — `XLWorksheet.cs:597` reads `"Used {nameof(DefinedName)} instead."`; should be
  "Use", matching its nine siblings.

---

## Task summary

| Task | Item(s) | Kind |
|------|---------|------|
| #1 | 7 | Bug — empty data validations written with empty `sqref` |
| #2 | 10 | Feature — pivot `chartFormats` round trip |
| #3 | 11 | Feature — pivot `filters` round trip |
| #4 | 2, 4 | Correctness — structured references in the dependency tree |
| #5 | 3 | Perf — register data-table formulas in the dependency tree |
| #6 | 1 | Perf — cache parsed reference addresses |
| #7 | 12 | Cleanup — `CacheId` nullability |
| #8 | 17 | Hardening — validate pivot field item values |
| #9 | — | Gap — complete the `SetDataValidation` deprecation |
| #10 | 5, 8, 9, 13–16, 18–27 | Decision — removal release for 17 obsolete members |
| #11 | — | Typo — `"Used"` → `"Use"` |
