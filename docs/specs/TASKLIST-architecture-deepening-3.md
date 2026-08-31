# Tasklist — Architecture deepening, round 3 (specs 36–51)

Progress board and parallel-execution plan for the sixteen architecture specs that came out of the
**2026-08-30** architecture review. Round 1 (specs 22–25) and round 2 (specs 26–34) have their own
boards in [TASKLIST-architecture-deepening.md](TASKLIST-architecture-deepening.md) and
[TASKLIST-architecture-deepening-2.md](TASKLIST-architecture-deepening-2.md).

**Update this file as tasks land.** Tick the boxes, and put the PR number next to the task.

## What this round found

The same shape as the last two rounds: **one fact has two or more implementations, kept in agreement
by hand.** Round 1 found it twice. Round 2 found it nine times, five already broken. Round 3 found it
fifteen times, and **most of the agreements have already failed in shipped code**.

Every defect below was **executed against a scratch build**, not inferred, except where the spec marks
a finding unverified. The repository was not modified during the review.

| Spec | Drift | Effect |
|---|---|---|
| 36 | Eight implementations of "where is this rectangle"; four disagree | **`SaveAs` throws** on a conditional format over a reversed range — the workbook cannot be saved at all. Plus: data validation lost on round trip, styling silently does nothing, negative row/column counts, tables with zero fields, consolidation dropping ranges |
| 37 | Eight copies of "reduce an argument to a scalar", three behaviours | **All 18 dynamic-array functions and 4 regression functions return `#VALUE!`** when a scalar parameter is a cell reference. `=SEQUENCE(A1)`, `=SORT(data,C1)`, `=XLOOKUP(k,a,b,,C1)`. Five of the eight copies contain a branch that provably never executes |
| 38 | Six enumerations of the sheet's view properties; four have drifted | Copying a sheet loses gridlines, zoom, view mode and tab colour, and picks up the *target* workbook's defaults. `sheetView/@view` is written and never read — Page Layout files reopen as Normal |
| 39 | Sixty pivot settings enumerated five times | `showLastColumn` is read from `showColStripes` — re-saving adds emphasis the user never asked for. `Title`/`Description` are public, settable and persisted nowhere. `CopyTo` resets 23 settings. `dataPosition` never set on load, against the subsystem's own documented crash condition |
| 40 | The dirty flag doubles as the BFS visited-marker | Calling the public `InvalidateFormula` makes a **later** edit under-recalculate — the walk treats an already-dirty node as visited and prunes its whole subtree. `BumpEditEpoch` is never called, so the documented three-state epoch model does not exist |
| 41 | Four serialisers of one pivot cache value union | An error value is written as a **boolean** element — Excel reports the cache as repairable. Grouping is **destroyed** on re-save and the part becomes structurally invalid; **XLibur then refuses to load its own output** |
| 42 | Two formula write paths, one skips invalidation | Replacing an array formula leaves its consumers stale — `SUM` over the array keeps the old total. The single-cell control is correct, which is why it was never noticed |
| 43 | Two readers of "a cell's current value", one spill-blind | Reading a spilled cell gives **a different answer depending on what was read before it**. Third copy of spill ownership maintained by hand |
| 44 | Thirteen validation settings written out **seven times** | Standard and extension writers already disagree on empty-formula emission and on the dirty gate. The consolidation equality test compares one property twice under two names — a dead comparison, and direct evidence the list is copied rather than derived |
| 45 | The text codec applied at 5 of 7 sites | **Save throws** on a control character with string sharing off. A literal `_x0041_` is silently rewritten to `A` by Excel — and XLibur skips the decode too, so its own round trip looks clean and no test can see it |
| 46 | A table writer with no reader module | A table's **colour filter filters by the wrong colour** after a round trip — a `dxfId` into a collection rebuilt every save. Sort state dropped; tables renamed on save; column/header/totals formatting destroyed. The mechanism is an **optional parameter** one call site omits |
| 47 | Four copies of the value-assignment protocol | `InsertData` keeps the leading apostrophe where `Cell.Value` strips it, when the inherited style already carries the quote prefix. The shipped public `SetCellValue` never invalidates and has no tests |
| 48 | Readers written against XLibur's writer, not the format | **Load crashes** on an Excel data bar with "negative fill same as positive". Save throws on a value object with no value. Bordered data bars come back borderless |
| 49 | Four parallel index-aligned collections on a public interface | One threshold with no type shifts every later entry onto the wrong type — an icon set silently re-saves meaning something else |
| 50 | One `Intersection`, two algorithms, two error modes | Disjoint returns a non-null `#REF!` without a predicate and `null` with one. Neither convention is documented or tested |
| 51 | Two consolidation engines, three methods byte-identical | **No divergence found** on normalised input across a 400-case fuzz. Prevention, not a fix — scheduled and presented as such |

**The through-line, again: the missing test is not an oversight, it is a consequence of the shape.**
Round 3 adds a second pattern worth naming — **XLibur's readers are written against XLibur's writer
rather than against the format.** Where the writer always emits an element, the reader assumes it
(48); where both halves of a codec are missing, the error cancels out and only Excel sees it (45);
where a fixture can only be authored by the library, the input that breaks it cannot be constructed
(41, 46). Several of these specs deliver **Excel-authored fixtures** as their real product.

---

## 1. Progress board

| Spec | Title | Effort | Blocked by | Status |
|---|---|---|---|---|
| [36](36-one-normalised-rectangle.md) | One rectangle, normalised once | M | — | 🟩 **Implemented** on `task/36` (`80a5b77a`, 15 commits; green on both TFMs at head — net8.0 14,268 verified by the orchestrator); not yet merged. Residual class recorded as D24. See Results |
| [37](37-scalar-argument-reduction.md) | One way to reduce an argument to a scalar | M | — (**must precede 32**) | 🟩 **Implemented** on `task/37` (`42bf46f8`, 13 commits, 14,285 green ×2 TFMs); not yet merged. See Results |
| [38](38-sheet-view-state.md) | Sheet view state gets one module | M | — (**conflicts with 31**) | 🟩 **Implemented** on `task/38` (`ba9af10d`, 11 commits, 28,516 green ×2 TFMs); not yet merged. Inverted-polarity premise disproved. See Results |
| [39](39-pivot-definition-attribute-table.md) | One attribute table for the pivot definition | M–L | — | 🟩 **Implemented** on `task/39` (`d5d73267`, 14 commits; 28,526 green ×2 TFMs at `f78a2fe8` verified by the orchestrator, net10.0 green at head after the owner's review pass); not yet merged. Eight golden fixtures regenerated. PR must be `feat!`/`fix!`: `ShowLastColumn` on `IXLPivotTable` breaks external implementers |
| [40](40-dirty-versus-visited.md) | The dirty flag stops doubling as a visited marker | S–M | — | 🟩 **Implemented** on `task/40` (`9a070669`, 9 commits, 14,265 green ×2 TFMs); not yet merged. **One open decision**: 4×/17× cost on the unsettled bulk-edit shape, accepted for correctness — see Results |
| [41](41-pivot-cache-value-codec.md) | One codec for a pivot cache value | M | — | ⬜ Ready |
| [42](42-formula-write-invalidation.md) | One formula write path, and it invalidates | S–M | soft: 40 (implemented, unmerged — branch off `task/40` or wait for its merge); **takes D26** | ⬜ Ready |
| [43](43-spill-aware-cell-read.md) | Reading a cell is spill-aware, once | M | soft: 40 (implemented, unmerged), 42 | ⬜ Ready |
| [44](44-data-validation-mapping.md) | Data validation: one mapping, two adapters | M | — (**conflicts with 48/49**) | ⬜ Ready |
| [45](45-text-codec-at-the-seam.md) | The text codec applies at the seam | **S** | — (**conflicts with 31**) | ⬜ Ready |
| [46](46-table-part-reader.md) | The table part gets a reader | M | — | ⬜ Ready |
| [47](47-value-assignment-protocol.md) | One implementation of assigning a value | S–M | soft: 42 | ⬜ Ready |
| [48](48-conditional-format-defects.md) | Conditional format defects | S–M | — (**before 49**) | ⬜ Ready |
| [49](49-conditional-format-value-object.md) | One conditional format value object | M–L | **48** | ⬜ Blocked |
| [50](50-intersection-one-convention.md) | One `Intersection`, one absence convention | **S** | soft: 36 | ⬜ Ready |
| [51](51-one-consolidation-engine.md) | One consolidation engine, two adapters | S–M | **36** (implemented, unmerged — unblocks on merge) | ⬜ Blocked |

## 2. Dependency graph

```
36 ──────────────► 51            (36 removes 51's only live divergence)
36 ┄┄┄┄┄┄┄┄┄┄┄┄┄► 50            (soft: consistent geometry, not required)

37 ──────────────► spec 32       (37 must land BEFORE 32; see conflicts)

40 ──► 42 ──► 43                 (one owner, one branch, this order)
       42 ┄┄► 47                 (soft: 47 calls the seam 42 relocates)

48 ──────────────► 49            (crash fix ships without an API decision)

38, 39, 41, 44, 45, 46  — independent of each other
```

## 3. Conflict map

**Read this before assigning two specs to run in parallel.**

| Pair | Conflict | Resolution |
|---|---|---|
| **37 ↔ spec 32** | 32 rewrites 411 registrations across the same function families 37 touches. Head-on collision | **37 lands first**, or 37 folds into 32's scope. 30 is file-disjoint from both and unaffected |
| **38 ↔ spec 31** | Both work inside the sheet-view writer | Sequence them. 31 is the larger; 38 is easier to rebase. Recommend **38 first** — it is smaller and 31 has not started |
| **45 ↔ spec 31** | 45 touches the cell XML writer and both sheet data writers | Small and surgical (one commit). Land **45 before 31 starts**, or accept a trivial rebase |
| **44 ↔ 48/49** | 44 moves the x14 validation reader *out* of the conditional format reader; 48 and 49 work *inside* it | Sequence. Recommend **44 first** — it removes code from the file the other two then edit |
| **48 ↔ 49** | Same files by design | Strictly sequential, one owner |
| **39 ↔ 41** | Both pivot, **file-disjoint** — definition vs cache | Genuinely parallel, two owners |
| **46 ↔ spec 28** | 46 reads the differential format collection 28 established | 28 has merged. No conflict |
| **40/42/43** | Same calc-engine staleness model | One owner, one branch, in order |

## 4. Suggested waves

**Wave A — parallel, four owners, no conflicts between them:**
36 · 37 · 41 · 46

**Wave B — after A:**
40→42→43 (one owner) · 39 · 45 · 48

**Wave C:**
38 · 44 · 47 · 50 · 51 · 49

**If only one thing is done this round: 37.** It is the widest live defect — an entire shipped feature
family unusable with cell references — and the deepening is unusually clean.

**If a second: 36.** An unconditional `SaveAs` throw is the only failure in the round a user cannot
work around.

**Cheapest real win: 45.** A crash and a silent corruption, fixed by moving two calls to a seam. Days,
not weeks.

---

## 5. Backlog notes — reviewed, deliberately not specced

Four candidates from the round-3 report that did not earn a spec. Recorded so the next review does not
re-walk them.

### 5.1 Extend the write-path agreement harness

**Not a spec because it is spec 31's harness, not new content.** The existing write-path agreement
tests cover two elements — the pane and the column — both claimed by spec 29. Meanwhile the streaming
path emits no `<dimension>`, no `<sheetFormatPr>`, no `_xlnm._FilterDatabase` defined name and no
theme part. All are schema-optional, so the SDK validator passes on every one.

**The theme omission has a visible effect:** a theme-coloured style resolves against Excel's built-in
theme instead of the workbook's, so the same `IXLStyle` renders differently depending on which path
wrote it.

**Action:** add to spec 31's task list — extend the harness to `<dimension>`, `<sheetFormatPr>`, a
cell-level comparison and the shared strings part, then decide per divergence whether streaming should
emit the element or the omission is deliberate.

### 5.2 Name the range index's two adapters

**Worth exploring; no divergence found that is reachable today.** Eight methods branch on whether the
quad-tree exists — two complete implementations of one interface in one class, switched by a hidden
one-way promotion at 20 ranges. They disagree on identity (the list path rejects a duplicate by
reference, the tree path by address) and on removal (all address matches versus one). Neither is
reachable today because the index holds strong references.

**The real finding is the test gap:** every test builds a 10,000-range fixture, so the flat-list
path — the one nearly every real workbook takes — has never been directly exercised.

**Action:** if the promotion threshold is ever changed, or the range repository's weak references ever
allow two live ranges with one address, this becomes urgent. Until then, a note.

**Related:** the quad-tree allocates 128 child quadrants and discards them on every add that does not
descend — which is every whole-row range. Perf, not architecture; belongs with spec 19's survey.

### 5.3 Name the pivot-CF handshake

**Worth exploring; severity rests on reading the schema, not a repro that was run.** A pivot's
conditional format lives in two parts joined only by `priority`. The sheet-side reader drops a rule
with no `@type` before registering it, so a malformed pivot rule — or two rules that both fall back to
the same sentinel priority — fails the **entire workbook load**. On save the handshake depends on an
unwritten ordering constraint: priorities are renumbered, and the pivot part must be written
afterwards. It is, but only because one line happens to precede another.

**Action:** could not construct the malformed input from XLibur's own writer. If an Excel-authored
fixture that triggers it turns up, promote to a spec. The degraded-but-loading failure mode is the
right target.

### 5.4 The nine `*Helper` files are file splits, not seams

**Speculative; structural observation only.** Seven of the nine have exactly one caller — the class
they were carved out of — take that class as their first parameter, and reach straight back into it.
The stated rationale is size: "to keep the main class smaller." Six are named nowhere in the test
project. Deletion test: inline any of them and complexity reappears in exactly *one* caller, which
already knew everything the interface asked for.

**Not "re-inline them."** Where a helper is worth keeping, narrow its interface to the data it actually
needs — an `Area` and a worksheet — which is what would make it testable without building a workbook.

**Exceptions:** the shift and insert helpers are genuine. Spec 26 gave them a real axis seam and tests
name them.

### 5.5 Incidentals

- `XLRanges.AddToNamed(string, XLScope)` **discards its `scope` argument** and hardcodes workbook
  scope. A worksheet-scoped named range silently becomes workbook-scoped. One-line fix, no spec.
- Dead code in the area type: `ShiftOrExtendRight`, `ShiftOrExtendDown` and `ToAreaList` have zero
  callers. `Overlaps` has no production caller and is misnamed — its body is a containment test.
- `ReadingOrder` is a raw enum cast in one styles-writer site and goes through the shared converter in
  another. They agree only because the enum's implicit values happen to match the format's. Latent;
  spec 23/28 territory.
- **Three second mapping sites bypass the shared enum converter.** The largest is the **pivot write
  path's own raw-string enum table** — nine enums duplicated from the shared converter, agreeing today
  including the two spellings that differ only in case, with four converter methods dead as a result.
  **Folded into spec 39**, which is already rewriting how that writer emits attributes. Two smaller
  ones have no spec and are recorded here: **`XLLineStyle` in the drawing part reader** and
  **`XLPictureFormat` in the rich data reader**. Both are one-directional local mappings that the
  shared converter also covers. Small enough to fold into whichever spec next touches those files —
  candidates would be spec 15 (shapes) or spec 17 (picture styling), both open.
- **The shared enum converter itself is clean, and this was checked mechanically.** All 43 pairs are
  symmetric and complete in both directions, verified by reflection against the real
  `DocumentFormat.OpenXml` 3.5.1 enum members rather than by reading the tables. No asymmetries, no
  collisions. The single exception is the `sheetView/@view` mapping, whose load-side call is missing
  entirely — that is spec 38's, not the converter's.
- **There is no `EnumConverterTests.cs`.** The only direct coverage is the lenient-casing path in the
  alignment tests, plus round-trip tests for the one deliberate gap. Given the tables are currently
  symmetric this is a latent risk rather than a defect, but it is why a fourth duplicate mapping could
  be added tomorrow without anything noticing.
- `XLCalcEngine.TryEvaluateSingleCell` and `ApplyFormula` share a 15-line literal duplicate and
  disagree on data tables — the fast path "falls back to the safe general path", which throws. **Not
  reproduced.** Recorded in spec 43 as a lead to confirm or dismiss with evidence.
