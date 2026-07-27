# Spec 05 — Structural-Edit & Bulk-Style Scalability (rows insert, range shift, style propagation)

**Area:** Architecture + Performance
**Effort:** L (2–3 weeks; can split into 3 independent PRs)
**Dependencies:** None.
**Status:** ✅ Done — see [Results](#results). Two acceptance criteria were disproved rather than met;
workstream C1 was declined on evidence. Read Results before acting on the design below, which
mis-attributes the cost it set out to fix.

## Summary

Row/column insert-delete and range-wide style changes are the two operations most at odds with the otherwise-excellent slice architecture. Inserting N rows one at a time is O(N × liveRanges·log(liveRanges)) because every insert materializes and sorts **every live range in the worksheet**; a range-wide style change materializes an `XLCell` object **per cell**. Both should be index-driven and slice-driven respectively.

## Current state

1. **Range-shift notification** — `XLibur/Excel/XLWorksheet.cs` (~lines 1166–1234), `NotifyRangeShiftedRows`/`NotifyRangeShiftedColumns`:
   ```csharp
   var rangesToShift = _rangeRepository
       .Where(r => r.RangeAddress.IsValid)
       .OrderBy(r => r.RangeAddress.FirstAddress.RowNumber * -Math.Sign(rowsShifted))
       .ToList();
   ```
   Full materialize + sort of every live range per single insert/delete, then a virtual dispatch per range.
2. **Per-cell formula shifting** — `XLRangeBase.Delete` loops cells calling `ShiftFormulaRows` per cell; `AllocationBenchmarks` shows `ShiftFormulaRows` as the worst micro-benchmark (2.61 ms / 3.6 MB per 1000 iterations).
3. **Style propagation** — `XLibur/Excel/Style/XLStylizedBase.cs` (145 lines): `ModifyStyle` → `GetChildrenRecursively` builds a `HashSet<XLStylizedBase>` of every descendant then `.GroupBy(...)` — one `XLCell` per cell in the range. `SetStyle(..., propagate: true)` recursive-walks `Children`.
4. **Two independent spatial indexes** — `XLibur/Excel/Ranges/Index/XLRangeIndex.cs` (265 lines) uses a QuadTree (`XLibur/Excel/Patterns/`, 391 lines) with a `MinimumCountForIndexing` threshold, below which `Contains`/`GetIntersectedRanges` degrade to linear `Any(...)` scans. The calc engine separately uses `RBush<AreaDependents>` (RBush.Signed package) in `DependencyTree`. Two structures, two maintenance paths.
5. **Misc:** `XLRanges.Ranges` re-enumerates `_indexes.Values.SelectMany(...)` on every access; `XLCellsCollection.FindUsedColumn` (~lines 340–368) builds a LINQ `Concat×4 → Where → Distinct → OrderBy` chain per `FirstColumnUsed`/`LastColumnUsed` call.

## Design — three independent workstreams

### A. Batch + index-driven range shift

1. **Index-driven query:** replace the `_rangeRepository.Where(...).OrderBy(...).ToList()` scan with a query against `XLRangeIndex`: only ranges intersecting or below/right of the shift line can be affected. Ranges fully above an inserted row don't need visiting at all.
2. **Batch semantics:** `InsertRowsAbove(n)` and worksheet-level bulk operations must notify **once** with the full shift delta, not n times. Audit call sites: `XLRow.InsertRowsAbove/Below`, `XLRangeBase.InsertRowsAbove`, delete equivalents, and column twins. If the public API already passes `numberOfRows` down, verify a single notification happens (write a counting test first — it may partially work already).
3. **Formula shifting:** `ShiftFormulaRows` per cell → shift at the `FormulaSlice` level: enumerate formulas in the affected region once, use `ClosedXML.Parser`-based transformation (see `CalcEngine/Visitors/FormulaTransformation.cs`, which already rents `ArrayPool<char>`) and skip formulas with no row-relative references cheaply.

### B. Slice-level bulk style application

`range.Style.Font.Bold = true` over 1M cells should not create 1M `XLCell` facades.

1. Add an internal bulk path on the style machinery: for a rectangular target, iterate `StyleSlice` positions directly (existing `Slice` enumerators), applying the transition `XLStyleValue → XLStyleValue` via the existing 8-slot transition cache on `XLStyleValue`. Cells with no explicit style entry get the (single, computed-once) transitioned inherited style — only written to the slice if it differs from the new inherited value.
2. Rework `XLStylizedBase.ModifyStyle`/`GetChildrenRecursively` so ranges/rows/columns/worksheets dispatch to the bulk path instead of materializing children. Non-cell stylized children (tables, CF) keep the object path.
3. Benchmark: new `BulkStyleBenchmarks` (style 100K×10 range; assert allocations don't scale with cell count).

### C. Spatial-index consolidation (evaluate, then do or document)

Evaluate replacing the QuadTree in `XLRangeIndex` with `RBush.Signed` (already a dependency, already proven in `DependencyTree`):
- Prototype `XLRangeIndex` on RBush; run range-heavy tests + a micro-benchmark (insert/query/remove 10K ranges).
- If RBush wins or ties: delete `XLibur/Excel/Patterns/` QuadTree (~391 lines) — one structure to maintain. If it loses (RBush removal cost is a known weakness): write the finding into this spec and keep QuadTree.
- Also fix `XLRanges.Ranges` re-enumeration (cache or expose a count-only path) and de-LINQ `XLCellsCollection.FindUsedColumn` (fold the 4-slice merge into a simple loop like the existing `ColumnsUsedKeys` cache did).

## Work plan

| # | Task | Size | PR |
|---|------|------|----|
| A1 | Counting test: notifications per `InsertRowsAbove(5)`; fix to single batch notify | S | 1 |
| A2 | Index-driven `NotifyRangeShifted*` | M | 1 |
| A3 | Slice-level formula shift | M | 2 |
| B1 | Bulk style application path + rework `ModifyStyle` | L | 3 |
| B2 | `BulkStyleBenchmarks` | S | 3 |
| C1 | RBush-vs-QuadTree prototype + decision | M | 4 |
| C2 | `FindUsedColumn` / `XLRanges.Ranges` de-LINQ | S | 4 |

## Acceptance criteria

1. Inserting 1,000 rows one-by-one into a sheet with 1,000 live ranges: ≥ 10× faster than main (add a benchmark demonstrating it).
2. `InsertRowsAbove(1000)` (batch) triggers exactly one shift notification pass.
3. Styling a 100K-cell range allocates O(distinct styles), not O(cells) — asserted in benchmark, and `GetChildrenRecursively` no longer enumerates cells for range targets.
4. Zero behavior change: full test suite green, especially `XLibur.Tests` range/insert/delete/named-range/CF-shift tests (recent fixes #158, #142, #303e27c0 are the regression corpus — do not break them).
5. If C1 consolidates: QuadTree deleted, RBush index passes all range-index tests.

## Risks

- Range-shift ordering matters (the `OrderBy` with `-Math.Sign(rowsShifted)` exists so overlapping ranges shift in the right order) — the index-driven replacement must preserve processing order semantics. Characterize with tests before changing.
- `ModifyStyle` rework touches a base class used by everything stylized — do B behind thorough tests; consider a temporary internal feature flag for A/B validation in tests.

## References

- Architecture survey §5 (pain points), `AllocationBenchmarks.cs` (`ShiftFormulaRows`), recent regression-relevant fixes: #158 (CF shift), #142 (named range shrink), #157 (page break extents).

## Results

### The spec attributed the cost to the wrong thing

The first thing built was not a fix but a measurement: `StructuralEditProfile`
(`-- profile structural`). It runs 1,000 single-row inserts and separates the two costs an insert
pays — the range-shift pass and the formula shift — by placing the sheet's content either *below* the
insert point (so the shift reaches it) or *above* it (so it cannot). The gap between those isolates
the cost of **visiting** candidates from the cost of **moving** them. The insert row is held fixed
across probes, because moving it would also change how many rows `XLRowsCollection.ShiftRowsDown`
renumbers and confound the comparison.

Baseline, 1,000 inserts into a sheet with 1,000 live ranges and 1,000 formula cells:

| probe | ms | share |
|---|---:|---:|
| empty sheet — no ranges, no formulas | 671 | 14% |
| formula shift | 3,244 | 68% |
| range-shift pass | 359 | 8% |
| — of which, visiting ranges the shift cannot reach | 83 | 2% |
| **total** | **4,753** | |

The spec's premise — "every insert materializes and sorts **every live range in the worksheet**" — is
literally true and accounts for 8% of the workload its own acceptance criterion names. Formula
shifting, which the spec files third under workstream A, is 68%.

### A2: the index was declined, and the reason is the point

A2 prescribes replacing the repository scan with a query against `XLRangeIndex`. That cannot be done
as written: `_rangeRepository` is a `ConcurrentDictionary<XLRangeKey, WeakReference<XLRangeBase>>`
holding every live range object, and it has **no spatial index at all**. `XLRangeIndex` covers only
ranges explicitly added to an `XLRanges` collection — merged ranges, CF targets — which is a small
subset. A2 was therefore not "use the existing index" but "build a new one".

It was not built, because the scan is not the expense. Filtering the candidates by the same condition
`XLRangeShiftHelper` uses to decide whether to move anything, and sorting only the survivors, made
unreachable ranges **free**:

| probe | before | after A2 |
|---|---:|---:|
| empty sheet | 592 ms | 665 ms |
| 1,000 ranges above the insert (no-ops) | 812 ms | 585 ms |
| 1,000 ranges below the insert (all move) | 849 ms | 888 ms |

The no-op case is now indistinguishable from an empty sheet, which says enumerating a thousand weak
references costs nothing measurable. What cost was the virtual dispatch and address arithmetic per
range, plus sorting a collection that mostly did not need visiting. An index would remove only the
enumeration that is already free, while adding a second source of truth over a weak-reference store
that is mutated *during* the iteration it would serve — every affected range relocates as the pass
runs, rewriting repository keys through `RelocateRange`. The ranges that remain expensive are the
ones that genuinely move, and no index avoids those.

### A3: the parser rewrite, and what was actually slow about the regex

The regex was not the problem. For every matched address the old shifter called
`Workbook.Worksheet(name).Range(addr)`, materialising a live `XLRange` through the range repository —
once per reference, per formula, per shift. That is what made 1,000 inserts over 1,000 formula cells
allocate 5.4 GB. `ClosedXML.Parser` hands back each reference already decomposed into row/column
values and relative/absolute markers, so no address is re-parsed, and it knows which spans are string
literals, so the quote-parity scan goes too.

| probe | before | after |
|---|---:|---:|
| 1,000 formulas below (affected) | 3,915 ms / 5,386 MB | 1,012 ms / 1,201 MB |
| 1,000 formulas above (no-ops) | 1,827 ms / 2,523 MB | 674 ms / 602 MB |
| ranges + formulas below | 4,753 ms / 6,091 MB | 1,539 ms / 2,079 MB |

Formulas the parser rejects (external workbook references such as `'[file.xlsx]Sheet'!A1`) fall back
to the regex implementation, kept in `XLCellFormulaShifter.Legacy.cs`.

Equivalence is pinned by `FormulaShifterCorpus.tsv`: 2,072 (formula, shifted range, shift, host sheet)
combinations generated from the old implementation, driven by `FormulaShifterCorpusTests`. Regenerate
with `-- profile shiftercorpus`, which writes the corpus to stdout and reports divergences from the
legacy path on stderr.

### A3 found a correctness bug, so criterion 4 has a documented exception

Nine of the 2,072 cases diverge, all one bug: a deletion that removes the **tail** of a reference
computed the new bottom edge as `last + shift` with no clamp to the row above the deletion.

| formula | deletion | was | now |
|---|---|---|---|
| `3:5` | rows 5–7 | `3:2` | `3:4` |
| `4:8` | rows 5–9 | `4:3` | `4:4` |
| `B3:B7` | rows 5–9 | `B3:B2` | `B3:B4` |
| `A2:A8` | rows 5–9 | `A2:A3` | `A2:A4` |
| `B:D` | cols 4–6 | `B:A` | `B:C` |

The outputs it replaces were inverted ranges or silently dropped a surviving row — `A2:A8` losing
row 4 is the one most likely to have been hit in practice. The other 2,063 cases are byte-identical.

Two things the corpus caught that a smaller test set would not have: cross-sheet resolution (an
unqualified reference means the sheet its formula lives on, which only matches when the formula is on
the sheet *being shifted* — the first cut compared the wrong two names and stopped shifting every
cross-sheet reference, caught by the `ShiftingFormulas` golden file), and
`ReferenceArea.GetDisplayStringA1()` collapsing a degenerate area, which would have rewritten
`A5:A10` as `A5` when a deletion narrowed it to one cell.

### Criterion 1 is not met, and cannot be met by workstream A

> Inserting 1,000 rows one-by-one into a sheet with 1,000 live ranges: ≥ 10× faster than main.

That exact workload — ranges, no formulas — went **1,030 ms → 888 ms (1.16×)**. Its cost is 665 ms of
fixed per-insert work on an *empty* sheet plus ~220 ms of range moves that genuinely have to happen.
The criterion describes a workload whose cost the spec never located.

The workload the spec was *reaching* for — ranges **and** formulas, which is what a real sheet has —
went **4,753 ms → 1,539 ms (3.1×)**.

Getting to 10× needs the 665 ms fixed cost, which is out of this spec's scope and not yet diagnosed.
`ShiftRowsDown`'s per-insert LINQ chain and sort were the obvious suspect and were removed; it made no
measurable difference. With ~1,000 materialised rows the sort is ~10k int comparisons, while the same
pass makes ~1,000 `SetRowNumber` calls, each running `OnRangeAddressChanged` → `RelocateRange`: a
`ConcurrentDictionary` probe plus a walk of the registered range indexes. For `XLRow` that probe
always misses — rows live in `RowsCollection` and are never stored in the repository. Removing that
waste needs `RelocateRange` to know which range types can actually be stored or indexed, which is a
change to the repository contract. **That is the next thing to do, and it is where the remaining 43%
of this workload is.**

### Criterion 3 is mis-stated, like spec 11's criterion 2 was

> Styling a 100K-cell range allocates O(distinct styles), not O(cells).

It allocates O(cells) — 324 KB, 3,236 KB, 32,264 KB for 10K, 100K, 1M — and no implementation can make
it otherwise: styling N cells writes N entries into the style slice and the slice must grow to hold
them. `BulkStyleBenchmarks` measures three sizes precisely because one size cannot distinguish
"allocates per cell" from "allocates per style".

What matters is the constant: **~33 bytes/cell**, all slice storage, against ~234 bytes/cell before
spec 11's Task 4. Both linear; only one mostly waste. A future revision should state the criterion as
a constant, not a complexity class.

### Workstream B was already done

B1 landed as spec 11's Task 4 (#185), which rewrote `XLStylizedBase.ModifyStyle`/`SetStyle` onto the
slice-walking path — the conflict map at the top of `docs/specs/README.md` had already flagged that
05 must rebase onto 11. Only B2, the benchmark, was outstanding.

### C1 not done

The RBush-vs-QuadTree prototype was not built. Nothing measured in this spec touches `XLRangeIndex`:
the range-shift pass reads the repository, not the index, and after A2 its remaining cost is real
range moves. A consolidation is still defensible as maintenance — two spatial structures, one of them
391 lines — but it is not a performance change, and this spec produced no evidence either way.

### Status by task

| # | Task | Outcome |
|---|------|---------|
| A1 | Counting test; fix to single batch notify | Test added. No fix needed — batch notify already correct; the spec guessed right that it "may partially work already". |
| A2 | Index-driven `NotifyRangeShifted*` | Done as a filter, not an index. Index declined on evidence (above). |
| A3 | Slice-level formula shift | Done, on `ClosedXML.Parser`. 3.9× on the isolated formula workload. Fixed a reference-shifting bug. |
| A4 | *(not in the original plan)* fixed per-insert cost | Sort removed from `ShiftRowsDown`; no measurable effect. Real cause identified, not fixed. |
| B1 | Bulk style path | Already shipped in #185. |
| B2 | `BulkStyleBenchmarks` | Done; disproves criterion 3 as worded. |
| C1 | RBush-vs-QuadTree | Not done. |
| C2 | `FindUsedColumn` / `XLRanges.Ranges` de-LINQ | Done. |
