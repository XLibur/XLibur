# Spec 05 — Structural-Edit & Bulk-Style Scalability (rows insert, range shift, style propagation)

**Area:** Architecture + Performance
**Effort:** L (2–3 weeks; can split into 3 independent PRs)
**Dependencies:** None.
**Status:** Proposed

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
