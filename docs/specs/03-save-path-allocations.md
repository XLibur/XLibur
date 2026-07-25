# Spec 03 — Save-Path Allocation Reduction (543 MB → target ≤ 350 MB)

**Area:** Performance (write time + memory)
**Effort:** M (1–2 weeks, mostly independent small tasks)
**Dependencies:** Coordinate with Spec 01 (both touch `SheetDataWriter`); land this first or rebase.
**Status:** Proposed

## Summary

The 50K-row formatted save (`CreateFormattedAndSave`) allocates ~543 MB and has barely moved while wall time halved — allocation is the remaining lever on the write path. The cell loop itself is already tuned (single slice pass, reusable cell-ref buffer, style memo); the remaining costs are number→string formatting, per-cell inherited-style resolution, style-key hashing, and a handful of known small items.

## Current state (verify line numbers before starting)

- `XLibur/Extensions/XmlWriterExtensions.cs` — `WriteNumberValue` formats doubles via `double.ToInvariantString()`, one string per numeric/date cell (~100K+ per benchmark save).
- `XLibur/Excel/XLWorksheet.cs` (~line 1671) — `GetStyleValue(point)` falls through to `GetInheritedStyleValue` for every cell without an explicit style; runs **before** `SheetDataWriter.ResolveCellStyleId`'s last-value memo can help.
- `XLibur/Excel/IO/SheetDataWriter.cs` — `CollectTableTotalCells` allocates a `HashSet<XLSheetPoint>` and probes it per cell whenever any table has a totals row.
- `XLibur/Excel/Style/XLStyleKey.cs` — `StyleKey_GetHashCode` benchmark: 40.7 ms/100K vs 12.0 ms for its worst component (`BorderKey`), 4.6/4.2/1.2 for Fill/Font/Color. The composite hash does redundant work.
- `XLibur/Excel/Cells/SharedStringTable.cs` — no `EnsureCapacity`; 50K unique strings rehash the dictionary repeatedly during save/populate.
- Benchmarks: `XLibur.Benchmarks/XLiburWorkbookBenchmarks.cs` (`CreateAndSave`, `CreateFormattedAndSave`), `StyleKeyHashCodeBenchmarks.cs`, `AllocationBenchmarks.cs`.

## Work plan (each task = one small PR, benchmark before/after in description)

| # | Task | Detail | Size |
|---|------|--------|------|
| 1 | Span-based number write | In `WriteNumberValue`: `Span<char> buf = stackalloc char[24]; value.TryFormat(buf, out int len, "R"/G17-equivalent, CultureInfo.InvariantCulture)` then `WriteRaw(char[], ...)` via a reusable `char[]` (XmlWriter.WriteRaw has no span overload — reuse a cached buffer). **Must produce byte-identical output** to `ToInvariantString` for round-trip stability — add a property-based test comparing both formatters over random doubles, ints, dates. | S |
| 2 | Inherited-style memo | In the save loop, memoize `GetStyleValue` results per (row-style, column-style) pair — for a typical sheet most cells inherit the same worksheet/column style. Cache last row's resolved inherited value; invalidate on row change. Profile first to confirm this is the hot allocation (dotMemory `profile` harness). | M |
| 3 | StyleKey hash caching | `XLStyleValue` is immutable — compute the full hash once in the constructor and store it (`_cachedHashCode`), or fix `XLStyleKey.GetHashCode` to combine pre-computed component hashes instead of re-hashing all fields. Verify with `StyleKeyHashCodeBenchmarks`. | S |
| 4 | Table-totals guard | Skip `CollectTableTotalCells` + per-cell `Contains` entirely when `worksheet.Tables.All(t => !t.ShowTotalsRow)` (or count == 0). | XS |
| 5 | SST `EnsureCapacity` | Add `internal void EnsureCapacity(int)` to `SharedStringTable`; call from bulk populate paths and from the SST load path (Spec 02 Task B uses it too). | XS |
| 6 | Repository dead-entry sweep | `XLRepositoryBase` (`XLibur/Excel/Caching/`) uses non-generic `WeakReference` per entry and never prunes dead ones. Switch to `WeakReference<T>` and add an opportunistic prune (e.g., every N misses). Long-lived processes creating many workbooks currently accumulate shells. | S |
| 7 | Profile-verify remainder | After 1–6, take a fresh dotMemory snapshot of `CreateFormattedAndSave`; file follow-up issues for anything ≥ 5% of remaining allocations (likely `XmlWriter` internals and `System.IO.Packaging` — note but don't chase; that's Spec 01 Phase 3 territory). | S |

## Measurement protocol

Same as Spec 02: BenchmarkDotNet filter `'*XLiburWorkbookBenchmarks*'` + dotMemory `profile` harness; before/after table in every PR description.

## Acceptance criteria

1. `CreateFormattedAndSave` allocated bytes reduced ≥ 30% total across the task series (543 MB → ≤ 380 MB; stretch ≤ 350 MB); wall time not regressed.
2. Saved output byte-identical (or semantically identical — same values after reload) versus main for the full test corpus; number formatting round-trip test added (task 1).
3. `StyleKey_GetHashCode` benchmark ≤ 15 ms/100K (from 40.7).
4. All tests green; no public API changes.

## Risks

- Task 1 formatting parity: `double.ToInvariantString` may use shortest-round-trip ("R"-like) semantics — match exactly; Excel files store shortest-round-trip doubles. The property test is mandatory, not optional.
- Task 2 can add complexity for marginal gain if the profile doesn't confirm — profile first, implement second.

## References

- Memory notes: prior allocation plan identified items 1, 4, 5 plus the (already-landed) cell-loop optimizations.
- Perf survey: 543 MB headline, StyleKey hash numbers, repository WeakReference accumulation.
