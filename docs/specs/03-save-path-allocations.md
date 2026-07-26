# Spec 03 — Save-Path Allocation Reduction (543 MB → target ≤ 350 MB)

**Area:** Performance (write time + memory)
**Effort:** M (1–2 weeks, mostly independent small tasks)
**Dependencies:** Coordinate with Spec 01 (both touch `SheetDataWriter`); land this first or rebase.
**Status:** In progress — see [Results](#results-2026-07-25) below.

## Summary

The 50K-row formatted save (`CreateFormattedAndSave`) allocates ~543 MB and has barely moved while wall time halved — allocation is the remaining lever on the write path. The cell loop itself is already tuned (single slice pass, reusable cell-ref buffer, style memo); the remaining costs are number→string formatting, per-cell inherited-style resolution, style-key hashing, and a handful of known small items.

## Current state (verify line numbers before starting)

- `XLibur/Extensions/XmlWriterExtensions.cs` — `WriteNumberValue` formats doubles via `double.ToInvariantString()`, one string per numeric/date cell (~100K+ per benchmark save).
- `XLibur/Excel/XLWorksheet.cs` (~line 1671) — `GetStyleValue(point)` falls through to `GetInheritedStyleValue` for every cell without an explicit style; runs **before** `SheetDataWriter.ResolveCellStyleId`'s last-value memo can help.
- `XLibur/Excel/IO/SheetDataWriter.cs` — `CollectTableTotalCells` allocates a `HashSet<Point>` and probes it per cell whenever any table has a totals row.
- `XLibur/Excel/Style/XLStyleKey.cs` — `StyleKey_GetHashCode` benchmark: 40.7 ms/100K vs 12.0 ms for its worst component (`BorderKey`), 4.6/4.2/1.2 for Fill/Font/Color. The composite hash does redundant work.
- `XLibur/Excel/Cells/SharedStringTable.cs` — no `EnsureCapacity`; 50K unique strings rehash the dictionary repeatedly during save/populate.
- Benchmarks: `XLibur.Benchmarks/XLiburWorkbookBenchmarks.cs` (`CreateAndSave`, `CreateFormattedAndSave`), `StyleKeyHashCodeBenchmarks.cs`, `AllocationBenchmarks.cs`.

## Work plan (each task = one small PR, benchmark before/after in description)

| # | Task | Detail | Size |
|---|------|--------|------|
| 1 | Span-based number write | In `WriteNumberValue`: `Span<char> buf = stackalloc char[32]; value.TryFormat(buf, out int len, "G15", CultureInfo.InvariantCulture)` then `WriteRaw(char[], ...)` via a reusable `char[]` (XmlWriter.WriteRaw has no span overload — reuse a cached buffer). **The format specifier is `G15`, not `"R"`/G17** — see the note below. **Must produce byte-identical output** to `ToInvariantString` for round-trip stability — add a property-based test comparing both formatters over random doubles, ints, dates. | S |
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

- Task 1 formatting parity — **resolved while implementing Spec 02, do not re-derive**:
  `ObjectExtensions.ToInvariantString` formats `double` with **`"G15"`** and `float` with `"G7"`,
  annotated "Specify precision explicitly for backward compatibility". So the current output is
  *not* shortest-round-trip: values wider than 15 significant digits already lose precision on
  save (e.g. `1234567890.123456` is written as `1234567890.12346` and reads back as
  `1234567890.1234601`). Matching `"R"`/G17 would change the bytes of essentially every workbook
  and look like a regression in diffs. Use `G15`. The property test is still mandatory — it is
  what pins this behaviour down.
- Task 2 can add complexity for marginal gain if the profile doesn't confirm — profile first, implement second.

## References

- Memory notes: prior allocation plan identified items 1, 4, 5 plus the (already-landed) cell-loop optimizations.
- Perf survey: 543 MB headline, StyleKey hash numbers, repository WeakReference accumulation.

## Results (2026-07-25)

Measured with `dotnet run -c Release --project XLibur.Benchmarks --framework net10.0 -- profile alloc`,
which reports GC-exact allocated bytes for the same two workloads as
`XLiburWorkbookBenchmarks`, split into a create phase and a save phase. Three runs per side;
allocation figures are stable to ±0.5 MB, elapsed is the min–max of three runs.

| Scenario | Create | Save | Total | Elapsed |
|---|---|---|---|---|
| `CreateAndSave` — before | 58.3 MB | 72.2 MB | 130.5 MB | 322–368 ms |
| `CreateAndSave` — after | 58.3 MB | **34.9 MB** | **93.2 MB** (−28.6%) | 276–329 ms |
| `CreateFormattedAndSave` — before | 305.6 MB | 237.1 MB | 542.7 MB | 1187–1290 ms |
| `CreateFormattedAndSave` — after | 305.5 MB | **117.9 MB** (−50.3%) | **423.3 MB** (−22.0%) | 1051–1167 ms |

`StyleKey_GetHashCode`: **6,952 µs → 320 µs per 100K** (−95%). The absolute numbers differ from the
40.7 ms in the spec body because that figure came from different hardware; the ratio is what
carries over, and it clears the ≤15/40.7 (−63%) bar by a wide margin.

Saved output is byte-identical to `main`: unzipping both packages and diffing shows every part
under `xl/` and `docProps/` matches exactly. Only the package-level `_rels/.rels` and the
`core-properties` part name differ, and those carry GUIDs that `System.IO.Packaging` regenerates
on every save regardless of this change.

### What landed

| # | Task | Outcome |
|---|------|---------|
| 1 | Span-based number write | Already landed for `double`; hardened (the `TryFormat` result was ignored) and extended to `int`/`uint`, which were still going through `XmlWriter.WriteValue` and allocating a string per cell. Parity test added: `XmlWriterExtensionsTests` compares the writer against `ToInvariantString` over 60K random doubles, serial dates and ints. |
| 2 | Inherited-style memo | **Not implemented — profile disconfirmed the premise.** Redirecting the sheet-data writer at `Stream.Null` drops its allocation from 96.0 MB to 0.0 MB, i.e. the cell loop (including `GetStyleValue` → `GetInheritedStyleValue`) allocates nothing at all; every byte is the package stream. A memo here would add state for no allocation win. |
| 3 | StyleKey hash caching | Implemented. Component hashes are memoised in the `init` accessors of `XLStyleKey` and folded together in `GetHashCode`; `Equals` now rejects on those ints before touching the components. The combining formula is unchanged, so hash *values* are identical to before and style ordering in the output is untouched. |
| 4 | Table-totals guard | Already landed — `CollectTableTotalCells` returns `null` when there are no tables and allocates only when some table has a totals row; the per-cell probe is behind that null check. |
| 5 | SST `EnsureCapacity` | Already landed on the load path (`XLWorkbook_Load`). No save-side bulk-populate path exists: the SST is filled incrementally by `IncreaseRef` as cells are assigned, and `SharedStringTableWriter` streams straight out of it. |
| 6 | Repository dead-entry sweep | Implemented. `XLRepositoryBase` uses `WeakReference<T>`, revives collected entries with an atomic `TryUpdate` instead of remove-then-add, and prunes the whole map once 512 lookups have hit collected entries. |
| — | Full-sheet `GetCells()` scans (not in the original plan) | Four save-path helpers materialised an `XLCell` wrapper for **every used cell** just to read one property. This was the bulk of the save-phase allocation. `CollectWorkbookStyles` now reads `GetStyleValue` off the slices; `EnsureDynamicArrayMetadata` and `CalculationChainPartWriter` walk the formula slice; comment/cell-image detection walks the misc slice. |

### Remaining hotspots (Task 7)

Save phase, `CreateFormattedAndSave`, 117.9 MB total:

- **~96 MB (81%) — `System.IO.Packaging` part buffering.** `worksheetPart.GetStream(FileMode.Create)`
  is backed by a `ZipArchive` in update mode, so the whole ~25 MB `sheet1.xml` is buffered in a
  `MemoryStream` that doubles as it grows. Writing the identical XML to `Stream.Null` allocates
  0.0 MB. As the spec anticipated: noted, not chased — it needs a different packaging strategy
  (Spec 01 Phase 3).
- **~13 MB — workbook-level parts** (styles + SST serialisation).
- **~9 MB — package dispose** (compression).

Create phase, 305.5 MB, is untouched by this spec and is where the next 30% lives:

- **~186 MB — cell population.** `ws.Cell(r, c).Value = x` costs ~372 bytes/cell, dominated by the
  `XLAddress`/`XLCell` wrappers minted per access. `XLWorksheet.SetCellValue(row, column, value)`
  already bypasses them internally.
- **~120 MB — styling.** Each `ws.Cell(r, c).Style` mints an `XLCell`, an `XLStyle` and its
  sub-wrappers; the sub-wrapper cache on `XLStyle` cannot help because the `XLStyle` itself is new
  every time. Fixing this means reworking how `IXLStyle` handles are vended, which conflicts with
  the "no public API changes" constraint here.

Consequence for acceptance criterion 1: the ≥30% target (≤380 MB) is **not reachable from the
save path alone** — the save path is now 117.9 MB in total, of which 96 MB is packaging. The
remaining 43 MB gap has to come from the create-phase items above, which belong in their own spec.
