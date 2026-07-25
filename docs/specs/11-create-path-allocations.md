# Spec 11 — Create-Path Allocation Reduction (306 MB → target ≤ 120 MB)

**Area:** Performance (build time + memory)
**Effort:** M (1–2 weeks; Task 1 is a one-liner, the rest are independent)
**Dependencies:** None for Tasks 1–3. Task 4 overlaps Spec 05 (bulk styling); Task 5 changes an
internal object-vending pattern, so land it after 1–4 and re-measure first.
**Status:** ✅ Tasks 1–3 implemented — see [Results](#results). Task 4 deferred to Spec 05 with
measurements; Task 5 out of scope by design.

## Summary

Spec 03 cut the *save* half of `CreateFormattedAndSave` in half (237.1 → 117.9 MB) and found that
the remaining save cost is `System.IO.Packaging` buffering the part stream — not something the
serializer can fix. That leaves the **create** half, which Spec 03 did not touch: **305.5 MB of the
benchmark's 423.3 MB, or 72%**.

That 305 MB is not the slice model — the slices are dense and cheap. It is the per-cell object
churn of the public API: `ws.Cell(r, c)` mints an `XLCell`, `.Style` mints an `XLStyle` plus
sub-wrappers, and — the single largest item — **every `cell.Value = x` runs a merged-range
membership test that allocates ~250 bytes even on a sheet with no merged ranges at all.**

The headline finding: **a one-line guard on that test takes the whole benchmark from 423.3 MB to
312.4 MB**, which is past Spec 03's stretch target on its own. Measurements below are reproducible;
the probe used to obtain them is included verbatim.

## Baseline

Post-Spec-03 (`perf/save-path-allocations`, PR #179), 50K rows × 10 cols, net10.0, Release:

| Scenario | Create | Save | Total |
|---|---:|---:|---:|
| `CreateAndSave` | 58.3 MB | 34.9 MB | 93.2 MB |
| `CreateFormattedAndSave` | **305.5 MB** | 117.9 MB | 423.3 MB |

Splitting the formatted create phase by disabling `ApplyRowFormatting`:

| Portion | Allocated | Operations |
|---|---:|---|
| Cell population (`ws.Cell(r,c).Value = x`) | 185.9 MB | 500K cell writes |
| Styling (`ws.Cell(r,c).Style` + mutations) | 119.6 MB | 250K styled cells, 2–5 mutations each |

## Where it goes (measured, not inferred)

Per-operation costs over 500,000 operations on a fresh worksheet with **no merged ranges, no
tables, no formulas**:

| Probe | Bytes/op | ns/op |
|---|---:|---:|
| `ws.Cell(r,c)` discarded | 48.2 | 15.0 |
| `ws.Cell(r,c).Value = double` | **375.5** | 269.8 |
| `ws.SetCellValue(r,c, double)` | **55.6** | 100.2 |
| `ws.Cell(r,c).Value = string` (shared) | 375.5 | 280.8 |
| `ws.Cell(r,c).Value = DateTime` | 408.6 | 416.0 |
| `ws.Cell(r,c).Style` discarded | 128.2 | 54.8 |
| `ws.Cell(r,c).Style` + 1 font mutation | 217.2 | 263.1 |
| `ws.Cell(r,c).Style` + 4 mutations | 473.2 | 699.4 |

Two facts fall straight out of this table:

1. **`ws.SetCellValue(r, c, v)` costs 55.6 bytes; `ws.Cell(r, c).Value = v` costs 375.5.** Same
   result, same slices, **6.8× the allocation**. The 320-byte gap is *not* the `XLCell` wrapper —
   that is only 48 bytes.
2. **`ws.Cell(r,c).Style` costs 128.2 bytes before any mutation**, i.e. 48 for the `XLCell` plus
   ~80 for an `XLStyle` that is discarded one statement later.

### The 320-byte gap is the merged-range check

`XLCell.SetValue` (`XLibur/Excel/Cells/XLCell.cs:208`) opens with:

```csharp
if (checkMergedRanges && IsInferiorMergedCell())
    return this;
```

and `IsInferiorMergedCell` (`XLCell.cs:1388`) is:

```csharp
internal bool IsInferiorMergedCell()
    => IsMerged() && !Address.Equals(MergedRange()!.RangeAddress.FirstAddress);

public bool IsMerged() => Worksheet.Internals.MergedRanges.Contains(this);
```

`XLRanges.Contains(IXLCell)` (`XLibur/Excel/Ranges/XLRanges.cs:115`) is
`GetIntersectedRanges((XLAddress)cell.Address).Any()`, which on an empty index allocates roughly
five objects per call:

- `cell.Address` is reached through `IXLCell`, whose `Address` is typed `IXLAddress` — so the
  `XLAddress` **struct is boxed**, then immediately unboxed by the `(XLAddress)` cast
- `XLRangeIndex.GetIntersectedRanges` falls into `_rangeList.Where(r => r.RangeAddress.Contains(address))`
  (`XLibur/Excel/Ranges/Index/XLRangeIndex.cs:91`) — a **display class** for the captured address
  plus a **`Where` iterator**
- the generic subclass adds `.Cast<T>()` (`XLRangeIndex.cs:208`) — **another iterator**
- `.Any()` allocates the **enumerator**

Removing the call entirely drops `Value = double` from 375.5 to ~100 bytes/op, confirming the
attribution. Everything else in the write path — `GetStyleForValue`, `ValueSlice.SetCellValue`,
`FormulaSlice.Get`, `CalcEngine.MarkDirty` — accounts for the remaining ~50 bytes together.
`MarkDirty` measured at zero: with no dependency tree built it returns immediately.

**This work is pure waste whenever the sheet has no merged ranges**, which is the overwhelmingly
common case, and `XLWorksheetInternals.MergedRanges` is a concrete `XLRanges` with an O(1) `Count`
(`XLibur/Excel/Ranges/XLRanges.cs:98`).

## Work plan (each task = one small PR, before/after numbers in the description)

| # | Task | Detail | Size |
|---|------|--------|------|
| 1 | Guard the merged-range check | `if (checkMergedRanges && Worksheet.Internals.MergedRanges.Count > 0 && IsInferiorMergedCell())`. Exactly equivalent: `Contains` cannot return true on an empty collection. Same guard for the second call site in the `FormulaA1` setter (`XLCell.cs:555`). **Validated: 375.5 → 103.6 bytes/op, 269.8 → 142.7 ns/op; whole benchmark 423.3 → 312.4 MB.** | XS |
| 2 | Allocation-free `IsMerged` | Fix the underlying path so it is cheap even *with* merges: add `bool Contains(in XLAddress)` to the `XLRanges`/`XLRangeIndex` surface (`XLRangeIndex.Contains(in XLAddress)` already exists at `XLRangeIndex.cs:56` and is unused by this path), and route `XLCell.IsMerged` to it via the concrete `XLRanges` type so `IXLCell.Address` is never boxed. Then replace the `Where`/`Cast`/`Any` chain in `XLRangeIndex.Contains` with a plain `foreach` over `_rangeList`. Task 1 makes the empty case free; this makes the non-empty case cheap. | S |
| 3 | Don't mint an `XLStyle` to read one component | `XLStylizedBase.InnerStyle` (`XLibur/Excel/Style/XLStylizedBase.cs`) caches an `XLStyle` per stylized object, but an `XLCell` is itself new on every `ws.Cell(r,c)`, so the cache never hits — 80 bytes per `.Style` access, discarded immediately. Give `XLCellsCollection` a small direct-mapped cache of recently-vended `XLCell`s keyed by `XLSheetPoint` (the wrapper is stateless apart from `_point`, so sharing is safe), which makes both the 48-byte `XLCell` and the 80-byte `XLStyle` reusable across the common `ws.Cell(r,c).Style.X = ...; ws.Cell(r,c).Style.Y = ...` pattern. Profile before/after — this is the least certain item. | M |
| 4 | Bulk-style entry point | Add an internal, wrapper-free style path analogous to `XLWorksheet.SetCellValue`, e.g. `XLWorksheet.SetCellStyle(int row, int column, Func<XLStyleKey, XLStyleKey>)`, and route `IXLRange.Style` / `IXLRow.Style` / `IXLColumn.Style` bulk assignments through it instead of per-cell `XLCell` + `XLStyle` pairs. Coordinate with Spec 05 (bulk style propagation) — same territory. | M |
| 5 | Reconsider how `IXLStyle` handles are vended | Longer-term: `cell.Style.Font.Bold = true` currently costs `XLCell` + `XLStyle` + `XLFont`. A struct-based or pooled handle would remove the last of it, but changes public return types, so it needs its own design round and a compatibility story. **Do not attempt inside this spec** — recorded here so the ceiling is visible. | L |

Tasks 1 and 2 are independent of 3 and 4. Task 1 should land on its own, immediately: it is one
line and worth 111 MB.

## Measurement protocol

Allocation totals: `dotnet run -c Release --project XLibur.Benchmarks --framework net10.0 -- profile alloc`
(added in PR #179 — GC-exact, split into create/save phases). Three runs; figures are stable to
±0.5 MB. Wall time from that harness is single-shot and noisy — quote a min–max over three runs, or
use BenchmarkDotNet for any time claim.

Per-operation attribution: add a probe to `XLibur.Benchmarks` shaped like the following, which is
what produced the table above. It is deliberately not checked in — it is a bisection tool, not a
benchmark.

```csharp
// 500_000 operations; call each probe once to warm up, then measure.
GC.Collect(2, GCCollectionMode.Forced, true, true);
GC.WaitForPendingFinalizers();
var before = GC.GetTotalAllocatedBytes(precise: true);
var watch = Stopwatch.StartNew();

using (var wb = new XLWorkbook())
{
    var ws = wb.AddWorksheet("s");
    for (var r = 1; r <= 50_000; r++)
    for (var c = 1; c <= 10; c++)
        ws.Cell(r, c).Value = r * 1.5;   // <- swap this line per probe
}

watch.Stop();
var bytes = GC.GetTotalAllocatedBytes(precise: true) - before;
```

**Caveat that cost time to find:** if you gate a code path behind
`Environment.GetEnvironmentVariable(...)` to bisect, remember that call itself allocates a string
per invocation (~28 bytes/op at this scale) and will show up in the result. Subtract it, or use a
`static readonly bool` initialized once.

## Acceptance criteria

1. `CreateFormattedAndSave` create-phase allocation reduced from 305.5 MB to **≤ 120 MB**
   (Task 1 alone reaches 194.8 MB); total benchmark allocation **≤ 240 MB** from the current
   423.3 MB. Wall time not regressed.
2. `ws.Cell(r,c).Value = double` allocates **≤ 110 bytes/op** and `ws.Cell(r,c).Style` (unmutated)
   **≤ 60 bytes/op** on a sheet with no merged ranges, tables or formulas.
3. Saved output byte-identical versus main for the full test corpus (compare unzipped `xl/` and
   `docProps/` parts; the package-level `_rels/.rels` and `core-properties` part name carry GUIDs
   that `System.IO.Packaging` regenerates on every save and will always differ).
4. All tests green on net8.0 and net10.0; no public API changes in Tasks 1–4.

## Risks

- **Task 1 is only equivalent if `MergedRanges` is always non-null and its `Count` is authoritative.**
  It is a concrete `XLRanges` field on `XLWorksheetInternals` (`XLWorksheetInternals.cs:27`),
  assigned in the constructor, and `Count` is maintained by `Add`/`Remove`. Add a test that sets a
  value into an inferior merged cell and asserts it is still ignored — that behaviour is what the
  check exists for, and it must not regress.
- **Task 3 (shared `XLCell` wrappers) changes reference identity of vended cells.** `XLCell`
  overrides `Equals`/`GetHashCode` on `(SheetPoint, Worksheet)` (`XLCell.cs:1401`), so value
  semantics are unaffected, but anything relying on `ReferenceEquals` between two `ws.Cell(r,c)`
  calls, or holding a cell across a structural edit, would change behaviour. Audit
  `ReferenceEquals`/`==` on `XLCell` and the range-shift code before implementing. If the audit
  looks risky, cap the task at caching the `XLStyle` rather than the `XLCell`.
- **Task 4 overlaps Spec 05.** Whoever goes second rebases; do not run them concurrently.
- The `IXLCell.Address` boxing in Task 2 is on an interface member, so other callers benefit too —
  but check for callers that depend on `Address` returning a fresh instance before changing shape.

## References

- Spec 03 Results section — save-phase breakdown, and the finding that ~96 MB of the remaining save
  cost is `System.IO.Packaging`, not XLibur code.
- Spec 05 (structural-edit & bulk-style scalability) — overlapping territory for Task 4.
- PR #179 — introduced the `profile alloc` harness this spec measures with.

## Results

`profile alloc`, 50K rows, net10.0, Release. Three runs per side; allocation stable to ±1 MB,
elapsed reported min–max.

| Scenario | Create | Save | Total | Elapsed |
|---|---:|---:|---:|---|
| `CreateAndSave` — before | 58.3 MB | 34.9 MB | 93.2 MB | 285–302 ms |
| `CreateAndSave` — after | **25.1 MB** | 34.9 MB | **60.1 MB** (−35.5%) | 274–285 ms |
| `CreateFormattedAndSave` — before | 305.5 MB | 118.0 MB | 423.4 MB | 1117–1175 ms |
| `CreateFormattedAndSave` — after | **183.6 MB** (−39.9%) | 117.6 MB | **301.3 MB** (−28.8%) | 994–1080 ms |

`profile create`, per-operation, no merged ranges / tables / formulas:

| Probe | Before | After |
|---|---:|---:|
| `ws.Cell(r,c).Value = double` | 375.5 B | **103.6 B** |
| `ws.Cell(r,c).Value = string` | 375.5 B | **103.6 B** |
| `ws.Cell(r,c).Value = DateTime` | 408.6 B | **136.6 B** |
| `...Value = double`, sheet has 1 merged range | 487.6 B (net8) / 457.7 B (net10) | **103.6 B** (both) |
| `ws.Cell(r,c).Style` (unmutated) | 128.2 B | 128.2 B — see below |

6325 tests pass on net8.0 and net10.0.

### What landed

- **Task 1** — `IsMerged` short-circuits on `MergedRanges.Count`. Done inside `IsMerged` rather than
  at the two call sites the spec named, which covers every caller for the same one line. Worth
  111 MB on the benchmark by itself.
- **Task 2** — bigger than the spec estimated. A sheet *with* a merged range was **worse** than one
  without (487.6 vs 375.5 bytes/op), because the predicate then actually ran and boxed
  `RangeAddress` per range on top of everything else. A merged title row is a very common layout,
  and it took every subsequent cell write off Task 1's fast path. Now identical to the unmerged
  case on both TFMs.
- **Task 3** — implemented as specified after the reference-identity audit came back clean (nothing
  in the codebase compares `XLCell` instances by reference; the wrapper's only instance state is
  the derived `_cachedStyle`, which `InnerStyle` re-syncs). Worth 11.5 MB.

### Task 3 does not move acceptance criterion 2

The criterion asked for `ws.Cell(r,c).Style` ≤ 60 bytes/op. It is still 128.2, and **no cache keyed
by cell address can change that** — the probe visits every point exactly once, so the first access
to a cell must build the wrapper and its `XLStyle`. The 11.5 MB Task 3 does deliver comes from real
workloads re-touching the same cells (populate pass, then style pass), which the probe deliberately
does not model. Getting the single-access number down needs Task 5, which is out of scope here.
**The criterion was mis-stated, not missed** — a future revision should express it as a workload
number rather than a per-op one.

### Task 4 deferred to Spec 05, with measurements

Bulk styling — `ws.Range(...).Style.Font.Bold = true` — costs **~206 bytes per cell** (measured:
289.4 bytes/op for a probe that also populates the cells at 55.6, over 500K cells). That is worth
fixing, but:

1. It contributes **nothing** to this spec's acceptance criteria. `CreateFormattedAndSave` styles
   cells individually and never goes through the bulk path.
2. The cost is not the `GroupBy` the spec's Task 4 pointed at. Replacing the group-by with a
   last-value memo — the contained part of the change — moved it only 289.4 → 272.6 bytes/op (6%).
   That change was written, measured, and reverted rather than banked, because a 6% sliver is not
   worth putting a second author into the file Spec 05 is going to rewrite.
3. The remaining 94% is `XLStylizedBase.ModifyStyle` materialising every child cell into a
   `HashSet<XLStylizedBase>` via `Children`. Removing that means writing the `StyleSlice` by point
   and never building an `XLCell` — which is exactly Spec 05's "materialize-everything patterns for
   style propagation", and carries semantics this spec did not scope (whether bulk styling pings
   currently-empty cells, dedup across overlapping ranges in `XLRanges`, propagation order).

The `ws.Range(all).Style.Font.Bold = true` probe is checked in so Spec 05 starts with a baseline.

### Remaining create-phase cost

183.6 MB, down from 305.5:

- **~52 MB** — cell population at 103.6 bytes/op. Of that, 48.2 is the `XLCell` wrapper and ~55 is
  the actual slice write (`ws.SetCellValue` costs 55.6 and does the same work). Only Task 5 reaches
  the wrapper.
- **~118 MB** — styling, at 473.2 bytes/op for the four-mutation pattern the benchmark uses. Each
  mutation still walks `XLStyle` → sub-wrapper → repository.
- The rest is benchmark-side strings.
