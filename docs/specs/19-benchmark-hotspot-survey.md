# Spec 19 — Benchmark Hotspot Survey (August 2026): five areas to investigate

**Area:** Performance (read + write time, memory)
**Effort:** L overall; each of the five areas is independently sized and independently ownable
**Dependencies:** None between the five areas. Area 3 overlaps Spec 01/03 territory; Area 4 overlaps Spec 18 task 5; Area 5 overlaps Spec 04. Those relationships are stated per area.
**Status:** Area 1 tasks 1.1–1.3 and Area 2 task 2.3 done (see [Area 1 results](#area-1-results),
[Area 2 results](#area-2-results)); areas 3–5 proposed.
This started as a survey rather than an implementation plan. Every area opens with what was measured
and closes with what has *not* been established, so an implementing agent knows which sentences are
evidence and which are hypotheses — Area 1's own ranking was wrong until task 1.1 measured it.

## Why this spec exists

The whole benchmark suite was re-run on 2026-08-07 against `docs/changelog-copyedit` (parent
`dd64819b`, library identical to `main` at `8c207377`). The specs that produced the current numbers —
02 (load allocations), 03 (save allocations), 05 (structural edits), 11 (create allocations), 18
(template round trip) — each optimised the thing they set out to optimise and then recorded what they
left behind. This spec reads the *whole board* instead, and the ranking that falls out is not the one
those specs would predict.

The headline: **on the read path, the public cell-access APIs now cost more than parsing the file
does.** Iterating 3.75 M cells through `CellsUsed()` costs 2.39 s and 1,020 MB *on top of* a load that
itself costs 3.72 s and 335 MB. The struct enumerator added alongside it walks the same cells for
0.36 s and **zero** additional bytes. That gap is the largest single number in the suite and it is not
a parsing problem, a packaging problem, or a styling problem — the three places the existing specs
point.

---

## Measured baselines

BenchmarkDotNet 0.15.8, net10.0 Release, `InProcessEmitToolchain`, default job unless noted.
Allocation figures from `[MemoryDiagnoser]`; the `profile` figures further down are
`GC.GetTotalAllocatedBytes(precise: true)` and are exact.

### Read path — `XLiburReadBenchmarks`, 250,000 × 15 = 3.75 M cells

Run with `--warmupCount 1 --iterationCount 3` (each operation is 4–6 s; the mutator applies to the
config's job rather than adding a second one — check the "Found N benchmarks" header if in doubt).

| Benchmark | Mean | Allocated | over `LoadWorkbook` |
|---|---:|---:|---|
| `LoadWorkbook` | 3.717 s | 334.54 MB | — |
| `LoadAndIterateEnumerateUsedCells` | 4.080 s | **334.54 MB** | +0.36 s / **+0.00 MB** |
| `LoadAndReadAllCells` | 5.262 s | 957.43 MB | +1.55 s / +622.9 MB |
| `LoadAndIterateCellsUsed` | 6.109 s | 1,354.38 MB | +2.39 s / +1,019.8 MB |

Per cell: load 0.99 µs / 93.5 B · `EnumerateUsedCells` 97 ns / 0 B · `Cell(r,c).GetValue<string>()`
413 ns / 166 B · `CellsUsed()` 638 ns / 272 B.

### Write path — `XLiburWorkbookBenchmarks`, 50,000 rows

| Benchmark | Mean | Allocated |
|---|---:|---:|
| `CreateAndSave` (50K × 3) | 244.4 ms · 255.5 ms | 60.52 MB |
| `CreateAndSaveFastestCompression` (new, see task 0) | **179.5 ms** | 60.59 MB |
| `CreateFormattedAndSave` (50K × 10, half the rows styled) | 1,005.5 ms · 1,097.3 ms | 322.0 MB · 325.1 MB |

Two figures are given where a benchmark was run twice, an hour apart, on identical code. **That
spread is the noise floor on this machine: 4.5% on `CreateAndSave`, 9% on `CreateFormattedAndSave`.**
Nothing below 10% is a result here without an A/B in one sitting.

`profile alloc` (exact, single iteration):

| Scenario | Create | Save | Total |
|---|---:|---:|---:|
| `CreateAndSave` | 25.1 MB | 34.9 MB | 60.0 MB |
| `CreateFormattedAndSave` | **204.2 MB** | 117.7 MB | 321.8 MB |

### Streaming writer — `StreamingWriteBenchmarks`, the same 50K × 3 data

| Benchmark | Mean | Allocated |
|---|---:|---:|
| `StreamingWrite` | 163.6 ms | 13.60 MB |
| `StreamingWriteInlineStrings` | 145.7 ms | 8.70 MB |
| `StreamingWriteFastestCompression` | **94.7 ms** | 17.01 MB |

### Template round trip — `TemplateRoundTripBenchmarks`

| Benchmark | Mean | Allocated |
|---|---:|---:|
| `OpenAndSaveRowHeavyUnchanged` (20,000 × 21) | 909.2 ms | 88.84 MB |
| `LoadRowHeavy` (same fixture) | 385.9 ms | 55.40 MB |
| `Open` (10 sheets, 20 names, 26 validations) | 4.47 ms | 1.67 MB |
| `OpenAndSaveUnchanged` | 8.38 ms | 3.17 MB |
| `RefreshLookupColumn` (1,000 values) | 9.99 ms | 3.83 MB |

All five reproduce Spec 18's post-task-1 figures within noise, with allocation identical to the byte.
Spec 18's results still stand.

### `profile create` — per operation over 500,000 operations, bytes exact

| Probe | Bytes/op | ns/op |
|---|---:|---:|
| `ws.Cell(r,c)` discarded | 48.1 | 19.8 |
| `ws.Cell(r,c).Value = double` | 103.5 | 148.3 |
| `ws.SetCellValue(r,c, double)` *(internal, no wrapper)* | 55.5 | 104.4 |
| `ws.Cell(r,c).Value = string` (shared) | 103.5 | 246.9 |
| **`ws.Cell(r,c).Value = DateTime`** | **136.5** | **487.7** |
| `ws.Cell(r,c).Style` discarded | 128.1 | 53.2 |
| `ws.Cell(r,c).Style` + 1 font mutation | 217.1 | 247.5 |
| `ws.Cell(r,c).Style` + 4 mutations | 473.1 | 603.6 |
| `ws.Range(all).Style.Bold` + populate | 88.6 | 246.7 |

### `profile template` — grid write, 20,000 × 21 = 420,000 cells

| Phase | Time | Allocated | per cell |
|---|---:|---:|---:|
| build the model (cells only) | 171.1 ms | 85.8 MB | 0.41 µs |
| save only | **438.1 ms** | 73.5 MB | **1.04 µs** |
| load the same file (`LoadRowHeavy`, BDN) | 385.9 ms | 55.4 MB | 0.92 µs |

### Supporting classes

`CellStylingBenchmarks` (100,000 rows) — `ValueOnly` 25.31 ms / 24.86 MB; `StyleFacadePerCell` 46.34 /
37.83; `StyleAssignedPerCell` 45.86 / 37.07; `TwoPropertiesPerCell` 56.67 / 43.17; `FacadeStyleOnly`
43.31 / 32.49; `FacadeChainOnly` 45.66 / 35.54; `StyleOnColumn` 20.30 / 16.25. Std-dev runs
2.0–6.3 ms, so the *time* column here separates only the coarse groupings; allocation is solid.

`SheetGeometryBenchmarks` (20,000 string cells) — `TallNarrow_Unique` 44.47 ms / 17.44 MB;
`TallNarrow_Repeated` 32.90 / 9.98; `ShortWide_Unique` 33.61 / 12.87; `ShortWide_Repeated` 17.44 /
5.52; write-only variants 10.24 / 7.45 and 8.67 / 6.15. Reproduces Spec 18 task 3; still explained,
still nothing to fix.

`FormulaEvaluationBenchmarks` (20,000 formulas) — `UniqueSameSheet` 16.41 ms; `SharedSameSheet` 13.60;
`SharedCrossSheet` 42.09. Allocation identical at 10.38 MB across all three. **See Area 5 — the third
row is confounded and cannot be read as a cross-sheet penalty.**

`StyleKeyHashCodeBenchmarks` (per 100,000) — `BorderKey` 2,572.9 µs; `FillKey` 872.9; `ColorKey`
282.0; **`StyleKey` 339.1**. The composite is now 7.6× *cheaper* than one of its own components,
because Spec 03 task 3 memoised the component hashes into `XLStyleKey`'s `init` accessors while a
freshly built `XLBorderKey` still hashes from scratch.

`AllocationBenchmarks` (per 1,000 calls) — `ToExcelFormat` 259.2 µs / 436,802 B (**437 B per call**);
`ShiftFormulaRows` 866.8 µs / 1,150,405 B (1,150 B per call); `IsEmptyConditionalFormats` 250.5 µs /
640,002 B (640 B per call); `EscapeSheetName` 96.6 µs / 45,601 B; `ToHex` 17.2 µs / 40,000 B;
`AddressToString` 18.7 µs / 35,200 B; `FixNewLines` 16.7 µs / 6,400 B; `AddressCreate` 31.2 µs / 0 B;
`CharCount` 2.6 µs / 0 B.

### Reproducing

```bash
dotnet build XLibur.Benchmarks/XLibur.Benchmarks.csproj -c Release -f net10.0

# Comment out [DotTraceDiagnoser] on XLiburWorkbookBenchmarks first, or the run takes hours.
XLibur.Benchmarks/bin/Release/net10.0/XLibur.Benchmarks.exe --filter "*XLiburWorkbookBenchmarks*"
XLibur.Benchmarks/bin/Release/net10.0/XLibur.Benchmarks.exe --filter "*XLiburReadBenchmarks*" --warmupCount 1 --iterationCount 3
XLibur.Benchmarks/bin/Release/net10.0/XLibur.Benchmarks.exe --filter "*TemplateRoundTrip*" "*StreamingWrite*" "*SheetGeometry*"
XLibur.Benchmarks/bin/Release/net10.0/XLibur.Benchmarks.exe --filter "*CellStyling*" "*FormulaEvaluation*" "*StyleKeyHashCode*" "*AllocationBenchmarks*"

XLibur.Benchmarks/bin/Release/net10.0/XLibur.Benchmarks.exe profile alloc
XLibur.Benchmarks/bin/Release/net10.0/XLibur.Benchmarks.exe profile create
XLibur.Benchmarks/bin/Release/net10.0/XLibur.Benchmarks.exe profile template
```

**Do not pass `--job short`.** `Program.cs` adds an explicit job for the in-process toolchain, and
`--job` *adds* a second one, so every benchmark runs twice. Individual characteristics
(`--warmupCount`, `--iterationCount`, `--launchCount`) are mutators and modify the existing job
instead — that is the supported way to shorten a run here.

---

## Area 1 — `CellsUsed()` costs more than loading the file, and buffers the whole sheet before yielding

**Size:** M · **Risk:** M (public enumeration semantics) · **Prize:** the largest number in the suite
**Status:** Tasks 1.1–1.3 done — see [Area 1 results](#area-1-results). 1.4 **declined on evidence**; 1.5 still open.

### The measurement

Three ways to visit the same 3.75 M cells, each already loaded:

| API | extra time | extra allocation |
|---|---:|---:|
| `ws.EnumerateUsedCells()` | +0.36 s | **+0.00 MB** |
| `ws.Cell(r,c).GetValue<string>()` in a nested loop | +1.55 s | +622.9 MB |
| `ws.CellsUsed()` | +2.39 s | +1,019.8 MB |

`EnumerateUsedCells` establishes the floor and it is *zero* — `LoadAndIterateEnumerateUsedCells`
allocates 334.54 MB, the same figure as `LoadWorkbook`, to the byte. Whatever `CellsUsed()` spends is
therefore overhead in the enumeration itself, not an unavoidable cost of reaching the data. It is
three times what parsing the file allocates.

### The mechanism

`XLibur/Excel/Cells/XLCells.cs`, `GetUsedCells` (line 73):

```csharp
var visitedCells = new HashSet<XLAddress>();
...
var cells = worksheetGroup.SelectMany(addr => GetUsedCellsInRange(addr, ws, usedCellsCandidates))
    .OrderBy(cell => cell.Address.RowNumber)
    .ThenBy(cell => cell.Address.ColumnNumber);
```

Three distinct costs, in descending order of confidence:

1. **The `OrderBy`/`ThenBy` materialises and sorts every cell before yielding the first one.** So
   `CellsUsed()` is not lazy in any useful sense — `ws.CellsUsed().First()` walks and sorts the whole
   sheet, and peak memory holds an array of 3.75 M `XLCell` references regardless of what the caller
   does with them.
2. **The sort is provably redundant for the single-range case.** `Slice<T>.Enumerator.MoveNext`
   (`XLibur/Excel/Cells/Slice.cs`, line 459, comment: *"The movement is columns first, then rows"*)
   already yields row-major — exactly the order `OrderBy(row).ThenBy(column)` produces. For one range
   on one sheet the comparison sort re-establishes an order the source guaranteed.
3. **`visitedCells` retains one `XLAddress` per used cell for the lifetime of the enumeration**, purely
   to deduplicate. Deduplication is only needed when the address ranges can overlap; a single range,
   or a set of disjoint ones, cannot produce a duplicate.

On top of all three, one `XLCell` wrapper is minted per cell — 48.1 B and 19.8 ns each by
`profile create`, i.e. 180 MB and 74 ms of the totals above before any of the LINQ.

### What to do

| # | Task | Status | Size |
|---|---|---|---|
| 1.1 | A benchmark that isolates the shapes: one range vs several disjoint vs several overlapping; and `.First()` / `.Take(10)` against a full walk, which is where the eager sort shows worst. Land this before any fix. | ✅ Done — `UsedCellEnumerationBenchmarks` | S |
| 1.2 | Skip the sort when the enumeration covers one range on one sheet and the candidate set is empty — the slice already yields row-major. Assert the property in a test rather than assuming it. | ✅ Done | S |
| 1.3 | Skip `visitedCells` when the range addresses provably cannot overlap (count == 1, or a pairwise disjointness check that is cheap for the small counts seen in practice). | ✅ Done — **better than specified**: adjacent-duplicate rejection removes the set for *every* shape, so no disjointness analysis was needed | S |
| 1.4 | Reuse the `XLCell` wrapper across the enumeration, or route `IXLCells` onto the slice enumerator with the wrapper vended lazily. **This is the risky one** — Spec 11 task 3 already audited `ReferenceEquals` on `XLCell` and found nothing depends on it, but a wrapper reused *within* one `foreach` is a stronger change than a cached one, because a caller who buffers `CellsUsed().ToList()` would get N references to one mutated object. Cap the task at 1.2/1.3 if the audit says so, and say why in the PR. | ❌ **Declined on evidence** — task 1.1 priced it at 24 MB of an 84.66 MB total, against the only semantic risk in the area | M |
| 1.5 | Same treatment for `LoadAndReadAllCells`'s path: `ws.Cell(r,c).GetValue<T>()` at 413 ns / 166 B. `SetCellValue` shows the write side already has a wrapper-free internal seam; find or add the read equivalent. | ⬜ Open — now the largest remaining read-path term | M |

### Acceptance criteria

1. `LoadAndIterateCellsUsed` allocation reduced ≥ 50% (1,354 MB → ≤ 677 MB) with time reduced ≥ 20%.
2. `ws.CellsUsed().First()` on the 250K × 15 fixture completes in time proportional to the position of
   the first used cell, not to the sheet — demonstrated by a benchmark, not asserted.
3. Enumeration order is unchanged for every shape in task 1.1, including overlapping ranges.
4. All tests green; no public API change.

### Not established (at the time of writing — now answered, see below)

- Whether the sort or the `HashSet` dominates. Task 1.1 exists to answer that before either is
  touched, and the answer decides whether 1.4 is worth its risk.
- `AllocationBenchmarks.IsEmptyConditionalFormats` reports 640 B per `IsEmpty(XLCellsUsedOptions.All)`
  call, and `GetUsedCellsInRange` calls `IsEmpty(_options)` per cell. Whether the default options
  reach the same path is **unmeasured** — 640 B × 3.75 M would exceed the whole measured overhead, so
  they almost certainly do not. Treat as a lead, not a finding.

<a id="area-1-results"></a>
### Area 1 results

#### Task 1.1 — the attribution, and how it re-ranked the rest

`UsedCellEnumerationBenchmarks`, 50,000 × 10 = 500,000 used cells. Each rung adds one layer to the
one beneath it, so a difference is that layer's cost.

| Rung | Mean | Δ time | Allocated | Δ alloc |
|---|---:|---:|---:|---:|
| L1 slice enumerator only | 5.20 ms | — | 88 B | — |
| L2 + the `XLCell` wrapper | 11.83 ms | +6.6 ms | 24.00 MB | +24.00 MB |
| L3 + `IsEmpty` filter | 17.93 ms | +6.1 ms | 24.00 MB | **+0** |
| L4 + `OrderBy`/`ThenBy` | 83.71 ms | **+65.8 ms** | 36.54 MB | +12.54 MB |
| L5 + `HashSet<XLAddress>` | 106.48 ms | **+22.8 ms** | 84.01 MB | **+47.47 MB** |
| L6 real `CellsUsed()` | 101.75 ms | ≈ L5 | 84.66 MB | ≈ L5 |

L5 reconstructing L6 within noise is what says the ladder accounts for the whole cost rather than
most of it.

**This inverted the spec's own ranking.** The sort and the visited set are **87% of the time and 71%
of the allocation**; the `XLCell` wrapper — which this spec sized as the M-effort, reference-identity-
risky task 1.4 — is 6.6 ms and 24 MB, and is the *floor* for an API that yields `IXLCell` handles.
Task 1.4 is therefore **declined**: it carries the only real semantic risk in the area and, per the
table, could at best remove 24 MB of an 84.66 MB total. The two tasks the spec sized S are where
everything was.

`IsEmpty` was also cleared: L3 − L2 is **zero allocation**. The 640 B/call lead recorded above does
not reach this path, exactly as suspected — `IsEmpty` short-circuits on `!IsContentEmpty()` and every
cell here has content.

#### Tasks 1.2 and 1.3 — what shipped

`XLCells.GetUsedCells` now dispatches: a single range on one sheet with no candidate cells streams
straight off `GetUsedCellsInRange`; everything else keeps the sorted path.

The justification is stronger than "the slice happens to be ordered". `XLCellsCollection.GetCells`
reads through `SlicesEnumerator`, a **k-way merge** over the value, formula, style and misc slice
enumerators: it selects the smallest `Point` each step and advances *every* enumerator sitting on it.
So its output is strictly ascending and already duplicate-free. `Point` packs the row above the
column and `CompareTo` compares the packed value, which makes ascending packed order exactly
row-major — the order `OrderBy(row).ThenBy(column)` produces. On that path the sort re-sorts sorted
input and the visited set can never reject anything.

The sorted path got a cheaper dedup rather than a fast path: once a sequence is sorted by
(row, column), duplicates are **adjacent**, so comparing against the previous address rejects exactly
what a set of every address seen would, at O(1) memory instead of one entry per used cell. That is
why the multi-range shapes improve too, without any disjointness analysis. (`default(XLAddress)`
packs to the same value as a relative `A1`, so the "previous" slot needs an explicit
have-we-seen-one flag — not a null check.)

`HasCandidates` tests the **sheet** as well as the options: asking for merged ranges on a sheet that
has none still leaves the candidate sequence empty, and that is the common shape for
`CellsUsed(XLCellsUsedOptions.All)`. Where it cannot cheaply prove emptiness it answers "yes" and
falls back, so it is conservative in the safe direction.

#### Measured, A/B in one sitting

`UsedCellEnumerationBenchmarks`, 500,000 cells:

| Benchmark | before | after | Δ time | Δ alloc |
|---|---|---|---:|---:|
| **`L6_CellsUsed`** | 101.75 ms / 84.66 MB | **20.01 ms / 24.00 MB** | **−80.3%** | **−71.7%** |
| `Shape_TenDisjointRanges` | 170.74 ms / 125.0 MB | 135.19 ms / 77.69 MB | −20.8% | −37.9% |
| `Shape_TenOverlappingRanges` | 160.80 ms / 82.77 MB | 138.63 ms / 59.17 MB | −13.8% | −28.5% |
| **`EarlyExit_CellsUsedFirst`** | 75.28 ms / 35.84 MB | **265.3 ns / 1,072 B** | **−99.9996%** | −99.997% |
| `L1_SliceOnly` *(control)* | 5.20 ms / 88 B | 4.92 ms / 88 B | −5.4% | 0 |
| `L2_PlusWrapper` *(control)* | 11.83 ms / 24,000,588 B | 11.04 ms / 24,000,598 B | −6.7% | +10 B |
| `L3_PlusIsEmptyFilter` *(control)* | 17.93 ms / 24,000,520 B | 16.39 ms / 24,000,676 B | −8.6% | +156 B |
| `EarlyExit_EnumerateFirst` *(control)* | 55.75 ns / 88 B | 56.29 ns / 88 B | +1.0% | 0 |

After the change `CellsUsed()` allocates 24.00 MB — **the same figure as L2**, i.e. the `XLCell`
wrappers and nothing else. That is the floor for this API without changing what it yields, and it is
reached.

`EarlyExit_CellsUsedFirst` is the number that matters most for real callers: `CellsUsed().First()`
went from walking and sorting all 500,000 cells to 265 ns. The eager sort meant every early-exit
pattern — `.First()`, `.Any()`, `.Take(n)` — paid for the whole sheet.

L4 and L5 are the ladder's own reconstruction code, not library code, and their allocation is
byte-identical across both arms (36,540,913 B for L4 in both runs). Their ~19% time movement is
machine state on the two noisiest rungs in the table — they allocate 36–84 MB per operation and
collect gen2 — and their before-arm standard deviations were 7.2 ms and 14.3 ms. Read them as
unchanged; the four genuine controls above move within the documented ±9%.

`XLiburReadBenchmarks`, 250,000 × 15, **A/B in one sitting**, 2 warmup + 12 iterations (the 3
iterations used for the survey table are not enough for a 6 s operation here — an early attempt
showed the *unchanged* control moving 17% with a ±8 s error bar):

| Benchmark | before | after | Δ time | Δ alloc |
|---|---|---|---:|---:|
| `LoadAndIterateCellsUsed` | 6.242 s / 1,354.38 MB | **5.245 s / 853.03 MB** | **−16.0%** | **−37.0%** |
| `LoadAndIterateEnumerateUsedCells` *(control)* | 3.868 s / 334.55 MB | 3.974 s / 334.55 MB | +2.7% | **0.00 MB** |

Gen2 collections per 1,000 operations fall 5,000 → 3,000, matching the control — the eager sort was
holding every wrapper in the sheet alive simultaneously, and they now die in gen0.

#### Acceptance criteria: 2, 3 and 4 met; **1 was mis-stated and is restated here**

> 1. `LoadAndIterateCellsUsed` allocation reduced ≥ 50% (1,354 MB → ≤ 677 MB) with time reduced ≥ 20%.

**Not met as written: −37.0% and −16.0%.** The criterion charged this area with costs that do not
live in it. Of the 853.03 MB that remain, only ~180 MB is the `XLCell` wrapper (24 MB per 500,000
cells, scaled), and the rest is the benchmark's `cell.Value` reads — which on this fixture include
250,000 formula cells whose values go through `XLCell`'s evaluating value path, not through the
enumeration at all. That is task 1.5's territory and Area 5's, not 1.2/1.3's.

A criterion this area can actually own: **`CellsUsed()` allocates no more than the wrappers it
yields**, which the L6 = L2 result above satisfies exactly. This is the third time in this spec
family that a criterion has been written against a workload number and turned out to price work the
task cannot reach — spec 11 criterion 2 and spec 05 criterion 3 were the first two.

2. ✅ `CellsUsed().First()` is now proportional to the position of the first used cell — 265 ns
   against 75.28 ms, plus a test that counts predicate invocations.
3. ✅ Enumeration order unchanged for every shape. `UsedCellEnumerationOrderTests` (9 tests) covers
   single range, disjoint, overlapping, the same range added twice, merged-range and data-validation
   candidates, the option-set-but-sheet-is-clean case, and predicate filtering. **Eight of the nine
   pass against the pre-change implementation**, which is what makes them characterisation tests
   rather than a restatement of the new code. The ninth is the laziness guarantee, and it fails
   against the old code with 10,000 predicate calls where the new bound is 4 — that failure also
   confirmed the double-invocation noted below.
4. ✅ 11,837 tests pass on net8.0 and net10.0; no public API change.

#### Left undone deliberately

- **The predicate is invoked twice per cell.** `GetUsedCellsInRange` passes `_predicate` to
  `GetCells`, which filters with it, and then tests `_predicate(cell)` again itself. Removing either
  is behaviour-visible in one edge case — the surviving call would run *after* `IsEmpty` rather than
  before, so a predicate that threw on an empty cell would stop throwing — and it is worth ~3% of the
  original figure. Left for a PR that can argue that edge case on its own.
- **Task 1.5** (`ws.Cell(r,c).GetValue<T>()` at 413 ns / 166 B) is untouched and is now the larger
  remaining term on the read path.

---

## Area 2 — Per-cell styling, and a transition cache that was eight slots deep

**Size:** S–M · **Risk:** L · **Prize:** the 204 MB create phase
**Status:** Task 2.3 done and it carried the win — see [Area 2 results](#area-2-results). 2.1 and 2.2
**disconfirmed by reading the code**; 2.4 still open.

> **Correction.** This area was first written as "style writes bypass the transition cache that
> exists to serve them", on the strength of reading `XLStyleValue.WithNumberFormat` in isolation and
> finding it does a full key hash and repository probe. Two of the three mechanisms proposed from
> that reading are wrong, and both were disproved by looking at the callers rather than by
> measuring. The original text is kept below each correction, because the way it failed is the
> point: a mechanism derived from one method without its call sites is a hypothesis, not a finding.

### The measurement that started it

From `profile create`, bytes exact:

| | Bytes/op | ns/op |
|---|---:|---:|
| `ws.Cell(r,c).Value = double` | 103.5 | 148.3 |
| `ws.Cell(r,c).Value = DateTime` | 136.5 | **487.7** |

Writing a date costs 3.3× writing a number and 33 bytes more. A date and a number are stored
identically — both are a serial `double` in the value slice — so the difference is not the value.

### ~~Task 2.1 — route the date rules through the transition cache~~ ❌ Disconfirmed

The claim was that `XLValueStyleRules.WithDateTimeFormat` → `XLStyleValue.WithNumberFormat` runs a
`with` over the seven-field `XLStyleKey`, a full composite hash and a repository probe **per date
cell**, while `XLStyle.ModifyNumberFormat` serves the identical transition from the cache.

`WithNumberFormat` does do all of that. It is just not called per date cell.
`XLWorksheet.GetStyleForDateTime` already holds a one-entry memo keyed on reference equality of the
source style (`_cachedDateOnlySourceStyle` / `_cachedDateOnlyResultStyle`, and the same pair for
date-with-time and for durations), so a column of dates under a stable base style calls
`WithDateTimeFormat` once. `GetStyleForText` goes further and skips reading the style at all for
text that needs no adjustment.

What a date write actually costs, over a number write:

- `GetStyleValue(point)` — a style-slice read that falls through to `GetInheritedStyleValue`, which
  is two dictionary probes and a `Combine` that short-circuits to the sheet style when the row and
  column styles match it (the common case, non-allocating);
- **a style-slice write**, because the cell now carries a number format it did not before.

The 33 B is that slice entry. Spec 11 measured bulk styling at **~33 bytes per cell of pure slice
storage** by a completely different route, which is a good independent check on the attribution.
It is the same conclusion spec 18 task 2 reached for per-cell styling: the allocation is the
style-slice write and it is inherent. **A date needs a format; a format needs a slice entry.** There
is no allocation win here.

The 487.7 ns is a single-shot probe figure and this spec's own probe output warns against reading it
as a time claim. Nothing has established that the *time* gap is anything but the slice write plus
two dictionary probes.

### ~~Task 2.2 — memoise `XLBorderKey.GetHashCode`~~ ❌ Not worth doing

The claim rested on `BorderKey_GetHashCode` measuring 2,572.9 µs per 100,000 against 339.1 µs for
the composite `StyleKey` — "7.6× cheaper than one of its own components".

The comparison is not like for like. `XLStyleKey` memoises each component's hash **in the `init`
accessor**, so the composite is cheap precisely because it never re-hashes anything, and a
`Key with { Border = k }` recomputes only the border hash and reuses the other five. There is no
redundant work to remove; `XLBorderKey` is simply a large struct (five colours plus four styles) and
25.7 ns is about what hashing it should cost — `XLColorKey` alone measures 2.82 ns.

Sizing it against the workload it was supposed to help: `CreateFormattedAndSave` performs roughly
250,000 border mutations, so the whole cost is ~6.4 ms of a 1,020 ms benchmark. **0.6%.** Declined.

### Task 2.3 — instrument the transition cache ✅ This is where the win was

Counters temporarily added to `GetTransition`/`StoreTransition`, over the create phase of
`CreateFormattedAndSave` (50,000 rows), after a warm-up pass:

| | count |
|---|---:|
| probes | 1,033,393 |
| hits | 783,242 (75.8%) |
| misses | 250,151 (24.2%) |
| — slot **evicted** (held a different transition) | **249,998** |
| — slot empty (cold) | 0 |
| — key mismatch | 0 |
| — no cache array yet | 153 |
| stores (one `TransitionEntry` allocation each) | 250,151 |

Every miss but 153 is an eviction. Not cold, not colliding on keys — **too small**. The 153 tells
the rest of the story: only 153 distinct base styles ever receive a transition in this workload, and
the benchmark applies about 109 distinct transitions to each of them. Eight slots cannot hold 109
entries, so they evicted each other and a quarter of a million style derivations were repeated and
re-allocated.

Sweeping the size:

| slots | hit rate | misses | create phase |
|---:|---:|---:|---:|
| **8** (original) | 75.8% | 250,151 | 211.8 MB |
| 16 | 98.4% | 16,672 | 188.6 MB |
| 32 | 96.8% | 33,365 | 190.0 MB |
| 64 | 98.4% | 16,677 | 189.2 MB |
| 128 | 98.4% | 16,677 | 189.4 MB |

16,677 is the **compulsory** miss floor — one per distinct transition per base style
(153 × ~109). It is reached at 16 slots and does not improve after.

Shipped at **64**, not the 16 that first reaches the floor, because 32 measured *worse* than 16.
That is hash-versus-modulus aliasing, and it means the exact value interacts with one fixture's hash
pattern; choosing the minimum that happened to work would be fitting that pattern. The array is
allocated lazily on first store, so the cost is 64 references per base style that actually receives
a transition — 153 of them here — not per style in the workbook.

<a id="area-2-results"></a>
### Area 2 results

BenchmarkDotNet, **A/B in one sitting**:

| Benchmark | before (8) | after (64) | Δ time | Δ alloc |
|---|---|---|---:|---:|
| `CreateFormattedAndSave` | 1,019.6 ms / 320.25 MB | **981.3 ms / 312.47 MB** | −3.8% | −2.4% |
| `CreateAndSave` *(control)* | 259.6 ms / 60.51 MB | 260.9 ms / 60.51 MB | +0.5% | byte-identical |
| `CreateAndSaveFastestCompression` *(control)* | 162.8 ms / 60.59 MB | 162.0 ms / 60.58 MB | −0.5% | −10 B |

`CellStylingBenchmarks` is unmoved to the byte across all seven variants — it applies one or two
distinct transitions, which eight slots already held. That is the control that shows the change
reaches only what it should.

`profile alloc`, cold single run, also A/B in one sitting:

| | before (8) | after (64) | Δ |
|---|---:|---:|---:|
| `CreateFormattedAndSave` create phase | 206.0 MB | **185.1 MB** | −10.1% |
| total | 323.7 MB | 302.7 MB | −6.5% |
| `CreateAndSave` *(control)* | 60.1 MB | 60.1 MB | — |

**The two disagree in magnitude and that is not resolved.** BenchmarkDotNet sees −7.8 MB where the
cold probe sees −21 MB. The transition caches hang off `XLStyleValue` instances in a process-wide
repository, so they survive across benchmark iterations and a benchmark process amortises the
population cost differently from a caller that builds one workbook — that is the likely cause, but
it predicts the gap in the wrong direction and has not been demonstrated. **The conservative
BenchmarkDotNet figure is the claim**; a caller creating one formatted workbook per process may do
better, and anyone who needs that number should measure it rather than take −10.1% from here.

### Still open

- **Task 2.4 — the per-mutation cost itself.** `profile create` puts four style mutations on one
  cell at 473.1 B / 603.6 ns against 128.1 B / 53.2 ns for building the façade and setting nothing,
  so roughly 85 B and 119 ns per mutation, of which ~33 B is the inherent slice entry. Bulk styling
  reaches 33 B/cell total. What the remaining ~52 B per mutation is has not been decomposed. Spec 18
  task 2 warns that per-step attributions in this code invert between 20K and 100K rows — do not
  motivate this from the `profile create` split alone.
- **Whether a bigger cache helps any workload but this one.** The 153-base-style / 109-transition
  shape is one fixture's. A template-driven workload with thousands of base styles would pay 64
  references for each one that receives a transition, and nobody has measured that shape.

---

## Area 3 — Saving costs more than loading, and a third of it is deflate

**Size:** M–L · **Risk:** M · **Prize:** ~30% of every save · **Overlaps:** Spec 01 phase 3, Spec 03 task 7

### The measurement

The same 420,000 cells, three ways (`profile template` grid probes plus `LoadRowHeavy`):

| | Time | per cell |
|---|---:|---:|
| build the in-memory model | 171.1 ms | 0.41 µs |
| load the saved file | 385.9 ms | 0.92 µs |
| **save it** | **438.1 ms** | **1.04 µs** |

And the compression knob, on the 50K × 3 workload, measured on both write paths:

| | `Optimal` (default) | `Fastest` | Δ |
|---|---:|---:|---:|
| `CreateAndSave` (`System.IO.Packaging`) | 255.5 ms / 60.51 MB | **179.5 ms** / 60.59 MB | **−30% time, 0% allocation** |
| `StreamingWrite` (own `ZipArchive`) | 163.6 ms / 13.60 MB | **94.7 ms** / 17.01 MB | −42% time |

Both paths pay ~70–76 ms of deflate on the same data, and on the ordinary path that is **30% of the
entire create-and-save benchmark** for zero allocation difference. `SaveOptions.CompressionLevel`
exists and defaults to `Optimal`; nothing in the suite measured it until this survey added
`CreateAndSaveFastestCompression`.

The second measurement in that table is the more uncomfortable one. The streaming writer produces the
same workbook from the same data in **163.6 ms and 13.6 MB** where the model path takes **255.5 ms and
60.5 MB** — 1.6× the time and 4.5× the allocation. Spec 03 task 7 already attributed ~96 MB of the
formatted save phase to `System.IO.Packaging` buffering whole parts in a doubling `MemoryStream`, and
Spec 01 found the fix while building the streaming writer: own the zip, write in `Create` mode, never
buffer a part. That fix is shipped — it is just not on the path `SaveAs` uses.

### What to do

| # | Task | Size |
|---|---|---|
| 3.0 | **Done in this survey.** `CreateAndSaveFastestCompression` added to `XLiburWorkbookBenchmarks`, which is where the −30% above comes from. | XS |
| 3.1 | Decide and document the compression default. `Optimal` → `Fastest` is 30% of save time; measure the file-size cost on the corpus and put both numbers in the docs so callers can choose. Changing the default is a product decision, not a perf one — bring the numbers, not a patch. | S |
| 3.2 | Measure how much of the model path's 92 ms / 47 MB gap to the streaming writer is packaging and how much is the model. Redirect `SaveAs` at `Stream.Null` (Spec 03 used this technique) to split them before designing anything. | S |
| 3.3 | Only if 3.2 says packaging: prototype `SaveAs` over `StreamingPackageWriter` for the clean-workbook case (no loaded package to preserve). **The hard constraint is `docs/round-trip-fidelity.md`** — saving a *loaded* workbook reopens the original package and rewrites only modelled parts, and that behaviour is pinned by tests. A second package writer must either preserve it or be restricted to workbooks with no origin package. | L |

### Acceptance criteria

1. A documented compression trade-off: time and output size for all four levels on at least three
   corpus workbooks.
2. Task 3.2 produces a split, in the PR description, of the `CreateAndSave`-versus-`StreamingWrite`
   gap into packaging and model.
3. If 3.3 is attempted: byte-identical output for clean workbooks, and every
   `docs/round-trip-fidelity.md` test still green.

### Not established

- **That 3.3 is feasible at all.** It is the reason this area is sized L and listed third rather than
  first. Tasks 3.1 and 3.2 are worth doing on their own and neither depends on it.
- Whether the deflate share holds for the formatted workload. `CreateFormattedAndSave` was not run
  with `Fastest`; its save phase is a larger fraction of a much larger total, so the share will differ.

---

## Area 4 — Load is ~1 µs and ~93 B per cell, and nothing has decomposed what remains

**Size:** M · **Risk:** M · **Prize:** the floor under every read number in the suite
**Overlaps:** Spec 02 (done), Spec 18 task 5 (open)

### The measurement

`LoadWorkbook` on 3.75 M cells: **3.717 s / 334.54 MB** — 0.99 µs and 93.5 B per cell.
`LoadRowHeavy` on 420 K cells corroborates independently at 0.92 µs / 138 B per cell.

Load is 71% of `LoadAndReadAllCells` and 61% of `LoadAndIterateCellsUsed`. Every improvement in Area 1
lands on top of this floor, and after Area 1 the floor *is* the number.

Spec 02 took load from 4.750 s / 1,020.92 MB to 3.968 s / 392.88 MB and recorded the reason it stopped:
*"after these three tasks, load time is no longer dominated by garbage, so further gains need a
different lever"*, naming per-sheet parallelism as the next structural one. The 334.54 MB now measured
has never been broken down. At 93.5 B per cell it is plausibly all real storage — the value slice, the
formula slice for 250 K formula cells, the shared strings — but *plausibly* is the operative word.

Spec 18 task 5 attacks the same code from the other end and is still open: a structurally empty
worksheet costs ~202 KB to round-trip, of which **41% is load** — the largest untouched term in that
spec, and its Results section explicitly names it as where a follow-up should start.

### What to do

| # | Task | Size |
|---|---|---|
| 4.1 | GC-exact decomposition of the 334.54 MB, the way `profile alloc` splits create from save: value slice, formula slice, shared strings, style cache, reader buffers, everything else. Nothing else in this area should start first. | S |
| 4.2 | A single-sheet load benchmark with the formula column removed, to price the 250 K `XLCellFormula` objects separately from the 3.5 M value cells. | S |
| 4.3 | Pipeline the sheet parse: the raw `XmlReader` pass over `<sheetData>` and the slice writes are separable, and the reader is I/O-and-tokenise while the writer is allocate-and-store. Prototype before designing — the workbook is not thread-safe and the win must clear that cost. | L |
| 4.4 | Per-sheet parallel load, which Spec 02 named. Note the benchmark fixture is **one sheet**, so this needs its own fixture and would not move any number in the table above. | M |
| 4.5 | Take Spec 18 task 5's load half with whatever 4.1 finds. The two are the same code seen at different scales. | M |

### Acceptance criteria

1. Task 4.1 publishes a decomposition table summing to within 5% of 334.54 MB, with every line ≥ 5%
   of the total named and attributed to a call site.
2. Any subsequent task cites that table for its premise. (This spec has three predecessors — 05, 11
   and 18 — whose first attributions were wrong; the rule exists because of them.)
3. `LoadWorkbook` time reduced ≥ 15% or the reason it cannot be is recorded with numbers.

### Not established

- Everything past 4.1. The tasks below it are candidate levers, in the order they look plausible, and
  4.1 exists to replace that ordering with evidence.
- Whether 93.5 B/cell has any slack in it at all. It may be near the floor for the data, in which case
  the honest outcome of this area is a documented "no" — which is worth having, because two specs
  currently assume otherwise.

---

## Area 5 — Formula evaluation: fix the benchmark before believing anything about it

**Size:** S to establish, unknown to fix · **Risk:** L for the benchmark work · **Overlaps:** Spec 04 (open)

### The measurement, and why it does not say what it appears to

| Benchmark | Mean | Allocated |
|---|---:|---:|
| `UniqueSameSheet` | 16.41 ms | 10.38 MB |
| `SharedSameSheet` | 13.60 ms | 10.38 MB |
| `SharedCrossSheet` | **42.09 ms** | 10.38 MB |

3.1× for a cross-sheet reference, at identical allocation, looks like a clean finding. **It is not
one.** From `FormulaEvaluationBenchmarks.Setup`:

```csharp
sharedSheet.Cell(row, 6).FormulaA1 = "SUM($D$1:$E$1)";                    //  2 cells
crossSheet .Cell(row, 6).FormulaA1 = $"SUM(Lookup!$A$1:$A${LookupRows})"; // 20 cells
```

The cross-sheet variant sums a ten-times-larger range. The comparison confounds sheet resolution with
range size, and on the arithmetic alone — 18 extra cells summed 20,000 times — most or all of the
28 ms gap could be the summing. No conclusion about cross-sheet references can be drawn until the
range sizes match.

This is exactly the failure mode Spec 18 records twice in its own history: a fixture whose intended
variable moved together with an unintended one, producing a confident wrong attribution. It is caught
here before it becomes a work item.

### The candidate mechanism, if the corrected benchmark still shows a gap

`PrefixNode.GetWorksheet` (`XLibur/Excel/CalcEngine/AstNode.cs`, line 247) calls
`wb.TryGetWorksheet(Sheet!, out …)` on **every evaluation**, which reaches
`XLWorksheets.TryGetWorksheet` (line 69) and runs `sheetName.UnescapeSheetName()` plus a dictionary
probe before returning. The `ReferenceNode` above it memoises the resolved `Reference` keyed on the
sheet instance (added by #286), so the address build is already cached — but the memo is only
consulted *after* the name lookup that produces the key. Allocation being identical across all three
benchmarks says `UnescapeSheetName` returns its input unchanged when there is nothing to unescape, so
whatever cost is there is CPU, not garbage.

### The larger open question this area sits next to

Spec 04 (demand-driven evaluation) is still proposed. Its premise — that reading one dirty formula
cell can trigger a full-workbook recalculation, building a dependency tree of ~176 MB to answer one
read — **has no benchmark anywhere in the suite.** `TryEvaluateSingleCell` has since landed
(`XLCalcEngine.cs`, line 256) and takes the single-cell fast path, but its `catch
(GettingDataException)` still falls through to `Recalculate(sheet.Workbook, null)` for any formula
whose precedent is dirty, which is the cliff Spec 04 describes. Nothing measures how often that
happens or what it costs.

### What to do

| # | Task | Size |
|---|---|---|
| 5.1 | Fix `FormulaEvaluationBenchmarks`: hold the summed range size constant across `SharedSameSheet` and `SharedCrossSheet`. Re-measure. If the gap collapses, record that and close the cross-sheet question. | XS |
| 5.2 | Only if a gap survives 5.1: hoist the sheet resolution behind the existing `_sheetReference` memo, or cache the resolved `XLWorksheet` on the `PrefixNode` with the same rename/delete invalidation the reference memo already uses. | S |
| 5.3 | Build the benchmark Spec 04 has always lacked: a workbook of ~100 K dirty formulas, read 100 random cells, and count evaluations with an internal counter. That number — evaluations per read — is Spec 04's whole case, and nobody has ever produced it. | M |

### Acceptance criteria

1. `SharedSameSheet` and `SharedCrossSheet` differ in exactly one variable, evidenced by the fixture
   code in the PR.
2. Task 5.3 publishes evaluations-per-read for the dirty-formula workload, whatever it turns out to
   be. Spec 04's task 6 asks for this benchmark; delivering it here would let Spec 04 be scheduled or
   declined on evidence rather than on its own estimate.

### Not established

- Any cross-sheet penalty at all, until 5.1.
- That Spec 04's cliff is reachable in practice on a loaded workbook. It might be common, it might be
  rare; 5.3 is the only way to find out and it is cheap relative to Spec 04's L estimate.

---

## Priority

| | Area | Prize | Confidence in the mechanism | Risk |
|---|---|---|---|---|
| 1 | **Area 1** — `CellsUsed()` enumeration | ✅ **done**: −80% time / −72% allocation on the enumeration; `.First()` 75 ms → 265 ns | — | — |
| 2 | **Area 2** — per-cell styling | ✅ **partly**: transition cache resized, −3.8% time / −2.4% allocation on `CreateFormattedAndSave`. Two of its three proposed mechanisms were wrong. | — | L |
| 3 | **Area 3** — deflate and packaging | 30% of save wall time, measured | High for 3.1/3.2, unknown for 3.3 | M |
| 4 | **Area 4** — the load floor | 3.72 s / 335 MB under everything | Low — undecomposed by design; 4.1 is the work | M |
| 5 | **Area 5** — formula evaluation | unknown | None yet — the benchmark is confounded | L |

Areas 1, 2 and 3 are fully independent and can run in parallel. Area 4 task 4.1 should start early
regardless, because Areas 1 and 4 stack on the same benchmark and whoever measures second will
otherwise be measuring the other's change. Area 5 task 5.1 is an hour's work and unblocks the rest of
its own area.

## Already ruled out, or explained, by this survey

- **The template round trip has not regressed.** All five `TemplateRoundTripBenchmarks` reproduce
  Spec 18's post-task-1 figures with allocation identical to the byte. Spec 18's results stand.
- **Sheet geometry and string uniqueness are still inherent.** `SheetGeometryBenchmarks` reproduces
  Spec 18 task 3's 2×2 exactly. A row costs what a row costs; a distinct string is stored once.
  Nothing to fix, and re-deriving it would be the third time.
- **`StyleKey.GetHashCode` is not a hotspot.** 339.1 µs per 100,000 after Spec 03 task 3. Its
  *components* are (Area 2, task 2.2), which is the opposite of the original finding and worth not
  mis-remembering.
- **Bulk styling is already fixed.** `ws.Range(all).Style.Bold` + populate is 88.6 B/op against
  473.1 B for four individual mutations. Spec 11 task 4 did this; the remaining gap is on the
  per-cell path, which is Area 2.
- **`ToExcelFormat` at 437 B per call** (`AllocationBenchmarks`) is the largest per-call allocation in
  the micro set and is *not* on the save hot path — Spec 03 task 1 moved cell number writing onto a
  span-based `TryFormat`. It is reached by callers formatting a value for display. Left alone
  deliberately; noted so the next survey does not re-find it and assume it is hot.

## Ground rules for implementing agents

These are the repo's, restated because this spec will be handed to agents who have not read
`docs/specs/README.md`:

- **Branch per area/task; never commit to main.** Commit style: `perf:`, `feat:`, `fix:`, `refactor:`.
- **Warnings are errors**; nullable is enabled.
- **Perf PRs must carry before/after BenchmarkDotNet tables**, A/B'd in one sitting by stashing only
  the library (`git stash push -- XLibur/`) so the benchmark project is byte-identical across arms.
- **The noise floor on the reference machine is 4.5–9%** on the write benchmarks and worse on
  `CellStylingBenchmarks` (std-dev 2.0–6.3 ms). A single run showing a 10% win has shown nothing.
  Spec 18 task 0's three checks — the `MinIterationTime` warning, the baseline's removed outliers, and
  gen2 collection in `[IterationCleanup]` — apply to every benchmark in this spec.
- **A/B the controls, not just the target.** A change that moves benchmarks it cannot reach is
  measuring the machine.
- **Do not upgrade SixLabors.Fonts** (license conflict).
- Line numbers here are from 2026-08-07 against `8c207377` — verify before editing.
