# Spec 21 — Hot-path class-to-struct candidates: the slice enumerators are the only large one

**Area:** Performance (read · enumeration) · Memory (allocation count, not bytes)
**Effort:** M overall; tasks 1 and 2 are ordered, task 3 is independent of both
**Dependencies:** Task 0 first (it is the before-arm). Task 1 → task 2. Task 3 independent. Task 4 gates the spec.
**Status:** ✅ **Complete — task 3 shipped, tasks 1 and 2 declined on measurement.** See
[Results](#results). Task 1 was implemented, measured at **+60% on the primary instrument**, and
reverted; that same measurement disproved task 2's premise before it was written. Task 3 delivered
−13.6% time and −99.996% allocation on the path it targets. **The task-4 decision rule as originally
written was wrong and is corrected in Results.**
The task text below is left as it was proposed, so the parts that turned out to be false stay
readable next to what replaced them.

## Why this spec exists

A review asked which classes on XLibur's hot paths could become structs. The survey covered the
coordinate layer, the cell/slice layer, the calc engine, the style layer, the IO readers and writers,
and the range layer. **The finding that shapes this spec is a negative one: almost everything that
should already be a struct already is.** `Point`, `Area`, `XLAddress`, `XLRangeAddress`,
`SheetPoint`, `ScalarValue`, `AnyValue`, `XLCellValue`, `XLUsedCell`, every style key, the
`SheetDataReadContext`/`CellProperties` load types and the `CellWriteContext` save types are structs
today. Specs 02, 03, 05, 11, 18 and 19 got there first.

So this is a short spec. Ten types were reviewed; **three are worth converting and seven are not**,
and the seven matter as much as the three because each carries a specific disqualifier that should
stop it being re-proposed.

**The one that matters: `Slice<TElement>.Enumerator`.** It sits under every cell walk in the
library, it is a `sealed class` whose every field is a value type, and it is reached through
`IEnumerator<Point>` — so it costs one allocation per enumeration *and* an interface dispatch per
cell. The allocation is trivial. The dispatch is not: `XLCellsCollection.SlicesEnumerator` is a
k-way merge that calls `.Current` and `.MoveNext()` through the interface on up to four enumerators
for every point it yields.

## What this will not do

**This is not a byte-count win.** All three conversions together remove on the order of one
allocation per enumeration and one per range construction — not per cell. Nothing in this spec
targets the per-cell `XLCell` wrapper (spec 19 area 1 task 1.4 declined that on measurement) or the
93.5 B/cell load floor (spec 19 area 4).

The claim being tested is **dispatch and inlining on the enumeration path**, which is a time claim
with no allocation signature. If it does not show up in time, there is nothing else to fall back on,
and task 4 should say so.

**The noise floor is against us.** Spec 19 measured 4.5–9% run-to-run spread on this machine and
records that its own L4/L5 rungs moved ~19% on byte-identical code. `L1_SliceOnly` is a 5 ms
benchmark with an 88-byte allocation profile, so it is one of the *quieter* ones in the suite — but
any claim here needs an A/B in one sitting, and a control that did not move.

---

## Measured baselines

From spec 19, `UsedCellEnumerationBenchmarks`, 50,000 × 10 = 500,000 used cells, post-Area-1:

| Rung | Mean | Allocated | what it exercises |
|---|---:|---:|---|
| `L1_SliceOnly` | 4.92 ms | 88 B | `EnumerateUsedCells` → one `Slice.Enumerator`, one dispatch pair per cell |
| `L2_PlusWrapper` | 11.04 ms | 24.00 MB | + `GetCells` → `SlicesEnumerator` over **four** boxed enumerators |
| `L6_CellsUsed` | 20.01 ms | 24.00 MB | the real `CellsUsed()` |
| `EarlyExit_EnumerateFirst` | 56.29 ns | 88 B | one `MoveNext`, so mostly construction cost |

And from `XLiburReadBenchmarks`, 250,000 × 15 = 3.75 M cells:

| Benchmark | Mean | Allocated | over `LoadWorkbook` |
|---|---:|---:|---|
| `LoadWorkbook` | 3.717 s | 334.54 MB | — |
| `LoadAndIterateEnumerateUsedCells` | 3.974 s | 334.55 MB | **+0.36 s / +0.00 MB** |

**That +0.36 s — 97 ns per cell to walk cells that are already in memory and allocate nothing — is
what task 1 targets.** The 88 B in `L1_SliceOnly` is the enumerator object itself, and it is the
whole allocation figure for the benchmark, which is what makes that rung a clean instrument.

---

## Task 0 — Establish the before-arm

**Effort:** XS · **Blocks:** every other task

Run, on the unchanged parent commit, in the sitting that will also run the after-arm:

```
dotnet build XLibur.Benchmarks/XLibur.Benchmarks.csproj -c Release -f net10.0
XLibur.Benchmarks/bin/Release/net10.0/XLibur.Benchmarks.exe --filter "*UsedCellEnumeration*"
```

`L1_SliceOnly` and `EarlyExit_EnumerateFirst` are the primary instruments; `L2_PlusWrapper` prices
the `SlicesEnumerator` path task 2 changes. `L4`/`L5` are the ladder's own reconstruction code and
spec 19 records them as the two noisiest rungs — **do not read them as controls.**

**Acceptance:** a before table in this file's Results section, from the same process invocation
sequence as the after table.

---

## Task 1 — `Slice<TElement>.Enumerator` and `ReverseEnumerator` become structs

**Effort:** S · **Depends on:** task 0 · **This is the task that matters.**

`XLibur/Excel/Cells/Slice.cs:435` and `:484`. Both are `sealed class : IEnumerator<Point>` whose
every field is already a value type:

```
Area                              _range              16 B
ColumnEnumerator                  _columnsEnumerator  ~40 B
Lut<RowData>.LutEnumerator        _rowsEnumerator     16 B
```

No inheritance, no reference identity, never stored in a field beyond the enumeration that created
it. This is the shape `List<T>.Enumerator` has in the BCL.

Convert both to `struct`. Then:

- **Give the concrete type a `public void Dispose()`.** Today it is `void IDisposable.Dispose()`, an
  *explicit* implementation. This is a clarity change, not a correctness one: it enables
  pattern-based disposal so `foreach` and `using` bind to the struct's own method rather than through
  `IDisposable`.

  **Verified, because the opposite was assumed first:** `using var e = <struct enumerator>` does *not*
  box and does *not* create a defensive copy, even when `Dispose` is implemented explicitly — the
  compiler emits a `constrained.` call, and a `using` local is not `readonly` for this purpose. A
  probe covering plain-local, `using var` with explicit `Dispose`, `using var` with public `Dispose`
  and `using (…)` block forms counts 3 of 3 in every case. So the eight `using` sites
  (`SheetDataWriter.cs:126,168`, `XLCalculationChain.cs:86`, `XLCalcEngine.cs:569`,
  `DependencyTree.cs:67`, `CalculationChainPartWriter.cs:30`, `XLWorkbook_Save.cs:334`,
  `FormulaSlice.cs:159`) and the two non-`using` ones in `XLFormulaShiftPass.cs` and
  `XLRowBatchDelete.cs` need no rework — but **confirm each still binds to the concrete type and not
  to `IEnumerator<Point>`**, because that assignment is where boxing genuinely occurs.
- **Delete `ReverseEnumerator.Dispose`'s `GC.SuppressFinalize(this)`.** The type has no finalizer;
  on a struct the call does not compile against `this` the same way and it was dead code regardless.
- `ISlice.GetEnumerator` returns `IEnumerator<Point>` and must keep doing so — that interface is how
  `SlicesEnumerator` holds four differently-typed slices today. **Keep it, boxing and all**, and add
  concrete-typed accessors beside it. The precedent already exists: `FormulaSlice.GetForwardEnumerator`
  (`FormulaSlice.cs:149`) returns the concrete type. Task 2 is what removes the remaining boxes.
- `ValueSlice.GetEnumerator` (`ValueSlice.cs:63`) forwards to the slice; add a concrete forward
  alongside it so `XLUsedCellEnumerator` can take it.

Then take the first consumer: **`XLUsedCellEnumerator`** (`Excel/Cells/XLUsedCell.cs:63`) holds
`IEnumerator<Point>? _inner` and lazily assigns it. This is the type whose doc comment says
*"without allocating a wrapper per cell"* and which spec 19 measured at **+0.00 MB** — and it still
allocates this one enumerator per iteration and pays two interface dispatches per cell. Replace the
field with the concrete struct plus an explicit `bool _started`, since a struct enumerator cannot use
`null` as the not-yet-created sentinel.

**Consequences to verify, not assume:**

- A mutable struct enumerator held in a `readonly` field or captured by a `foreach` over an
  interface silently operates on a copy and never advances. Every site that stores one must hold it
  in a plain field or local and call through `ref`. This is the failure mode that makes mutable
  structs dangerous, and it is silent — an infinite loop or a zero-iteration walk, not a compile
  error.
- `public ref readonly TElement Current` returning a ref into the LUT's backing array is unchanged
  by the conversion, but confirm the struct is not copied between obtaining the ref and using it.
- `~72 B` is large for a struct. Every copy is a 72-byte memcpy, so `in`/`ref` discipline is not
  optional here — see the task 2 note.

**Acceptance:**
- `L1_SliceOnly` allocation falls from 88 B to **0 B**, and `EarlyExit_EnumerateFirst` likewise.
  This is the part that cannot be noise, so it is the criterion that gates the rest.
- `L1_SliceOnly` time improves, or the reason it does not is recorded with numbers.
- Full suite green on net8.0 and net10.0; no public API change beyond `XLUsedCellEnumerator`'s
  internals.

---

## Task 2 — `SlicesEnumerator` stops going through `IEnumerator<Point>`

**Effort:** M · **Depends on:** task 1 · **This is where the dispatch actually is.**

`XLCellsCollection.SlicesEnumerator` (`XLCellsCollection.cs:615`) is already a struct, but it holds
`IEnumerator<Point>[]` and merges up to four slices — value, formula, style, misc — by selecting the
smallest `Point` each step:

```csharp
private Point FindNextPoint()
{
    var current = _enumerators[0].Current;          // interface dispatch
    for (var i = 1; i < _count; i++)
    {
        var candidate = _enumerators[i].Current;    // interface dispatch
        ...
```

`FindNextPoint` and `AdvanceMatchingEnumerators` both run per yielded point. On a four-slice merge
that is **eight interface calls per cell**, none of them inlinable, and on the 3.75 M-cell fixture
roughly 30 M dispatches for one `CellsUsed()` walk. It also allocates twice on construction — the
`params` array at the call site, then a second array copied from it.

Replace the array with four explicitly-typed struct fields plus a count/liveness mask, or with a
generic merge over a struct-constrained enumerator interface (`where TEnum : struct, IPointEnumerator`)
so the JIT devirtualises each arm. Both work; the four-field version is simpler and the arity is
fixed at four by `XLCellsCollection`'s own slice set.

**Watch:**
- Four `Slice.Enumerator` structs inline is ~288 B of struct. `SlicesEnumerator` must then never be
  copied — pass it by `ref`, and check that `GetCells` and `ForValuesAndFormulas` do not return it
  by value into a `foreach` that copies. **If the copies cannot be eliminated, this task is a
  regression and should be abandoned rather than tuned** — that is a real outcome, and task 4 should
  record it.
- `AdvanceMatchingEnumerators` compacts the array by swapping the last live entry down. With fixed
  fields that becomes a mask or a swap over a small inline buffer; keep the exhausted-enumerator
  semantics identical, because the k-way merge's duplicate-free guarantee (spec 19 area 1) depends
  on advancing *every* enumerator sitting on the current point.

**Acceptance:**
- `L2_PlusWrapper` time improves by more than the documented noise floor, with its 24.00 MB
  allocation figure **unchanged to the byte** — that figure is the `XLCell` wrappers, which this task
  does not touch, so any movement in it means something unintended changed.
- `L6_CellsUsed` improves in step with `L2`.
- `UsedCellEnumerationOrderTests` (9 tests, spec 19 area 1) green unchanged — they characterise
  enumeration order across single/disjoint/overlapping shapes and are the guard on this merge.

---

## Task 3 — `XLRangeParameters` becomes a `readonly record struct`

**Effort:** XS · **Independent of tasks 1 and 2**

`Excel/Ranges/XLRangeParameters.cs:3` is a two-property parameter object (`XLRangeAddress`,
`IXLStyle`), `internal`, never mutated despite its `private set`, and appears only as a constructor
argument at eight sites. Convert to `readonly record struct` and pass `in`.

The reason it earns a task is not its own size. It is `XLRangeBase.GetRange`
(`XLRangeBase.cs:809-810`):

```csharp
var newRangeAddress = new XLRangeAddress(newFirstCellAddress, newLastCellAddress);
var xlRangeParameters = new XLRangeParameters(newRangeAddress, Style);   // <-- Style façade built here
if (newFirstCellAddress.RowNumber < RangeAddress.FirstAddress.RowNumber || ...)
    throw new ArgumentOutOfRangeException(...);                          // <-- bounds check after
return ...GetOrCreateRange(xlRangeParameters);                           // <-- repository lookup after
```

`Style` builds a style façade, which `profile create` prices at **128.1 B / 53.2 ns** for
`ws.Cell(r,c).Style` discarded. So every `range.Range(...)` pays for a parameter object and a style
façade *before* the bounds check runs and *before* the repository is consulted — including on a cache
hit, which is the common case. Move both below the bounds check.

**Watch:** `XLWorksheet.GetOrCreateRange` reads `xlRangeParameters.DefaultStyle != null`. `IXLStyle`
is a reference type so that check survives the struct conversion unchanged, but confirm no call site
relies on passing `null` where a struct default now arrives instead.

**Acceptance:** allocation on a `range.Range(...)`-in-a-loop probe falls by the parameter object and
the façade; full suite green. **If no benchmark in the suite covers this path, add the probe or
state plainly in Results that the change is unmeasured** — spec 19's rule.

---

## Task 4 — Measure, then decide

**Effort:** S · **Depends on:** tasks 1–3 · **Empowered to revert.**

Re-run task 0's command in one sitting against the parent commit and the branch.

**Decision rule, agreed in advance:**

- **Task 1's allocation criterion is binary** — 88 B → 0 B on `L1_SliceOnly` either happens or it
  does not, and it cannot be noise. If it happens, keep task 1 regardless of what the time column
  says: a truly allocation-free `EnumerateUsedCells` matches what its documentation already claims.
- **Task 2 keeps only on a time win above the noise floor**, with `L2_PlusWrapper`'s allocation
  byte-identical. It is the largest and riskiest change here and it buys nothing but dispatch.
  Under the floor, revert it — four inline struct enumerators are a real complexity cost to carry
  for an unmeasurable gain.
- **Task 3 keeps if the probe moves and the suite is green.** It is small enough that
  "no measurable change, but strictly less work on a hot path" is an acceptable verdict; say so
  rather than claiming a win.
- **Any regression** — revert to the last green task and record which one caused it.

**Acceptance:** a Results section in this file stating what was measured, on what commit, and which
tasks survived — in spec 19's format, including the disproved parts.

---

<a id="results"></a>
## Results

Measured 2026-08-08 against parent `66307307`, branch `perf/spec-21-hot-path-structs`.
BenchmarkDotNet 0.15.8, net10.0 Release, `InProcessEmitToolchain`, AMD Ryzen 9 5950X.
Every A/B below is one sitting.

**Outcome: task 3 shipped. Tasks 1 and 2 declined on measurement — task 1 was implemented, measured,
and reverted; task 2 was never written, because task 1's measurement disproved its premise too.**

### Task 1 — the struct conversion is neutral; *embedding* it costs 60%

`UsedCellEnumerationBenchmarks.L1_SliceOnly`, 500,000 cells. Five variants, the same instrument:

| Variant | Mean | Allocated |
|---|---:|---:|
| **baseline** — `Enumerator` is a class, held via `IEnumerator<Point>` | **5,053,270 ns** | 88 B |
| `Enumerator` is a struct, still **boxed** via `IEnumerator<Point>` | 5,068,989 ns | 88 B |
| struct **embedded by value**, reached through a wrapper type | 8,463,473 ns | **0 B** |
| the same, wrapper methods `[MethodImpl(AggressiveInlining)]` | 7,928,567 ns | **0 B** |
| the same, **wrapper type removed entirely** | 8,108,683 ns | **0 B** |

Row 2 is the load-bearing one. **Converting `Slice<TElement>.Enumerator` from a class to a struct is
free — 5,069 µs against a 5,053 µs baseline, inside noise.** Everything that went wrong went wrong
at the next step: embedding those ~72 bytes by value inside `XLUsedCellEnumerator` costs **+60%**,
and rows 4 and 5 show that neither the wrapper layer nor the JIT's inlining decisions explain it.
The wrapper was the first hypothesis and it was wrong; force-inlining it recovered 6% of a 67% hole,
and deleting it recovered nothing.

`EarlyExit_EnumerateFirst`, which is construction-dominated rather than walk-dominated, moved the
other way: 61.30 ns / 88 B → 50.03 ns / **0 B**. That is the shape of the whole finding — the
conversion helps where an enumerator is *created* and hurts where one is *driven*, and real
enumerations are overwhelmingly the second.

**Why (hypothesis, not established).** Boxed, the enumerator's state is one heap object reached
through a pointer the JIT keeps in a register, and the per-row `_columnsEnumerator = …` write is a
heap-field store. Embedded, the enclosing `XLUsedCellEnumerator` becomes an address-taken stack
object whose nested fields cannot be promoted, so every `MoveNext` re-loads them. The interface
dispatch this spec set out to remove was also cheaper than assumed — one implementing type is
observed at each of these sites, so dynamic PGO was already devirtualising it. **The spec's premise
was that the dispatch was the cost. It was not.**

### Task 2 — declined without being written

Task 2 would have embedded **four** of these enumerators by value in `SlicesEnumerator`. Task 1
measured the cost of embedding one at +60% on the walk. The spec pre-committed to this outcome —
*"if the copies cannot be eliminated, this task is a regression and should be abandoned rather than
tuned"* — and that is the call. Writing it to confirm would cost a day to reproduce a result already
in hand.

`L2_PlusWrapper` and `L6_CellsUsed`, which route through `SlicesEnumerator`, were 11,390 µs and
24,986 µs at baseline and are untouched.

### Task 3 — shipped

A new probe, `AllocationBenchmarks.SubRangeOfRange` (1,000 `range.Range(...)` calls over ten reused
addresses, so every call after the first is a repository **cache hit**), added because nothing in the
suite exercised this path:

| | before | after | Δ |
|---|---|---|---:|
| `SubRangeOfRange` | 90.48 µs ± 3.90 / **78.19 KB** | **78.18 µs ± 1.95 / 3 B** | **−13.6% time, −99.996% alloc** |

~80 B per call, all of it the `XLRangeParameters` heap object, on a path that was already going to
find its range in the repository. Error bars do not overlap.

`XLRangeBase.GetRange` also stopped building the parameters and materialising `Style` *before* the
bounds check; that half is unmeasured on its own, because the probe never takes the throwing path.

### The decision rule in task 4 was mis-stated, and is corrected here

> **Task 1's allocation criterion is binary** — 88 B → 0 B on `L1_SliceOnly` either happens or it
> does not, and it cannot be noise. If it happens, keep task 1 regardless of what the time column says.

**That rule is wrong and was not followed.** It priced an 88-byte *per-enumeration* constant as if it
were per-cell, and licensed keeping a 60% time regression to remove it. Spec 19 records
`LoadAndIterateEnumerateUsedCells` at +0.00 MB over load across 3.75 M cells — which is precisely the
evidence that those 88 bytes are invisible at workload scale. A criterion written before the
measurement it governs should not be allowed to override the measurement; this is the fourth time in
this spec family (spec 19 criterion 1, spec 11 criterion 2, spec 05 criterion 3) that a criterion has
priced work its task cannot reach.

**The rule that should have been written:** on an enumeration path, a per-enumeration allocation is
worth removing only if the walk does not get slower. It did, so task 1 goes.

### What is worth taking from a declined task

The reusable finding, which is not what this spec expected to produce:

- **A struct enumerator is not automatically cheaper than a class one.** Where the enumerator is
  large and gets embedded in another struct, the class can win outright — and here it wins by 60%.
- **Dynamic PGO had already devirtualised the interface calls** this spec was written to remove, at
  every one of these single-implementation sites. "It goes through an interface" is not, by itself,
  a cost.
- `Slice<TElement>.Enumerator` staying a class is now a **measured** decision rather than an
  unexamined one.

### Verification

- 11,841 tests pass, 4 skipped, 0 failed, on **both** net8.0 and net10.0.
- No public API change. `XLRangeParameters` is `internal`.
- `PublicAPI.Shipped.txt` / `Unshipped.txt` untouched.

---

## Reviewed and rejected

Seven types that look like candidates and are not. Recorded so they are not re-proposed.

| Type | Why it looks like one | Disqualifier |
|---|---|---|
| `Formula` (`CalcEngine/Formula.cs:6`) | Two readonly refs, immutable, sealed, no interface | Stored in `ConditionalWeakTable<string, Formula>` (`ExpressionCache.cs:14`), whose `TValue` is constrained `: class`. Hard no. |
| `XLStyleValue.TransitionEntry<TKey>` (`Style/XLStyleValue.cs:128`) | 250,151 allocations measured in spec 19 area 2 — the biggest count in the survey | **Deliberate.** Its own comment: hash, key and result travel in one object so a cache slot fills with a single atomic reference write. As a struct, two threads storing into one slot can tear and hand a reader a key and result from different transitions — a silently wrong style, not a miss. |
| `XLCell` (`Cells/XLCell.cs:21`) | Two fields (16 B payload), 48.1 B/instance, the largest per-cell allocation on the read path | Extends `XLStylizedBase` and is vended as `IXLCell`. Spec 19 area 1 task 1.4 already declined the softer version (reusing the wrapper) after pricing it at 24 MB of an 84.66 MB total against the only semantic risk in that area. |
| `XLCellFormula` (`Cells/XLCellFormula.cs:28`) | 1,076 B allocated per formula cell on load — ~50% of load cost (spec 19 area 4) | Mutable and tracked by reference: `XLCalcEngine.SpillFootprint` holds `XLCellFormula owner`, and the dependency tree and calculation chain key on the instance. The lazy `FormulaExtra` split is the right fix for its size and is already done. |
| `Reference` (`CalcEngine/Reference.cs:18`) | Small, immutable, sealed, no interface | Already carries inline-first storage (its comment records removing a ~64 B `List`). It is the payload arm of the `AnyValue` union, which holds it by reference; embedding it would inflate every `AnyValue` on the evaluation stack. |
| `XLBorderValue` / `XLFillValue` / `XLFontValue` / `XLNumberFormatValue` (`Excel/Style/*Value.cs`) | Immutable wrappers over a key | Repository-deduplicated and compared by reference — `XLStylizedBase.ReferenceEqualityComparer<T>`, and `range.StyleValue == StyleValue` at `XLWorksheet.cs:1687`. Value semantics would change what those comparisons mean. The *keys* are the right target; that is spec 20. |
| `XLStyle` (`Style/XLStyle.cs:6`) | `ws.Cell(r,c).Style` discarded costs 128.1 B — a large per-operation figure | Implements `IXLStyle`, is mutable (`Value { get; private set; }`), and holds a container back-reference. A struct would box at every vend. Task 3 attacks the same number from the other side: stop building the façade when nothing reads it. |

`Criteria` (`CalcEngine/Functions/Criteria.cs:12`) was reviewed and is **structurally viable** —
immutable, three fields, no interface, no inheritance — but it is allocated once per
`SUMIF`/`COUNTIF`/`DSUM` call across eight sites, not per cell. Correct and near-worthless. Take it
only if you are already editing the file; do not open a PR for it.

## Explicitly out of scope

**The `XLCell` wrapper and the load-path formula cost.** The two biggest per-cell numbers in the
suite, both already owned: spec 19 area 1 task 1.5 and area 4 respectively.

**Style key sizes.** Spec 20, and its types are already structs — that spec shrinks them rather than
converting them.

**Making `ISlice.GetEnumerator` generic.** It would let every slice vend a struct enumerator without
the concrete-accessor duplication task 1 adds, but it changes an interface implemented by four slice
types and consumed across the calc engine, IO and range layers. If tasks 1 and 2 land and the numbers
justify going further, that is a separate spec.

## Ground rules

Standard for this repo (see `docs/specs/README.md`): branch per task, never commit to main, warnings
are errors, nullable annotations required, no compound shell commands. Perf PRs carry before/after
numbers. Line numbers above were verified against `66307307` on 2026-08-08 — re-verify before editing.
