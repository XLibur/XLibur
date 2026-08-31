# The dirty walk and the bulk-edit cost

Notes on the performance side of spec 40 ("the dirty flag stops doubling as a visited
marker"), written for review. The correctness story is in the CHANGELOG entry and the spec;
this document covers only what the change did to bulk-edit cost, and why the expensive shape
is expensive on purpose.

Branch: `task/40`. Profile: `XLibur.Benchmarks`, `profile bulkedit`
(`XLibur.Benchmarks/BulkEditDirtyWalkProfile.cs`).

## The mechanism that changed

`DependencyTree.MarkDirty` walks the dependency graph breadth-first from an edited area and
marks every dependent formula dirty. It needs some "already visited" marker, or a cyclic
graph loops forever.

It used to use the formula's own dirty flag as that marker: if a dependent was already dirty,
the walk skipped it and everything beneath it. Spec 40 replaced that with a walk id stamped
on each formula it enqueues, so the walk tracks its own visits and a formula's dirty state no
longer influences traversal.

That fixed the correctness bug, and it also removed a saving that nobody had written down.

## Before: the walk short-circuited *between* edits, not just within one

The dirty flag is not reset until a formula is evaluated. So a formula left dirty by one
`MarkDirty` call was still dirty when the next call arrived, and the next walk skipped it —
along with everything downstream.

For a sequence of edits with no read in between, that meant:

```
edit 1  ->  walk marks the whole downstream model dirty
edit 2  ->  first dependent is already dirty -> stop after one hop
edit 3  ->  stop after one hop
...
```

Roughly `O(edits + model)` for the whole sequence, because only the first edit paid for the
traversal.

This was never sound. It is the same shortcut that produced the bug spec 40 exists to fix: a
formula dirty for an unrelated reason (a prior `InvalidateFormula`, a rename's reference
rewrite, a shift, a range move) was indistinguishable from one this walk had already visited,
so a legitimate walk stopped at it and left real dependents stale. The cheapness and the bug
were the same behaviour seen from two sides.

## After: every edit walks the closure it actually affects

Each `MarkDirty` gets a fresh walk id, so nothing carries over between calls. The same
sequence is now `O(edits × closure)`.

There is no cheap way to get the saving back correctly. A cross-call skip needs a marker that
distinguishes *"this subtree is dirty because a walk reached it"* from *"dirty for some
unrelated reason"* — a third state, e.g. a `_dirtiedByWalk` bit separate from `_isClean` and
cleared by `MarkClean()`. The dirty flag alone cannot carry that distinction, which is the
whole finding of the spec. Reintroducing a skip on the dirty flag would reintroduce the bug.

## Measurements

Three probes, same machine, back to back. **Allocation figures are exact and were
byte-identical across repeated runs.** Times are single-shot Stopwatch numbers on a machine
with roughly 40% run-to-run variance — treat them as order-of-magnitude, not as a benchmark
result. Times below are medians of three runs.

| Probe | `main` | fix, queue per walk | fix + pooled queue (HEAD) |
|---|---|---|---|
| 200k edits, 200×10 chains (real walk) | 821.5 MB / 1588 ms | 821.5 MB / 1531 ms | **791.0 MB** / 1848 ms |
| 200k edits, no dependents (walk finds nothing) | 46.5 MB / 41 ms | 46.5 MB / 37 ms | **16.0 MB** / 46 ms |
| 20k unsettled edits, 50-deep shared model | 12.6 MB / 48 ms | 224.4 MB / 184 ms | **221.3 MB** / 216 ms |

Independently reproduced during code review on a throwaway worktree of `upstream/main` with
the benchmark file copied across: 49 ms / 12.6 MB on `main` versus 191 ms / 221.3 MB on HEAD
for the third probe. That agreement is the main reason to trust these numbers despite the
timing noise.

### Reading the table

- **Row 1 and row 2 are the normal cases and did not regress.** Anything that reads or saves
  between edits, and any edit to a cell nothing depends on, settles the graph and never hits
  the cross-call skip. Row 2 actually improves by ~3x on allocation versus `main` — see
  pooling below.
- **Row 3 is the whole cost.** 17.6x allocation, ~4x time. This is the shape that used to be
  free and now is not.
- The row 3 allocation is **not** the walk queue. It is one RBush `FindDependentsAreas` search
  list per hop: 20,000 edits × 50 hops = 1M searches instead of 20k. It scales linearly with
  model depth, so a model deeper than 50 costs proportionally more.

### The pooled queue

`MarkDirty` allocated a fresh `Queue<SheetArea>` per call and regrew it as the walk went.
Reusing one queue per tree (the walk is not re-entrant) removes that. It is a small win where
the walk has real work to do and a large one where it does not — row 2 is the common case
when bulk-writing plain data, and it drops to about a third of what `main` allocated.

The queue is cleared on the way out rather than the way in, so a walk that threw does not
leave its entries reachable, and its backing array is released when it grows past
`WalkQueueRetainedCapacity` (4096) so one unusually wide walk does not retain a slot per
visited node for the tree's lifetime.

## The open question

**Is row 3 acceptable to ship?** "Fill an input range beneath a model, then read or save once"
is a real and common pattern, and 17.6x allocation on it is not nothing. The judgement made
on the branch was: take correctness now, measure and document the cost rather than gamble on
a soundness argument for a cross-call skip under time pressure. That is a product call, made
autonomously, and it is the one most worth a second opinion.

If it needs to be recovered, the `_dirtiedByWalk` third state described above is the shape of
the answer, and the third probe in `BulkEditDirtyWalkProfile` is the guard that any such
change has to move.

---

# Follow-ups

Two items from the branch's "least sure of" list, recorded here so they are not lost.

## 1. The interference matrix was organised by the wrong axis

**What happened.** Spec 40's test plan is a matrix over *dirtiers* — `InvalidateFormula`,
sheet rename, row/column insert, range move — each with a control. That matrix was built,
went red before the fix and green after, and missed a regression the fix itself introduced.

The regression: the walk-id counter initially lived on `DependencyTree`. But
`XLCalcEngine.Purge` discards and rebuilds the whole tree on a sheet add or rename and on
every row/column insert or delete, while the `XLCellFormula` objects carrying the walk stamps
survive. A per-tree counter restarted at zero and handed surviving formulas an id they were
already stamped with, so the first walk after every rebuild pruned at its first hop — the
original defect, in a narrower window. Code review caught it; the matrix could not, because
**every test in it builds its graph after the last purge**.

Fixed in `b08faf00` (process-wide counter) with
`Walk_ids_are_not_reused_after_the_dependency_tree_is_rebuilt` as cover.

**Why it matters beyond this one bug.** The matrix enumerates *what dirties a formula*. The
defect lived in *the lifetime of the structure doing the walking*. Those are different axes,
and only one was enumerated. Worth asking, before the next spec in this area:

- What else in the calc engine has state that outlives the object that issues it? The walk
  stamp on `XLCellFormula` versus the tree is one instance; the calculation chain and the
  spill footprints are worth the same question.
- Which tests would survive a `Purge` in the middle of them? Almost none currently exercise
  the rebuild boundary at all.
- `Purge` is blunt enough to mask narrow bugs (it already masks two of the four dirtiers in
  the matrix — see the note in `DirtyPropagationTests.cs`). Anything it masks is invisible to
  public-API tests and needs a white-box test against the primitive.

**Suggested action.** A small second matrix over *lifecycle events* (purge/rebuild, sheet
add, sheet delete, workbook load) crossed with "does a stale marker survive this?", rather
than more dirtiers.

## 2. `WalkQueueRetainedCapacity` is a guess, not a measurement

**What it is.** `DependencyTree.WalkQueueRetainedCapacity = 4096`. Above this, `MarkDirty`
trims the pooled queue's backing array on exit instead of retaining it.

**What is actually known.** Only that it is high enough not to fire on any of the three
profile probes — verified, since allocations are byte-identical with and without the trim. So
it is not currently costing anything.

**What is not known.** Where real workbooks' dependency closures actually sit. The constant
was chosen to sit "above any ordinary closure" on the strength of the existing comment in
`MarkDirty` that the longest chain seen in the wild is ~1000 formulas — but a *closure* is not
a *chain*, and a wide fan-out (one input feeding thousands of formulas) reaches 4096 far more
easily than a deep one does. Both failure modes are mild:

- Too high: a wide walk retains up to 4096 `SheetArea` slots plus a sheet-name reference each
  for the tree's lifetime. Bounded and small.
- Too low: frequent trims on a workload that repeatedly walks a large closure, giving back
  the pooling win.

**Suggested action.** Cheap to resolve — instrument peak queue depth across the existing
benchmark corpus and the round-trip fidelity fixtures, then set the constant from the
distribution rather than from a comment about chain length. Low priority; it is a tuning
constant with no correctness role.
