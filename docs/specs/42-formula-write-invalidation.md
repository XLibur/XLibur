# Spec 42 — One formula write path, and it invalidates

**Area:** Architecture · **Defect (stale reads after an array-formula edit)**
**Effort:** S–M (~3 days)
**Dependencies:** None hard. Shares the calc-engine staleness model with specs 40 and 43; one owner
should take all three, in the order 40 → 42 → 43.
**Status:** Proposed. From the 2026-08-30 architecture review (round 3).

## Problem Statement

Writing a formula into a cell has two halves: store it, and tell the calc engine that everything
depending on that cell is now out of date. The single-cell path does both. The array-formula path
does the first and only part of the second — it marks the formula itself, and never reaches the
engine's dependency walk.

What a user sees:

```
A1..A3 = 1,2,3 · D1:D3 {=A1:A3*2} · F1 "=SUM(D1:D3)"    ->  D = 2,4,6   F1 = 12
replace D1:D3 with {=A1:A3*10}
                                                          ->  D = 10,20,30   F1 = 12
```

`F1` should be 60. The array's own cells update; everything that depends on them does not. The same
happens the other way round: read a formula that references a region, *then* put an array formula in
that region, and the reader keeps its old answer.

The equivalent single-cell edit is correct, which makes this hard to spot — the behaviour a developer
would test first is the behaviour that works.

## Solution

The module that stores formulas is the one that invalidates dependents. Callers hand it a formula and
a location; they do not also have to remember to notify the engine afterwards, and they cannot get the
order wrong, because there is no order left to get wrong.

Loading a file is the one case that legitimately skips invalidation — a freshly loaded workbook has
nothing stale. That becomes an explicit, named mode rather than a caller-side omission that looks
identical to a bug.

## User Stories

1. As a library consumer, I want a formula that sums an array formula's range to update when I replace
   that array formula, so that my totals are correct.
2. As a library consumer, I want the same when I first write an array formula into a region something
   already references, so that the order in which I build a sheet does not matter.
3. As a library consumer, I want removing an array formula to update its dependents, so that deletion
   is as safe as replacement.
4. As a library consumer, I want an array formula that depends on another array formula to update, so
   that chains of array formulas work.
5. As a library consumer, I want a cell formula and an array formula to behave identically with respect
   to invalidation, so that I do not have to know which kind I used.
6. As a library consumer, I want a formula written through a range operation to invalidate its
   dependents, so that bulk authoring is as correct as cell-by-cell authoring.
7. As a library consumer, I want a formula written during a copy or paste to invalidate its dependents,
   so that duplicated regions are current.
8. As a library consumer, I want a formula rewritten by a row or column insertion to invalidate its
   dependents, so that structural edits leave a consistent sheet.
9. As a library consumer, I want loading a workbook not to trigger a full recalculation, so that
   opening a file stays fast.
10. As a library consumer, I want a loaded workbook's cached values to be trusted until something is
    edited, so that load performance is unchanged by this fix.
11. As a library consumer, I want to be able to read a dependent value immediately after writing an
    array formula, without calling a full recalculation first, so that the object model behaves the way
    the single-cell path already does.
12. As an XLibur maintainer, I want invalidation to sit next to the write, so that the two cannot be
    separated by a caller.
13. As an XLibur maintainer, I want a new formula write path to inherit invalidation, so that the next
    one cannot forget it.
14. As an XLibur maintainer, I want the storage module's interface to stop exposing an ordering rule
    between its own methods, so that callers have nothing to sequence.
15. As an XLibur maintainer, I want the load-time skip to be explicit and named, so that it reads as a
    decision rather than an oversight.
16. As an XLibur maintainer, I want tests that assert on a *consumer* of an array formula, so that this
    defect class is pinned rather than only the array's own cells.
17. As a contributing agent, I want one way to write a formula, so that I do not have to discover which
    of several methods also notifies the engine.

## Implementation Decisions

**The seam is the existing formula storage module.** It already registers formulas in the dependency
tree and the calculation chain; adding invalidation to it puts all three effects of a write in one
place. No new type.

**Its interface shrinks.** It currently offers separate operations for setting a single formula,
setting an array formula, setting during load, and marking dirty — and leaves callers to combine them
correctly. After this spec, marking dirty is not a caller-facing operation at all, and the load-time
variant is the only way to opt out.

**Load remains the sole opt-out.** Loading writes formulas for a workbook that by definition has no
stale dependents, and the load path must not pay for a graph walk per formula. That mode stays, named
for what it is.

**The single-cell path is the reference behaviour.** It already does the right thing. This spec moves
the call rather than inventing a new policy, so the risk is concentrated in *where* the call happens,
not in *what* it does.

**Performance is a gate.** Formula writing is on the create path and the structural-edit path. The
change must not regress the create-path or structural-edit benchmarks; measure before and after. If
the array path's newly-added invalidation proves expensive for large arrays, the fix is to mark the
array's whole footprint once rather than per cell — not to skip it.

**Interaction with spec 40.** That spec fixes what happens once the dependency walk is entered; this
one fixes a path that never enters it. They are independent and can land in either order, but 40 first
means this spec's tests exercise a correct walk.

**No public API change.**

## Testing Decisions

**What makes a good test here.** A good test writes formulas through the public object model, reads a
*dependent* cell, and asserts its value — without calling a full recalculation in between. The moment
a test inserts a full recalculation, it stops being able to see this defect, which is precisely what
the current suite does.

**The centrepiece is a write-kind by operation matrix.** For each kind of formula write — single cell,
array formula, formula set through a range, formula produced by a copy, formula rewritten by a
structural edit — and for each operation — first write, replacement, removal — assert that a
downstream consumer of the written region is current afterwards.

**Consumer-side assertions, not producer-side.** The existing array formula tests assert on the array's
own cells and on expected exceptions. Both are correct and stay. What is missing, in all
twenty-five of them, is any assertion about a cell that *reads* the array's range. That is the
assertion this spec adds.

**A no-recalculation rule for the new tests.** They must not call a full recalculation. If a test needs
one to pass, it is not testing this.

**A load-path guard test.** Load a workbook with many formulas and assert that loading does not trigger
evaluation — otherwise the opt-out could be silently lost later and only show up as a performance
regression.

**Prior art.** The array formula tests are the right home and the right shape; they need consumer
assertions. The single-cell equivalents already demonstrate the correct behaviour and make good
control cases — every new test should have one.

**Test seam.** `IXLCell.FormulaA1`, `IXLRange.FormulaArrayA1`, and reading `.Value`. No new seam.

## Out of Scope

- What the dependency walk does once entered — spec 40.
- Reading a spilled cell's value — spec 43.
- Demand-driven evaluation — spec 04.
- The dependency tree's construction and what gets registered in it.
- Changing when a workbook recalculates as a whole.

## Further Notes

The asymmetry here is one missing call, which makes it sound trivial. It is worth being precise about
why it happened: the single-cell path assembles the sequence in the cell type, and the array path
assembles a different sequence in the range type. Neither is wrong on its own terms; there is simply
no module that says what the sequence *is*. The fix is not to add the missing call to the second
caller — that leaves a third caller free to make the same mistake — but to remove the possibility of a
caller assembling the sequence at all.

Both stale reads were reproduced against a scratch build, along with a single-cell control that
behaves correctly, before this spec was written.
