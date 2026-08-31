# Spec 43 — Reading a cell is spill-aware, once

**Area:** Architecture · **Defect (order-dependent value)**
**Effort:** M (~4–5 days)
**Dependencies:** Shares the calc-engine staleness model with specs 40 and 42; one owner should take
all three, in the order 40 → 42 → 43.
**Status:** Proposed. From the 2026-08-30 architecture review (round 3).

## Problem Statement

A dynamic-array formula occupies one anchor cell and spills its result into neighbouring cells. Those
neighbours hold values but no formula of their own.

Two different pieces of code answer the question "what is this cell's current value". The one the calc
engine uses when evaluating a formula knows about spilling: for a cell with no formula it checks
whether the cell belongs to a spill whose anchor is stale, and forces the anchor to evaluate first.
The one the public interface uses does not: it recomputes only if the cell has a formula of its own,
and a spilled cell has none. So it returns whatever was last written there.

What a user sees:

```
A1:A4 = 1,2,3,4 · C1 {=UNIQUE(A1:A4)} spilling into C1:C4 · E1 "=C4&\"\""
edit A4 so the result shrinks to three rows
   [1] read C4 directly   -> '4'     stale
   [2] read E1            -> ''      correct — the formula read forces the anchor
   [3] read C4 again      -> ''      the same call as [1], now different
```

Reading a cell gives a different answer depending on what was read before it. That is worse than a
consistently stale value, because caching or reordering in the caller's code changes the result.

A third copy of the same knowledge sits alongside: the engine keeps its own list of which cells belong
to which spill anchor, maintained by hand across three methods, and duplicating information the
formula's own recorded range already implies. Clearing that list drops half the answer.

## Solution

There is one module that resolves a cell location to its current value. It knows the three things
that matter — whether the location has a formula, whether it belongs to a spill, and whether what it
depends on is stale — and it is the only thing that knows them.

Its two callers differ in one respect only: what to do when the value is stale. The public read
recomputes and returns; the in-formula read signals the engine to evaluate the anchor first and come
back. That single difference is the whole of the distinction between them.

Spill ownership becomes something that module derives rather than a list two other types maintain.

## User Stories

1. As a library consumer, I want reading a spilled cell to give the same answer regardless of what I
   read before it, so that my code's behaviour does not depend on access order.
2. As a library consumer, I want a spilled cell to be current after I edit something its anchor depends
   on, so that I do not have to force a full recalculation.
3. As a library consumer, I want a spill that shrinks to clear the cells it no longer occupies, so that
   stale values do not remain visible.
4. As a library consumer, I want a spill that grows to fill the cells it now occupies, so that the new
   region is populated.
5. As a library consumer, I want `NeedsRecalculation` on a spilled cell to tell me the truth, so that I
   can decide whether to recalculate.
6. As a library consumer, I want the cached-value accessor on a spilled cell to be consistent with the
   value accessor, so that the two do not disagree.
7. As a library consumer, I want a saved workbook to contain the current spilled values, so that the
   file matches what the object model reports.
8. As a library consumer, I want enumerating a range that overlaps a spill to yield current values, so
   that bulk reads agree with individual reads.
9. As a library consumer, I want a spill that becomes blocked to report the blocking error, so that I
   can detect the condition.
10. As a library consumer, I want reading a spilled cell to be no slower in the common case where
    nothing is stale, so that correctness here does not cost throughput.
11. As a library consumer, I want an anchor cell to keep behaving exactly as it does today, so that the
    case that currently works is not disturbed.
12. As an XLibur maintainer, I want one implementation of "is this value current", so that the public
    and internal readers cannot disagree.
13. As an XLibur maintainer, I want spill ownership derived from the formula's recorded range, so that a
    separate list cannot fall out of sync or be cleared independently.
14. As an XLibur maintainer, I want the spill test suite to assert on cell values directly, so that it
    stops routing around the defect with a full recalculation.
15. As an XLibur maintainer, I want the two staleness policies stated side by side, so that the
    difference between the public and in-formula reads is a visible decision.
16. As a contributing agent, I want one place to ask what a cell's value is, so that I do not pick the
    naive reader by accident.

## Implementation Decisions

**The seam is a single value resolver.** Both existing readers become callers. The resolver answers
"the current value at this location", and takes the caller's staleness policy as the one thing that
varies — recompute, or signal to reorder evaluation.

**Spill ownership is derived, not stored.** The engine's separate ownership list duplicates what a
formula's own recorded range already says. Deriving it removes a third copy and removes the failure
mode where clearing the list silently disables spill awareness for the public reader. If measurement
shows the derivation is too slow on the read path, a cache is acceptable — but as a cache owned by the
resolver, invalidated by it, not as a second source of truth maintained by other types.

**The in-formula reader's behaviour is the reference.** It is the one that is correct. The public
reader adopts its knowledge and differs only in its response to staleness.

**The naive checks move behind the resolver.** The public "does this need recalculating" predicate and
the cached-value accessor currently ask whether the cell has a formula. They ask the resolver instead,
which is what makes them correct for spilled cells.

**Performance is a gate.** Cell reading is the hottest path in the library — the read-heavy benchmark
is dominated by it. The resolver must add nothing measurable when no spill is present, and the
fast-path check for "this workbook has no spills at all" must be preserved. Measure before and after;
a regression here is a blocker, not a trade.

**The adjacent duplication is recorded, not fixed here.** The engine's fast single-cell evaluation
path and its general formula-application path share a substantial literal duplicate and disagree about
data tables — the fast path defers to a general path that throws. This was not reproduced. It is noted
as a lead for whoever takes this spec, to be confirmed or dismissed with evidence, not fixed blind.

**No public API change.** The observable change is that spilled cells report current values.

## Testing Decisions

**What makes a good test here.** A good test edits a precedent and then reads a spilled cell through
the public interface, asserting its value — with no full recalculation in between. Every existing
mutating spill test inserts one, which is exactly why none of them can see this defect. The new tests
must not.

**The centrepiece is the read-order test.** Edit a precedent, then read the spilled cell *first*,
before anything else touches the engine, and assert it is current. Then repeat the sequence reading a
dependent formula first. Both orders must give the same answer. That equality is the property the
defect violates.

**Footprint-change tests.** A spill that shrinks, one that grows, one that becomes blocked, and one
that becomes unblocked — asserting both the newly occupied and the newly vacated cells.

**Consistency tests across accessors.** For the same spilled cell, the value accessor, the cached-value
accessor and the needs-recalculation predicate must agree.

**A save-path test.** Edit a precedent, save without recalculating, reopen, and assert the file
contains the current spilled values.

**A performance guard.** The read-heavy benchmark before and after, recorded in the spec's results.

**Prior art.** The spill evaluation tests are the right home. They are well constructed and their use
of a full recalculation is deliberate — it makes them deterministic. The new tests are a different
category and should be grouped separately so that nobody later "fixes" them by adding the recalculation
back.

**Test seam.** `IXLCell.Value`, `CachedValue`, `NeedsRecalculation`, and a save/reload round trip. No
new seam.

## Out of Scope

- The dependency walk's pruning behaviour — spec 40.
- Formula writes that never invalidate — spec 42.
- Demand-driven evaluation — spec 04.
- Spill phase B work, which has its own note.
- The data-table disagreement between the two application paths, recorded above as an unverified lead.
- Changing spill semantics. This spec makes reads report what the engine already computes.

## Further Notes

An order-dependent read is a distinctive failure. A consistently stale value is a bug a caller can
work around once they know about it; a value that changes depending on what was read first cannot be
worked around at all, because the workaround depends on knowing the whole program's read order.

The suite's use of a full recalculation between mutation and assertion is worth calling out without
criticism — it is the obvious way to make a spill test deterministic, and it was almost certainly
added for that reason. It is also, exactly, the thing that hides this. That pattern is worth looking
for elsewhere: a test helper that makes tests reliable by forcing the system into a known state also
stops the tests from noticing when the system does not reach that state on its own.

The order-dependent read was reproduced against a scratch build before this spec was written.
