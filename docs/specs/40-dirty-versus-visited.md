# Spec 40 — The dirty flag stops doubling as a visited marker

**Area:** Architecture · **Defect (propagation pruned)**
**Effort:** S–M (~3 days)
**Dependencies:** None hard. Shares the calc-engine staleness model with specs 42 and 43; one owner
should take all three, in the order 40 → 42 → 43.
**Status:** ✅ Merged — [#418](https://github.com/XLibur/XLibur/pull/418) (squash `2c7150c7`, 2026-08-31). From the 2026-08-30 architecture review (round 3). See Results.

## Problem Statement

When a cell changes, the calc engine walks the dependency graph marking everything downstream as
needing recalculation. That walk uses a formula's *dirty* flag as its "have I already visited this
node" marker.

Those are two different facts. A formula can already be dirty for reasons that have nothing to do
with the current walk — the public invalidation method on a cell sets it, and so do sheet renames,
reference shifts and range moves. When the walk reaches such a node it concludes it has already been
there, stops, and never visits anything downstream of it.

What a user sees:

```
A1=1 · B1 "=A1+1" · C1 "=B1+1" · D1 "=C1+1"     ->  2, 3, 4
ws.Cell("C1").InvalidateFormula();               // documented public API
ws.Cell("A1").Value = 10;
   B1 = 11  correct    C1 = 12  correct    D1 = 4   <-- should be 13
```

Calling a public method whose entire purpose is to force recalculation causes a *later* edit to
under-recalculate. The deeper the graph beyond the invalidated cell, the more cells go stale. Nothing
reports it; the values are simply old.

A second, smaller problem sits alongside it: the code documents a three-state epoch model for
tracking staleness, and the method that would advance the epoch is never called anywhere. The epoch
is a constant, so the documented model does not exist and the flag is really a boolean.

## Problem the user actually hits

Any application that invalidates a cell — to force a refresh, after a bulk edit, or as part of its own
caching — silently degrades the correctness of every subsequent recalculation in that region of the
sheet.

## Solution

The traversal keeps its own record of what it has visited, separate from whether a formula is stale.
Marking dirty means marking dirty; visiting means visiting. A node that was already dirty is still
traversed, so everything downstream of it is reached.

The epoch machinery is resolved at the same time: either it is wired up and does what it documents, or
it is removed and the flag is honestly a boolean. Leaving code that documents a model it does not
implement is the thing that made this defect hard to see.

## User Stories

1. As a library consumer, I want editing a cell to recalculate everything that depends on it, however
   deep the chain, so that the values I read are current.
2. As a library consumer, I want calling `InvalidateFormula` on a cell to have no effect on how later
   edits propagate, so that a refresh hint does not corrupt subsequent calculations.
3. As a library consumer, I want invalidating an intermediate cell in a chain to leave the rest of the
   chain reachable, so that partial refreshes are safe.
4. As a library consumer, I want a sheet rename not to reduce how far a later edit propagates, so that
   renaming is not silently destructive.
5. As a library consumer, I want a row or column insertion, which shifts references, not to prune a
   later recalculation, so that structural edits and value edits compose.
6. As a library consumer, I want a moved range not to prune propagation, for the same reason.
7. As a library consumer, I want a cyclic dependency to still terminate the walk, so that fixing this
   does not reintroduce a hang.
8. As a library consumer, I want a diamond-shaped dependency graph to recalculate every path, so that a
   node reachable by two routes is not skipped on the second.
9. As a library consumer, I want propagation across worksheets to behave the same as within one, so
   that sheet boundaries are not a special case.
10. As a library consumer, I want the recalculation cost of an edit to stay proportional to what
    actually depends on it, so that correcting this does not make edits slow.
11. As an XLibur maintainer, I want cycle termination to be a property of the traversal, so that no
    other subsystem can affect it by setting a flag.
12. As an XLibur maintainer, I want the dirty flag to mean exactly one thing, so that reading the code
    does not require knowing which meaning is in play.
13. As an XLibur maintainer, I want the epoch model either implemented or deleted, so that the comments
    and the code agree.
14. As an XLibur maintainer, I want a test that dirties an intermediate node and then edits the root, so
    that this defect class is pinned.
15. As an XLibur maintainer, I want the public invalidation method to have tests at all, so that a
    shipped public method is not entirely uncovered.
16. As a contributing agent, I want the difference between "stale" and "visited" to be visible in the
    types, so that I do not reuse one for the other.

## Implementation Decisions

**The seam is the existing dependency-tree walk.** Nothing new is introduced. The walk gains its own
visited set — or a generation counter compared per node, if measurement shows the allocation matters —
and stops consulting the formula's dirty flag for that purpose.

**Marking stays idempotent.** A node already marked dirty by this walk is not re-enqueued; a node
marked dirty by anything *else* is. That distinction is the whole fix.

**Cycle termination is preserved and tested explicitly.** The dirty flag currently provides
termination as a side effect. The replacement must provide it deliberately, and the test for it comes
before the change, not after.

**The epoch decision is made on evidence.** The method that would advance the edit epoch has no
callers, so the epoch is constant and the two states the code distinguishes are identical. The spec's
default is to delete the machinery and document the flag as boolean, because that is what the code
does. If implementing the epoch turns out to be cheap and to buy something measurable, that is a
finding to report rather than a silent substitution.

**No public API change.** The observable change is that more cells recalculate — which is to say, the
right ones.

**Performance is a gate, not a goal.** The walk is on the edit path. The change must not regress the
structural-edit or bulk-edit benchmarks; a measurement before and after is part of the work.

## Testing Decisions

**What makes a good test here.** A good test builds a dependency chain through public API, performs
public operations, and asserts cell values. It does not inspect the dirty flag, the visited set or the
dependency tree — those are the implementation. The defect is fully observable as wrong values.

**The centrepiece is the interference matrix.** For each way a formula can become dirty outside a walk
— the public invalidation method, a sheet rename, a row or column insert, a range move — dirty a node
partway along a chain, then edit the chain's root, then assert every downstream value. That is the
test the current shape cannot pass and, once passing, the one that keeps it fixed.

**Graph-shape tests.** A straight chain, a diamond, a cross-sheet chain, and a cycle. The cycle test
asserts termination and must be written *before* the change, so that it is known to be meaningful.

**A control case in every test.** The same graph and the same edit without the interfering operation,
asserting the correct values. Without the control, a test that passes proves nothing about whether the
interference was the cause.

**Prior art.** The calc engine's recalculation tests are the right home. Note that the public
invalidation method currently has no tests anywhere in the suite, so this spec creates its first
coverage rather than extending existing coverage.

**Test seam.** `IXLCell` value and formula assignment, `InvalidateFormula`, and worksheet-level
structural operations. No new seam.

## Out of Scope

- Which cells are *registered* in the dependency tree, and how the tree is built. Only the walk over
  it changes.
- The formula write path's failure to invalidate at all — that is spec 42, and the two are
  independent: 42 adds a missing call, 40 fixes what happens once it is made.
- Spill-aware reads — spec 43.
- Demand-driven evaluation, which is spec 04.
- Making recalculation faster. This spec makes it correct and must not make it slower.

## Further Notes

Reusing one bit for two meanings is a classic and usually harmless economy. It is harmless when the
bit has exactly one writer. Here it has at least four, and three of them are in unrelated subsystems
that have no reason to know the dependency walk exists.

The dead epoch method is worth treating as evidence rather than tidying. Code that documents a
three-state model and implements one state is a sign that the model was designed, partly built, and
then left — and the comment is now actively misleading anyone who reads the flag and believes it.

Both wrong answers above were reproduced against a scratch build before this spec was written.

## Results

**Implemented 2026-08-30 on `task/40` (worktree `xl-wt-40`), head `9a070669`, 9 commits, cut from
`upstream/main` `37c986bb`; PR [#418](https://github.com/XLibur/XLibur/pull/418) opened 2026-08-31 at head `edc3b88c` (10 commits, docs sync). The PR body also points at `docs/dirty-walk-bulk-edit.md` for the follow-ups.** Full suite green on both TFMs: 14,265 per
TFM, 0 failed, 5 pre-existing skips; Report 481, Fonts.SixLabors 31, Fonts.SkiaSharp 37 all green.

**The commit order the brief demanded held**: cycle-termination safety net green first (`6fbab32a`),
interference matrix red (`8c610852`), graph shapes red (`50595721`), the fix (`576e76da`), the epoch
deletion on evidence (`f7244615`), then two perf commits and a review-driven fix.

**The fix.** `MarkDirty` carries a process-wide walk id and `XLCellFormula.TryVisit` compares a
per-node generation against it; dirty state is no longer consulted for traversal. The spec's default
(a `HashSet` visited set) was measured first — ~7× allocation, ~3× time — and the generation counter
replaced it on that evidence. The epoch machinery: `BumpEditEpoch` had zero callers, `EditEpoch` was
pinned at 1, and `_evalEpoch` was a two-state bool in disguise; deleted and replaced by
`bool _isClean` across 8 call sites in 5 files.

**What the spec predicted that turned out wrong.** Two of the four named dirtiers — sheet rename and
row/column insert — **cannot reproduce the defect through their public operations**: both trigger
`XLCalcEngine.Purge`, which marks the whole workbook dirty and masks the narrow bug. Their live-API
tests are green before and after and are kept as regression cover; the primitive both actually use
(a formula text rewrite) is isolated in `Formula_marked_dirty_by_a_text_rewrite_does_not_prune_a_later_edit`,
which is genuinely red/green. `InvalidateFormula` and range move reproduce exactly as the spec says.

**The in-branch review found a High-severity regression the fix had introduced**: walk ids were reused
across a dependency-tree rebuild, so a graph built after a purge could be pruned. The interference
matrix could not see it — every test built its graph after the last purge. Fixed in `b08faf00` with
`Walk_ids_are_not_reused_after_the_dependency_tree_is_rebuilt`. Same lesson as every round: run the
review inside the task.

**The performance gate — one shape regressed, deliberately.** Three-run medians:

| Probe | main | HEAD |
|---|---|---|
| 200k edits, 200×10 chains | 821.5 MB / 1588 ms | 791.0 MB / 1848 ms |
| 200k edits, no dependents | 46.5 MB / 41 ms | 16.0 MB / 46 ms |
| **20k unsettled edits, 50-deep model** | **12.6 MB / 48 ms** | **221.3 MB / 216 ms** |

The third row is real and was independently reproduced by the reviewer. Reusing the dirty flag also
short-circuited *between* calls — bulk writes into a shared, still-dirty model skipped the walk after
the first edit. That saving was never sound; it is the same shortcut that pruned legitimate walks.
The allocation is RBush search lists, one per hop, not the queue. **The agent chose correctness and
documented the cost in the changelog rather than restore the skip; it flagged this as a product call
for a second opinion — open as of 2026-08-31.** Restoring a *sound* cross-call skip needs a third
state — a `_dirtiedByWalk` bit distinct from `_isClean` — and `BulkEditDirtyWalkProfile`
(`profile bulkedit`, permanent) is the guard for whoever attempts it. Structural-edit benchmarks are
within noise, as expected — `Purge` never calls the changed method.

**Found, not fixed — recorded as D25 and D26.** `XLRange.TransposeRange` swaps its offsets on the
wrong axes (silent data misplacement). `XLCell.InvalidateFormula` marks only its own formula and runs
no walk, so dependents stay stale until an unrelated edit sweeps them — a missing call, spec 42's
territory, outside "only the walk changes".

**Least sure of (agent's own list).** The matrix is organised by dirtier and the regression was in the
tree's lifecycle — other lifecycle axes may be unenumerated. Whether the unsettled bulk-edit cost is
shippable. `WalkQueueRetainedCapacity = 4096` is a judgement, not a measurement.

**What the next consumer inherits.** Spec 42 adds the missing invalidation on the array-formula write
path and now has a walk that will actually propagate it; it should also take D26. Spec 43 builds on
both. `Mark_dirty_stops_at_dirty_cell` was rewritten — it had asserted the defect's symptom as
intended behaviour.

**Merged 2026-08-31** as [#418](https://github.com/XLibur/XLibur/pull/418) (squash `2c7150c7`,
branch tip `edc3b88c`), the first of the five. **The open performance decision was resolved by
merging as-is**: the merged CHANGELOG documents the cost in user terms (20k unsettled edits of a
50-deep shared model ~50 ms → ~190 ms; reading or saving between edits unaffected, allocations
otherwise lower). A sound cross-call skip (`_dirtiedByWalk` third state, guarded by
`BulkEditDirtyWalkProfile`) remains an un-specced follow-up. Specs 42 and 43 build on the merged
walk.
