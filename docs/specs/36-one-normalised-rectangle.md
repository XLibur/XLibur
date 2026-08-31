# Spec 36 — One rectangle, normalised once

**Area:** Architecture · **Defect (5 shipped, 1 fatal)**
**Effort:** M (~5–6 days)
**Dependencies:** None hard. **Should land before spec 51** — the consolidation engines' only live
divergence is the one this spec removes.
**Status:** 🟩 Implemented on `task/36` (2026-08-30), unmerged — see Results. From the 2026-08-30 architecture review (round 3).

## Problem Statement

A user can name a range by its corners in either order. `ws.Range("B5:E2")` and
`ws.Range(ws.Cell("C3"), ws.Cell("A1"))` are both accepted, and Excel accepts the same thing. What
XLibur does with such a range then depends on which part of the library it reaches.

Five failures, all shipped, all reachable from ordinary public API:

1. **The workbook cannot be saved.** Adding a conditional format to a range whose rows are reversed
   but whose columns are not makes `SaveAs` throw. Not a wrong result — a hard failure, on every
   save, for the life of the workbook.
2. **Data validation is silently lost.** A validation created on such a range survives in memory,
   writes an invalid reference on save, and comes back as nothing after reload.
3. **Styling silently does nothing.** Applying a style to such a range writes no cells, while
   assigning a value to the *same* range writes all of them.
4. **Counts go negative.** `RowCount()` and `ColumnCount()` return negative numbers while the range
   address's spans and `Cells().Count()` are correct. A table built over such a range reports zero
   fields, because it derives its field count from the negative one.
5. **Consolidation drops the range.** `Consolidate()` returns an empty collection where it should
   return the equivalent forward range.

The user wrote nothing invalid. They wrote corners in an order the library accepts.

## Solution

The rectangle a range occupies is decided in exactly one place, and every part of the library that
needs geometry asks that one place rather than working it out from the address.

The range address keeps its current observable behaviour — it may still be reversed, and
`IsNormalized` still reports so, because callers depend on that and tests pin it. What changes is
that nothing downstream computes geometry from the address any more. Everything goes through the
value-typed rectangle, and that rectangle is normalised per axis at the single point where an
address becomes one.

After this spec, a reversed range behaves exactly like the forward range with the same corners:
it saves, it styles, it validates, it counts, it consolidates.

## User Stories

1. As a library consumer, I want a range written with its corners in either order to save without
   throwing, so that my application does not fail on input a user could reasonably produce.
2. As a library consumer, I want a conditional format on a reversed range to be written to the file,
   so that the format I asked for reaches Excel.
3. As a library consumer, I want a data validation on a reversed range to survive a save and reload,
   so that reopening my own output does not lose rules.
4. As a library consumer, I want applying a style to a reversed range to style its cells, so that
   styling and value assignment agree about which cells a range contains.
5. As a library consumer, I want `RowCount()` and `ColumnCount()` to return the same magnitudes as
   the address's row and column spans, so that I can size a loop without checking the sign.
6. As a library consumer, I want a table created over a reversed range to have the same fields as one
   created over the equivalent forward range, so that table features work regardless of corner order.
7. As a library consumer, I want `Consolidate()` to include a reversed range in its output, so that
   merging ranges does not silently discard some of them.
8. As a library consumer, I want a range's merged, sorted and cleared behaviour to be identical for
   both corner orders, so that corner order is never load-bearing.
9. As a library consumer, I want a reversed range used as a formula reference to evaluate rather than
   throw, so that the calc engine accepts everything the object model accepts.
10. As a library consumer, I want the quad-tree range index to find a reversed range, so that
    intersection queries do not depend on how the range was constructed.
11. As a library consumer, I want `IsNormalized` to keep reporting whether the address I built was
    normalised, so that existing code that inspects it keeps working.
12. As an XLibur maintainer, I want one place that decides what "reversed" means, so that fixing a
    geometry bug fixes it everywhere at once.
13. As an XLibur maintainer, I want new geometry consumers to be unable to receive a reversed
    rectangle, so that the next consumer cannot reintroduce this defect class.
14. As an XLibur maintainer, I want the defensive normalisation currently scattered through the cell
    enumerator to become unnecessary, so that the code stops paying for an invariant twice.
15. As an XLibur maintainer, I want the calc engine's reference type to stop throwing defensively on
    un-normalised input, so that the exception it currently needs can be deleted.
16. As an XLibur maintainer, I want a single property test that exercises every geometry consumer
    against all four corner orders, so that a regression in any one of them fails a test.
17. As a contributing agent, I want the geometry seam to be obvious from the type signatures, so that
    I can tell where a rectangle is guaranteed normalised without reading every consumer.

## Implementation Decisions

**The seam is the value-typed rectangle.** The area type is the single conversion point from a range
address to geometry, and it normalises per axis. This was chosen over normalising at range-address
construction, which would be a higher seam but would change public behaviour and break the tests that
deliberately pin an un-normalised address.

**Normalisation is per axis, not per corner.** The current conversion tests whether *either* the row
or the column is inverted and, if so, swaps *both* corners together. That is correct when both axes
are inverted and wrong when only one is — it produces a rectangle with a negative width, which is the
root of every defect above. The range address type already implements the correct per-axis rule,
including the fixed-row and fixed-column bookkeeping that goes with it; that behaviour is the
reference, and the conversion adopts it.

**Consumers read the rectangle, not the address.** The row and column count methods on the range base
derive from the rectangle rather than subtracting address corners. Style application, cell
enumeration, consolidation, the quad-tree's coverage test, and reference writing all take the
rectangle. No consumer computes a span from two addresses.

**No public behaviour changes.** The range address remains constructible in reversed form and
`IsNormalized` continues to report it. The counts change sign, which is a bug fix, not an API change:
they become the magnitudes the spans already report.

**The calc engine's reference precondition becomes redundant.** The reference type currently rejects
un-normalised input with an exception. Once the rectangle is the only way geometry arrives, that
precondition cannot fire; it is removed rather than left as unreachable code.

**Ordering.** The conversion fix lands first, with the tests that would have caught each defect. The
consumer migration follows, one consumer per commit, so a regression bisects to a single consumer.

## Testing Decisions

**What makes a good test here.** A good test drives the public object model and asserts on observable
results — a saved file, a reloaded workbook, a cell's style, a count, a collection's contents. It
does not assert that a particular internal method normalised its argument. The defects are all
observable from outside; the tests should be too.

**The property test is the centrepiece.** For every rectangle, and for each of its four corner
orderings, the following must agree: cell count, row and column counts, the address's spans, the set
of cells returned by enumeration, the set of cells a style reaches, the result of consolidation, the
ranges a data validation reports after a save and reload, and a table's field count. One test, one
loop, covering all five defects and any sixth that has not been found yet.

**Each defect also gets a named regression test**, because a property test that fails tells you
something is wrong but not which thing. Five named tests, each reproducing the user-visible symptom:
the save that throws, the validation that vanishes, the style that does nothing, the negative count,
the dropped consolidation.

**Prior art.** The range address tests already pin the per-axis normalisation rule in isolation,
including the mixed-inversion cases; those assertions are correct and stay. What is missing is any
test that runs the same cases through a *live* range, which is exactly the gap this spec closes. The
area tests cover the value type's own behaviour and are the right place for the conversion fix.

**Test seam.** Everything above is reachable through `IXLWorksheet`, `IXLRange`, `IXLRanges` and a
save/reload round trip. No new test seam is introduced.

## Out of Scope

- Changing whether a reversed range address can be constructed, or the meaning of `IsNormalized`.
- The two consolidation implementations. Collapsing them is spec 51; this spec fixes only the
  normalisation input that makes them disagree today.
- The quad-tree's allocation behaviour on whole-row ranges, and the promotion policy between its two
  adapters. Recorded as backlog notes from the same review.
- The range index's duplicate-detection asymmetry, which is latent and not reachable today.
- Performance. The conversion is already on the hot path; the fix must not regress it, but this spec
  claims no improvement.

## Further Notes

The suite currently pins the *existence* of un-normalised addresses without ever pinning what a live
range built from one should do. That is the shape this review round found repeatedly: the fact has
two implementations, and the test sits on one side of the seam rather than across it.

The fatal case — the save that throws — is worth treating as the lead symptom when writing the
commit history, because it is the one a user cannot work around. The other four are silent, which is
worse in principle but easier to live with in practice.

The five defects were reproduced against a scratch build before this spec was written; the observed
values are recorded in the round-3 architecture review report.

## Results

**Implemented 2026-08-30 on `task/36` (worktree `xl-wt-36`), head `80a5b77a`, 15 commits, cut from
`upstream/main` `37c986bb`. Not yet pushed or merged.** Full suite green on both TFMs: XLibur.Tests
28,524 / 0 failed / 10 pre-existing skips; Report 962, Fonts.SixLabors 62, Fonts.SkiaSharp 74, all
green. Cell-enumeration benchmark (`L6_CellsUsed`, 50K×10) 19.45 ms → 19.49 ms median-of-three, noise.

**The mechanism was exactly as predicted.** `Area.FromRangeAddress` swapped both corners when either
axis was inverted; per-axis `Min`/`Max` fixed it (`ef849725`). **What the spec did not anticipate:**
three of the five defects (the `SaveAs` throw, the lost validation, the styling no-op) were fixed by
that one commit alone, because those consumers already read `Area` and only inherited its bug. Only
`RowCount`/`ColumnCount` and `Consolidate` computed from `FirstAddress`/`LastAddress` directly and
needed their own migration. "Reference writing" needed **no** change — an attempt to write
`Area.ToString()` directly changed a 1×1 array formula's `ref` from `B6:B6` to `B6` and broke two
golden files; reverted.

**Named regression tests** (`ReversedRangeGeometryTests.cs`, all seen red in `65acc1b4`):
`SavingConditionalFormatOnRangeWithReversedRowsDoesNotThrow`,
`DataValidationOnReversedRangeSurvivesSaveAndReload`, `StyleAppliedToReversedRangeStylesItsCells`,
`RowCountAndColumnCountOnReversedRangeReturnPositiveMagnitudes`, `ConsolidateIncludesReversedRange`,
plus `QuadTreeFindsReversedRangeThatSpansAQuadrantBoundary` and
`FormulaReferencingReversedRangeEvaluates`. Property test
`AllCornerOrdersAgreeAcrossEveryGeometryConsumer` (`ReversedRangePropertyTests.cs`) covers four
shapes × four corner orders across count, enumeration, style, consolidation, validation, table
fields, merge, table save and index intersection.

**The in-branch review widened the scope, correctly.** `/code-review` plus the agent's own quad-tree
probe found three more sites in the same class, all inside this spec's user stories 8 and 10, and
they were fixed in a second pass with the seam decision unchanged (consumers read `Area`;
`XLRangeAddress.Contains`/`Intersects` stay un-normalised): the range index's flat-list path
(`16542066`), the quadrant's point-containment path — `ws.Range("B5:E2").Merge()` then writing B3 did
not see the merge (`1d7e55ed`), relative `Range(int,int,int,int)`/`GetRange` anchoring plus the table
part writer — a table over a reversed range threw on save (`73b8e5f0`), and
`XLRangeColumn`/`XLRangeRow.CellCount` (`c2d085c6`). `Area.FromRangeAddress` gained a non-generic
`XLRangeAddress` overload to keep the per-cell-write merge check allocation-free; both overloads share
one `FromCorners`, so there is still one normalisation.

**Deliberately not done — recorded as D24.** A grep of every `.FirstAddress`/`.LastAddress` read in
`XLibur/` found ~12 further sites that compute geometry from raw corners: `XLTable` structural edits
(totals/header toggles, `ExpandTableRows`), `XLTableRange` row bounds, `XLCellCopyHelper`
(`CopyTo`/`AutoFill` origin and counts), `XLBorder` segment indexing, `XLRangeSetOperationsHelper`
`Grow`/`Shrink`, `XLRangeCellsHelper` clamping, `XLRange` insert/copy/transpose destinations and
`Row(int)`/`Column(int)`, and — uncertain — `XLAutoFilter`, `XLWorkbook_Load`, and the two formula
shifters. Plus `XLHelper.IsValidRangeAddress` rejecting a reversed address, which makes a reversed
data-validation list source get quoted as a literal. Not fixed here: the volume exceeded what could be
verified red-first on this branch, and one earlier attempt had already regressed a golden file.

**Third pass, 2026-08-30 evening, `80a5b77a` — driven by the owner in the agent's tab from a
`/code-review` on the branch. Two findings applied, two scoped out.** (1) The first pass moved the
*counts* onto the rectangle but left every member that addresses a cell *relative* to the range —
`Cell(in XLAddress)`, `ColumnQuick`, `Row(int)`, `Column(int)`, `FirstCell`/`LastCell`, `Rows()`/
`Columns()`, and the four `First`/`LastRow`/`ColumnUsed` probes — anchored on `RangeAddress.FirstAddress`.
With the two disagreeing, `LastCell()` on `B5:E2` returned `E8` and `Rows()` yielded `B5:E5..B8:E8`,
three rows outside the range, so a style or value written through `Rows()` landed on unrelated
cells. All now measure from the rectangle. (2) **A regression the branch itself introduced:** routing
the counts through `SheetRange` made `RowCount()`/`ColumnCount()` throw `InvalidOperationException`
for a `#REF!` range (both corners `InvalidAddress`, which `XLRangeShiftHelper` produces when a delete
swallows a range whole) — and the save and copy paths call them on stored ranges. Fixed by
`SheetRangeUnchecked`, the same per-axis normalisation without the validity guard. Six red-first
tests in `ReversedRangeGeometryTests.cs`; suite green on net10.0 across all four projects, and on net8.0
(14,268 / 0 failed, run by the orchestrator). **Lead for D24 / whoever touches `Reference.cs`:** its
`List<XLRangeAddress>` ctor doc still promises normalised areas and `RangeOp.ExpandBoundingBox`
still says "Areas are normalized", with no guard behind either now that the precondition is gone.

**Least sure of (agent's own list).** The `CellCount` fix was not independently cycled red — it was
reasoned by exact analogy to the verified `RowCount` fix. The audit's classification of the formula
shifters and `XLAutoFilter`/`XLWorkbook_Load` is unverified.

**What the next consumer inherits.** `Area.FromRangeAddress` is the one place normalisation happens;
any new consumer reads `Area`/`SheetRange` bounds, never `RangeAddress.FirstAddress`/`LastAddress`.
Spec 51 now receives normalised input to both consolidation engines. Spec 50 (`Intersection`) should
build on `Area`. D24 is the follow-on spec.
