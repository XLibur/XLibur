# Spec 04 — Demand-Driven Formula Evaluation (kill the full-recalc cliff)

**Area:** Performance (read time + memory) + Architecture (calc engine)
**Effort:** L (2–3 weeks; subtle correctness work)
**Dependencies:** None. Read `docs/dynamic-array-spill-phase-b.md` first — spill ownership interacts with recalc.
**Status:** Proposed

## Summary

Reading `cell.Value` on a dirty formula cell can trigger **full-workbook recalculation**: `XLCell.Evaluate` tries `TryEvaluateSingleCell` and, on any `GettingDataException` (formula references another dirty cell), falls back to `CalcEngine.Recalculate(wb, null)` — which lazily builds the full `XLCalculationChain` + `DependencyTree` (~176 MB for 250K formulas) and iterates every formula in the workbook. The fix: evaluate the *precedent closure* of the requested cell recursively instead of falling back to everything.

## Current state

- `XLibur/Excel/CalcEngine/XLCalcEngine.cs` (670 lines) — `TryEvaluateSingleCell` (single-cell fast path, no tree), `Recalculate` (full chain), `MarkDirty` (builds `DependencyTree` lazily when `_needsDependencyTree`).
- `XLibur/Excel/Cells/XLCell.cs` — `Evaluate(force)` (~line 338): clean-check via `Formula.IsDirty(workbook)` is cheap (epoch compare); dirty → `TryEvaluateSingleCell` → fallback `Recalculate(wb, null)`.
- `XLibur/Excel/CalcEngine/DependencyTree.cs` (404 lines) — real dependency graph; per-sheet `RBush<AreaDependents>` R-trees; built by full workbook scan in `CreateFrom`.
- `XLibur/Excel/CalcEngine/XLCalculationChain.cs` (411 lines) — intrusive linked chain + `Dictionary<XLBookPoint, Link>`; uses `GettingDataException` to reorder evaluation when a formula reads a dirty precedent.
- Dirty tracking is epoch-based: `XLCellFormula._evalEpoch` vs `XLWorkbook.EditEpoch` — any edit conceptually dirties the whole book; RBush + explicit-dirty flags refine it.
- Known history: an earlier attempt to lazily build the tree inside `MarkDirty` failed on unparseable formulas (external workbook refs, legacy TABLE) and unexpected build timing (`Mark_dirty_wont_crash_on_cycle`). Design around this: **the demand-driven path must not require the tree at all.**

## Design

### Core: recursive demand evaluation with cycle guard

Replace the `TryEvaluateSingleCell`-or-full-recalc dichotomy with:

```
EvaluateOnDemand(cell):
    if formula is clean → return cached
    push cell onto evaluation stack (detect cycles → #REF!/circular handling per existing semantics)
    evaluate AST; when the evaluator reads another cell:
        if that cell has a dirty formula → EvaluateOnDemand(it) first (recurse)
        else → read its slice value directly
    store result, stamp _evalEpoch, pop stack
```

This is textbook demand-driven recalc (what `GettingDataException` + chain-reordering approximates globally, done locally). Concretely:

1. Add an evaluation-stack (reused `List<XLBookPoint>` + `HashSet<XLBookPoint>` on the engine, cleared per top-level call) for cycle detection. On cycle: match Excel semantics already implemented for the chain (see how `XLCalculationChain` + cycle tests handle it — `Mark_dirty_wont_crash_on_cycle` and friends define expected behavior).
2. The hook point is where `CalculationVisitor`/`CalcContext` dereferences a cell value and currently throws `GettingDataException` for dirty cells. Instead of throwing, call back into demand evaluation. Keep the exception path intact for the full-`Recalculate` chain mode — add a mode flag on `CalcContext` rather than changing global behavior.
3. Depth limit (e.g. 10_000) to avoid stack issues on pathological chains — it's an explicit iterative stack, not CPU recursion, so the limit is about runaway work, not StackOverflow. On hitting the limit, fall back to full `Recalculate` (current behavior) — correctness is preserved.
4. **Unparseable formulas** (external refs, TABLE): if a precedent's formula can't be parsed, treat as current behavior (its cached value is used / full-recalc fallback). Never let demand evaluation crash where full recalc wouldn't.

### Secondary: don't build the dependency tree for read-only workloads

`MarkDirty` currently wants the tree. For the load-then-read pattern there are no edits, so the tree is never needed — verify that plain `Load` + `Value` reads with this change never construct `DependencyTree`/`XLCalculationChain` (assert via test with an internal counter or by checking `_dependencyTree is null` after the read pass).

Edits after reads must still propagate dirtiness correctly: the epoch mechanism already dirties everything on edit, so post-edit reads re-evaluate via the demand path. Confirm the two historically-failing scenarios are covered by tests: `EditCellInvalidatesDependentCells`, `RecalculateAllFormulas_recalculates_all_formulas_in_sheet_and_leaves_rest_dirty`.

### Interaction with spill (dynamic arrays)

A dirty spill-owner formula must be evaluated as a whole (its footprint cells derive from it). When demand evaluation hits a cell inside a spill footprint, resolve to the owner formula and evaluate that. See `TryGetDirtySpillOwner` in `XLCalcEngine` — reuse it. Add tests: read a cell in the middle of a dirty spill range.

## Work plan

| # | Task | Size |
|---|------|------|
| 1 | Read the chain/`GettingDataException` flow end-to-end; write a short design note in the PR confirming/adjusting the hook point | S |
| 2 | Demand-evaluation with explicit stack + cycle detection behind `CalcContext` mode flag | L |
| 3 | Wire `XLCell.Evaluate` to demand path; remove full-recalc fallback except for depth-limit/parse-failure cases | S |
| 4 | Spill-owner resolution in demand path + tests | M |
| 5 | Tests: cycles, cross-sheet precedents, dirty ranges (`SUM(A:A)` with dirty cells in area), post-edit invalidation, unparseable precedents | M |
| 6 | Benchmark: 250K×15 `LoadAndReadAllCells` read phase; add a formula-heavy read benchmark (e.g. 100K formula cells, read 100 random) demonstrating O(precedents) not O(workbook) | S |

## Acceptance criteria

1. Reading one dirty formula whose precedents are plain values: no `DependencyTree`, no `XLCalculationChain` construction, no workbook-wide iteration.
2. Reading 100 random formula cells out of 100K dirty formulas evaluates ≤ (100 × average precedent closure) formulas — verified with an internal evaluation counter in a test.
3. All existing calc-engine tests green, including cycle tests and the two dirty-propagation tests named above.
4. 250K×15 read-phase time and allocations improve or hold (the earlier prototype suggested large wins when formulas reference values; formula-chain sheets are the new win).
5. Full `Recalculate()`/`RecalculateAllFormulas()` public behavior unchanged.

## Risks

- Correctness. This is the highest-risk spec in the set: evaluation order and cycle semantics must match Excel/current behavior. Mitigate by keeping the chain-based full recalc untouched and diffing results: add a test harness that evaluates a corpus of workbooks both ways and compares every cached value.
- Volatile functions (NOW, RAND, OFFSET/INDIRECT dependencies): check how `DependenciesVisitor` flags them today; demand evaluation must not cache-and-skip volatiles differently than full recalc. Add explicit tests.

## References

- Prior in-progress work (`TryEvaluateSingleCell`) and its two failing tests — that attempt is the starting point, this spec is its completion with a sounder strategy.
- `docs/dynamic-array-spill-phase-b.md` for spill architecture.
