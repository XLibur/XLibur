# Spec 37 — One way to reduce an argument to a scalar

**Area:** Architecture · **Defect (22 functions unusable with references)**
**Effort:** M (~4–5 days)
**Dependencies:** None hard. File-disjoint from spec 30, which owns the function-definition array
path. **Must land before spec 32** — 32 rewrites 411 registrations across the same function families
and would collide head-on.
**Status:** 🟩 Implemented on `task/37` (2026-08-30), unmerged — see Results. From the 2026-08-30 architecture review (round 3).

## Problem Statement

Every dynamic-array function returns `#VALUE!` when one of its scalar parameters is a cell reference
instead of a literal. So do the four regression functions.

```
=SEQUENCE(A1)               ->  #VALUE!     Excel: 1;2;3
=SEQUENCE(3,B1)             ->  #VALUE!     Excel: a 3x2 grid
=TAKE(A1:A3,B1)             ->  #VALUE!     Excel: 3;1
=SORT(A1:A3,B1)             ->  #VALUE!
=UNIQUE(data,B1)            ->  #VALUE!
=XLOOKUP(k,a,b,,C1)         ->  #VALUE!
=LINEST(A1:A4,B1:B4,D1)     ->  #VALUE!     Excel: 2
=TREND(A1:A4,B1:B4,D2,D1)   ->  #VALUE!     Excel: 10
```

The same functions work when the argument is written as a literal — `=SEQUENCE(3)`, `=TAKE(A1:A3,2)`,
`=LINEST(A1:A4,B1:B4,TRUE)`. Putting the count, the index, the sort order or the flag in a cell and
referring to it is the ordinary way people write these formulas, and it is the way that fails.

Eighteen dynamic-array functions and four regression functions are affected. The feature ships
unusable for anything but hard-coded arguments.

## Solution

There is one module that answers "give me this argument as a single scalar value, or the error". It
knows every way an argument can arrive — already a scalar, a one-cell array, a single-cell reference,
a multi-cell reference that needs implicit intersection, a multi-area reference that cannot be
reduced — and it returns the same answer for all of them regardless of which function asked.

Every function that needs a scalar argument calls it. None of them re-derives the reduction.

## User Stories

1. As a spreadsheet author, I want `=SEQUENCE(A1)` to produce a sequence of the length held in A1, so
   that I can change the size of a spill by editing a cell.
2. As a spreadsheet author, I want `=SEQUENCE(3,B1)` to use B1 as the column count, so that both
   dimensions can be driven from cells.
3. As a spreadsheet author, I want `=TAKE(range, B1)` and `=DROP(range, B1)` to take a count from a
   cell, so that I can build a parameterised extract.
4. As a spreadsheet author, I want `=SORT(range, C1)` to take its sort index from a cell, so that a
   user of my sheet can change the sort column without editing a formula.
5. As a spreadsheet author, I want `=SORTBY`, `=FILTER`, `=UNIQUE`, `=CHOOSECOLS`, `=CHOOSEROWS`,
   `=EXPAND`, `=TOCOL`, `=TOROW`, `=WRAPCOLS`, `=WRAPROWS`, `=HSTACK` and `=VSTACK` to accept cell
   references in every scalar slot, so that the whole dynamic-array family behaves consistently.
6. As a spreadsheet author, I want `=XLOOKUP` and `=XMATCH` to take their match mode and search mode
   from cells, so that lookup behaviour can be configured on the sheet.
7. As a spreadsheet author, I want `=TEXTSPLIT` to accept a reference for its delimiter, so that the
   delimiter can be edited without touching the formula.
8. As a spreadsheet author, I want `=LINEST`, `=LOGEST`, `=TREND` and `=GROWTH` to accept a reference
   for their constant and statistics flags, so that regression options can be driven from cells.
9. As a spreadsheet author, I want a scalar argument that refers to a whole column to resolve by
   implicit intersection the way Excel does, so that formulas copied from Excel behave the same.
10. As a spreadsheet author, I want an implicit intersection that finds no cell to give me Excel's
    error rather than a different one, so that I can diagnose it using what I already know.
11. As a spreadsheet author, I want a multi-area reference in a scalar slot to give `#VALUE!`, so that
    a genuinely unreducible argument is reported rather than guessed at.
12. As a spreadsheet author, I want a blank cell in a scalar slot to behave as Excel treats a blank,
    so that empty inputs do not need special handling in my formulas.
13. As a spreadsheet author, I want a text function given an unreducible argument to return an error
    rather than throw, so that one bad cell does not abort a whole calculation.
14. As an XLibur maintainer, I want the reduction rule to exist once, so that a change to implicit
    intersection semantics is one edit rather than eight.
15. As an XLibur maintainer, I want a newly registered range-accepting function to get correct
    reference handling without writing any reduction code, so that the next function cannot ship with
    this defect.
16. As an XLibur maintainer, I want the dead reduction branches deleted, so that reading a function
    family does not suggest a code path that never executes.
17. As an XLibur maintainer, I want the value model to stop handing callers a type they cannot use, so
    that the type system enforces the step callers currently have to remember.
18. As a contributing agent, I want one documented answer to "how do I get a scalar out of an
    argument", so that I do not copy whichever nearby helper I happen to find.

## Implementation Decisions

**The seam is on the value model.** All eight existing copies already hold the same argument type, so
the reduction becomes a member of that type. No new abstraction is introduced and no call site has to
learn a new concept — it swaps one call for another.

**The reduction is an ordered ladder, stated once:**

1. already a scalar — return it;
2. an array — take its first element;
3. a reference to a single cell — read that cell;
4. a reference to more than one cell — apply implicit intersection against the calling formula's
   address, then read the resulting cell;
5. a multi-area reference, or an intersection that selects nothing — return the appropriate error.

Only one of the eight existing copies implements all five steps. Two implement step 1 alone. Five
implement steps 1–3 and then attempt step 4 through an idiom that can never succeed.

**The root cause is an interface, not a typo.** The existing implicit-intersection member returns a
*reference* for the single-cell case, and the scalar accessor its callers then use rejects references
by definition. Every caller is therefore obliged to remember an extra unwrap that nothing in the
signature asks for, and seven of eight forgot. The fix is not to add the unwrap seven times: the
intersection member stops being part of the public reduction path — it becomes an implementation
detail of the new one.

**Existing correct behaviour is the reference.** The calc engine's own cell-content reduction already
implements the full ladder correctly and has done so all along. Its behaviour is what the new module
must reproduce; it then becomes a caller like the others.

**The function-definition intersection pass stays.** Arguments to functions registered as scalar are
already reduced before the function body runs, which is why five of the copies escape the defect
today. That pass is not removed by this spec — it is why the dead branches are dead. Removing it is
spec 30's and spec 32's territory.

**No public API change.** Everything here is internal to the calc engine. The observable change is
that formulas which returned `#VALUE!` now return values.

## Testing Decisions

**What makes a good test here.** A good test is a formula and its expected result. It sets cell
values, sets a formula, reads the answer, and compares against what Excel produces. It does not
assert that a particular reduction helper was called, and it does not test the reduction module in
isolation from the functions that use it — the defect lived precisely in the gap between a correct
rule and the functions that failed to apply it.

**The centrepiece is a matrix over argument shapes.** For each function that takes a scalar
parameter, and for each of six input shapes in that parameter — a literal, a single-cell reference, a
one-cell array, a row range requiring implicit intersection, a column range requiring implicit
intersection, and a multi-area reference — assert the result. That is the test the current shape
cannot express, because there is no single place where the rule lives.

**Every affected function gets at least one reference-argument case.** The suite currently passes
literals in every scalar slot of every dynamic-array function, which is exactly why 22 broken
functions shipped. A test per function, using a reference, is the minimum that would have caught it.

**Excel is the oracle, not the current implementation.** Expected values are Excel's answers,
recorded in the test. Where XLibur's answer differs and the difference is deliberate, the test says
so.

**Prior art.** The dynamic-array function tests and the regression function tests are the right homes;
they already build a worksheet, set a formula and assert a spilled result. The change is to their
inputs, not their shape.

**Test seam.** Setting a formula through `IXLCell` and reading the result. Highest available; no new
seam.

## Out of Scope

- The function registration overloads and the argument encoding they carry. That is spec 32, which
  this spec must precede.
- Per-element array application inside the function definition — spec 30.
- Excel's 2016-versus-365 difference in implicit intersection semantics. This spec reproduces one
  behaviour consistently; choosing which is a separate question, and the new module is where that
  choice would later be made.
- Adding functions. This spec makes existing functions work with existing argument kinds.
- The text-function throw noted in spec 30, which has its own owner.

## Further Notes

This is the widest live defect the round-3 review found, and the cheapest to characterise: eight
implementations of one rule, five of them containing a branch that provably never executes.

The test gap is the interesting part. Nothing about the suite is careless — there are tests for every
one of these functions. They all pass literals, because a literal is the natural thing to write when
you are demonstrating what a function does. The defect lives in the difference between demonstrating
a function and using one.

All the wrong answers above were reproduced against a scratch build before this spec was written.

## Results

**Implemented 2026-08-30 on `task/37` (worktree `xl-wt-37`), head `42bf46f8`,
13 commits, cut from `upstream/main` `37c986bb`. Not yet pushed or merged.** Full suite green on
both TFMs: 14,285 total, 14,280 passed, 5 pre-existing skips, per TFM.

**What the spec predicted that turned out wrong.**

- **22 affected functions is 20.** `HSTACK` and `VSTACK` take no scalar parameter at all. The live
  set is 16 in `DynamicArray.cs`, `TEXTSPLIT`, and the four regression functions.
  `CHOOSEROWS`/`CHOOSECOLS` already worked — their index goes through `TryPickCollectionArray`,
  which happens to handle a single-cell reference — but they got reference-argument tests anyway.
- **"Five copies contain a branch that provably never executes" is right about the idiom, wrong about
  the consequence.** Four of those five (`DateAndTime.ToScalar`, `SampleStatistics.TryGetScalarNumber`,
  `Financial.TryScalarNumber`, `MathTrig.TryGetAggregateArgument`) were already correct for a
  single-cell reference through a `TryGetSingleCellValue` step the spec's description elides, and
  were *unreachable* dead code besides — every caller is registered `AllowRange.Except`/`Only`, so the
  function-definition intersection pass pre-reduces the argument first. The live bugs were
  `DynamicArray.TryScalarArg` and `Text.TryOptionalInt` (step 1 only) and `Regression.TryGetBoolean`
  (no step 3). The eighth copy, `XLCalcEngine.ToCellContentValue`, is the full ladder, as stated.
- **The fix itself introduced a regression the suite could not see**: `SORTBY`'s order-vs-`by_array`
  disambiguation broke for a one-row sort once a single-cell reference became a valid order. Found by
  the in-branch `/code-review`, pinned by `SortBy_TwoSingleCellByArraysStayByArraysWhenTheSortedRangeIsOneRow`,
  fixed in `389972b5`. Same lesson as round 2: run the review inside the task.

**What was done.** `AnyValue.TryReduceToScalar(ctx, out scalar, out error)` is the one ladder
(scalar → array[0] → single-cell ref → implicit intersection against the calling cell → error).
`ToCellContentValue` became a caller. All eight copies plus four more found by the same sweep now
delegate to it. Semantics reproduced: the pre-existing Excel-2016-style single implicit intersection
via `Reference.ImplicitIntersection(IXLAddress)`; no 365 `@`/spill semantics introduced. That choice
lives only in `AnyValue.TryReduceToScalar`.

**Deliberately not done.** The function-definition intersection pass (`FunctionDefinition.IntersectArguments`)
and `AnyValue.ImplicitIntersection` that backs it are untouched, per the brief — spec 30/32 territory.
The shape matrix is full six-shape for `SEQUENCE`, partial for `TEXTSPLIT` and `LINEST`, and one
reference case each for the other 17 — a scope call disclosed in `c760d375`'s body.

**Found, not fixed — recorded as D23.** A defined name's own formula is evaluated with
`FormulaAddress: null`, so any function inside it that needs a real implicit intersection throws
`MissingContextException` instead of returning `#VALUE!`. Pre-existing; spec 37 narrows one
`TEXTSPLIT` path onto it slightly.

**Least sure of (agent's own list).** The `SORTBY` fallback order (try `by_array` shape first, then
order) is checked against two constructed cases, not against Excel across pair counts; the eight-copy
census came from targeted greps, not an end-to-end read of `Functions/`.

**What the next consumer inherits.** Spec 32 rewrites registrations across these families and must
rebase onto this branch; every scalar slot now has a reference-argument test that will notice if a
rewrite regresses it. Spec 30's per-element array application should call `TryReduceToScalar` rather
than reintroduce a ladder.
