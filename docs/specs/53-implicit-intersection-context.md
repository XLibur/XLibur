# 53 — implicit intersection on operator operands

**Status:** ✅ **Merged** 2026-09-01 as [#427](https://github.com/XLibur/XLibur/pull/427)
(`0908c05e`). Suite green. Worktree `xl-wt-d38` and branch `task/d38-implicit-intersection` are
removed. Its in-branch review raised nine findings: two were fixed as D43, two were its own and
fixed here, and **five reached `main` as D44–D48** — two of which re-open defects #426 recorded as
fixed.
**Defect:** [D38](../DEFECTS.md) · **Investigation:**
[`D38-investigation-implicit-intersection-on-operands.md`](D38-investigation-implicit-intersection-on-operands.md)

**In one line:** a legacy formula now passes an operator's reference operand through implicit
intersection, so `=A1+B1:B3` at `C3` answers `47` instead of `142` — and stops materialising the
whole array to get there.

---

## The investigation's design premise was wrong, and Excel says so

The investigation proposed a four-row context table, whose third row read:

> Inside an argument to a parameter that allows ranges → array → no intersection.

**Excel does not do that.** Setting a formula through `Range.Formula` — the legacy property — makes
Excel insert the `@` implicit-intersection operator exactly where Excel 2016 would have intersected,
which turns Excel into an oracle for the rule. Measured on Excel 16.0 build 20228, formula at `K2`,
`A1:A9` and `B1:B9` populated:

| Typed | Excel stores | Excel shows |
|---|---|---|
| `=SUM(A1:A9-B1:B9)` | `=SUM(@A1:A9-@B1:B9)` | 3 |
| `=MIN(A1:A9-B1:B9)` | `=MIN(@A1:A9-@B1:B9)` | 3 |
| `=AVERAGE(A1:A9-B1:B9)` | `=AVERAGE(@A1:A9-@B1:B9)` | 3 |
| `=STDEV.S(A1:A9-B1:B9)` | `=STDEV.S(@A1:A9-@B1:B9)` | `#DIV/0!` |
| `=SUM(A1:A9)` | `=SUM(A1:A9)` | 60 |
| `=SUMPRODUCT(A1:A9-B1:B9)` | `=SUMPRODUCT(A1:A9-B1:B9)` | 81 |

Read rows 1–4 against row 6. SUM, MIN, AVERAGE and STDEV.S all accept ranges, and Excel intersects
inside every one of them. **SUMPRODUCT is the exception, and the reason is not that it accepts
ranges — it is that its parameters are array-typed.** Row 5 shows the other half: a *bare* reference
argument is never intersected, only an operator's operand is.

So Excel's rule is:

> An operator's reference operand is intersected wherever the operator sits, **except** inside an
> argument to an array-typed parameter.

That is both simpler and stricter than the investigation assumed.

## Why this spec does not implement Excel's rule in full

Two obstacles, and they point the same way.

**XLibur has no data that tells SUM apart from SUMPRODUCT.** Both register as
`FunctionFlags.Range, AllowRange.All`, as do MIN, MAX and COUNT. Excel's distinction is a
parameter's *type* (array vs range), which XLibur does not model. Implementing the full rule means
adding that classification to ~400 registrations, and a misclassification silently changes answers.

**The suite's expectations are written for array semantics inside range-accepting functions.**
`TTest_MatchesTheTDistributionOfEachTestStatistic` asserts on
`AVERAGE(A1:A9 - B1:B9) / (STDEV.S(A1:A9 - B1:B9) / SQRT(9))` in cell `D2`. Excel's legacy answer
for that cell is `#DIV/0!`; the test asserts the array answer. The brief for this work names that
test as one of six that must pass **unmodified**, and says that needing to change one is a finding
to argue rather than an edit to make. **This is that finding** — recorded here, not acted on.

## What was implemented: intersect at the top level only

An operator intersects its reference operands when it is **not inside a function argument**, in a
**legacy** formula that **has a cell to intersect against**.

| Position | Context | Intersects |
|---|---|---|
| Top level of a legacy formula anchored at a cell | scalar | **yes** — new |
| Anywhere in a dynamic-array or CSE array formula | array | no — unchanged |
| A formula with no anchoring cell (`Evaluate(expr)` with no address) | array | no — unchanged |
| Inside any function argument | array | no — unchanged |

Row 3 is not a compromise, it is a necessity: implicit intersection needs a row and a column, and
`worksheet.Evaluate("MIN(A1:A2-B1)")` supplies neither. Reading `CalcContext.FormulaAddress` without
one throws `MissingContextException`.

Row 4 is the conservative half. Excel would intersect inside SUM and MIN; XLibur does not, yet. The
gap is documented at the call site and in the table above.

### The mechanism

`CalcContext.IntersectOperands`, a mutable per-position flag, is the per-argument context the
formula-level `IsArrayCalculation` could not be.

- `XLCalcEngine.EvaluateFormula(expression, wb, ws, address, …)` — the funnel for every legacy
  entry point, cell formulas included — sets it when `UseImplicitIntersection && address is not
  null`. `EvaluateArrayFormula` builds its context separately and leaves it false, so array and
  dynamic-array formulas are untouched.
- `CalculationVisitor.Visit(CalcContext, FunctionNode)` saves it, clears it around argument
  evaluation, and restores it before the call. It must **restore**, not assume false: a function
  call can sit inside an operand of a top-level operator, as in `=A1+SUM(B1:B3)*C1:C3`, and the
  operator after the call still has to intersect.
- `AnyValue.BinaryOperation` and `AnyValue.UnaryOperation` call the existing
  `AnyValue.ImplicitIntersection(CalcContext)` on each operand when the flag is set.

`AnyValue.ImplicitIntersection` was rewritten as a tag test, not because it was wrong but because of
where it is now called from. It was a `Match` over all seven cases whose reference arm captured the
context, so calling it allocated a display class and a delegate — a cost paid once per scalar
function argument before, and twice per operator afterwards. **In-branch code review caught the
regression**; measured over 20,000 evaluations, bytes per evaluation:

| | at `720bae1a` | first D38 commit | after the rewrite |
|---|---|---|---|
| `A1+B1` | 128.1 | 304.1 | **128.1** |
| `A1+B1*B2-B3` | 128.1 | 656.1 | **128.1** |
| `A1&B1` | 224.1 | 400.1 | **224.1** |

The baseline column was measured in a throwaway worktree at the branch point, so "back to baseline"
is a measurement rather than an assertion.

**No new intersection logic was written.** `Reference.ImplicitIntersection` already implements
Excel's rule exactly — column vector on the formula's row, row vector on its column, `#VALUE!`
when the formula lies outside the span, `#VALUE!` for a multi-area reference. It had exactly one
caller, `FunctionDefinition.IntersectArguments`, and now has two. Its 2D branch returns `#VALUE!`
where Excel would intersect on both axes, which is unreachable in practice: a formula inside both
the row span and the column span of a rectangular range is inside the range, so that case is always
a circular reference.

### Cost, for free

An intersected operand is a single cell, so the array is never built. `worksheet.Evaluate("A1+B:B",
"C3")` goes from **692 ms / 24.1 MB** to **0 ms / 0.00 MB**, and the fuzzer's 566-column shape from
366 s / ~13.6 GB to an instant `#VALUE!`.

**This did not on its own retire the lazy-`Array.Apply` work.** The eager materialisation remained
reachable two ways, and neither can intersect: **through a function argument**, where this spec
deliberately suppresses intersection, and **through `Evaluate(expression)` with no address**, which
is the path the fuzz target uses. That work followed and is written up below.

## The performance half: lazy element access

Landed after the correctness fix, as its own commits, and **behaviour-preserving** — the suite is a
complete gate for it and no answer changed.

`Array.Apply` filled a `ScalarValue[height, width]` for the whole rectangle whatever the consumer
wanted from it. Both overloads now return a lazy view — `BinaryArray` and `MappedArray` — in the
shape `ReferenceArray` and the broadcast views already had.

Then the fuzz target, no longer dying on its seed, ran 108,404 executions and immediately found the
**same defect in a second place**: `T(B1:C1/V+AM/U/+QU:B%+1)` took **59,757 ms and 11,088 MB**,
while the inner expression alone was 0 ms and 0.00 MB. Two survivors, both sized by an arbitrary
input range rather than by a result:

| | What it did | Reachable? |
|---|---|---|
| `Text.TArray` | filled `ScalarValue[array.Height, array.Width]` | **yes** — `T` accepts a range, so a 458-column operand is 480M elements; this is the one the fuzzer hit |
| `Reference.Apply` | filled `ScalarValue[height, width]` for a whole area | **no — it has no callers.** Latent, not live |

`Reference.Apply` was rewritten anyway, because leaving one member of the family eager is how the
next caller reintroduces the defect, and delegating to `ReferenceArray` is what it should have said
in the first place. But the commit message that landed it claimed a live cost it does not have; the
correction is recorded here and in the method's own remarks.

Both are now lazy. Allocation around a single evaluation:

| Expression | Before | After |
|---|---|---|
| `SUMPRODUCT(B:B*2)` | 24.00 MB | **0.00 MB** |
| `SUMPRODUCT(B:Q*2)` | 384.00 MB | **0.00 MB** |
| `SUM(B:B+0)` | 24.00 MB | **0.00 MB** |
| `T(B1:C1/V+AM/U/+QU:B%+1)` | 11,088 MB / 59,757 ms | **0.00 MB / 7 ms** |

Every other `new ScalarValue[…]` in the calc engine is sized by a result count — bins, modes,
matched rows — not by an input span, so this is the end of the family rather than the next
instalment.

**What is not fixed, and should not be:** time for a genuine aggregate. `SUMPRODUCT(B:Q*2)` still
scans 16M cells in ~900 ms, because that work is real rather than wasted; 566 columns is ~30 s of
honest iteration. Excel is no different. The trade the lazy form makes is that a consumer reading an
element twice computes it twice; element access is cheap and allocation-free, which is what makes
that acceptable.

## Acceptance

| # | Criterion | Result |
|---|---|---|
| 1 | `ImplicitIntersection_AppliesToAnOperandOfABinaryOperator` passes with `[Skip]` removed | ✅ `A1+B:B` 47, `A1+B1:B10` 47, `A1*B:B` 210 |
| 2 | The six listed tests still pass, unmodified | ✅ none edited |
| 3 | `Evaluate("A1+B:B", "C3")` is single-digit ms and megabytes | ✅ 0 ms, 0.00 MB |
| 4 | `SetDynamicFormulaA1("A1+B1:B3")` still spills 142, 49, 47 | ✅ |
| 5 | The fuzzer's input added to the seed corpus | ✅ `XLibur.Fuzz/corpus/formula/operand-implicit-intersection-whole-column` |
| 6 | A short `formula` run completes without a timeout artifact | ⚠️ **partly** — the allocation timeouts are gone; a dense-scan one remains |

### Criterion 6 belonged to the performance half, not the correctness half

The `formula` fuzz target evaluates through `sheet.Evaluate(formula)` — **with no formula address**
(`XLibur.Fuzz/FuzzTargets.cs`, `RunFormula`). That is exactly the row of the context table that
cannot intersect, because there is no cell to intersect against.

**The investigation expected otherwise.** It said an intersected operand "never builds the array at
all, and the lazy view then costs nothing", which is true — but only for a formula anchored at a
cell, and the fuzzer's own reproducer is not one. So the ordering the brief gave was right for a
stronger reason than it stated: **the lazy views are not merely a separable half, they are the only
half that could close the fuzz target.**

Three 420-second runs tell the story:

| After | Result |
|---|---|
| the correctness fix alone | dies on its own seed at 11 s — one timeout, no progress |
| lazy `Array.Apply` | 108,404 executions, then a **second** timeout: `T(B1:C1/V+AM/U/+QU:B%+1)` |
| lazy `TArray` and `Reference.Apply` | a full run, no timeout — ends on a 3-byte crash, `[ ]`, recorded as D42 |
| the same, run again on the rebased base | a **new** timeout: `SUM(1,TB:U/3,2,B2,)` |

The target is unblocked and finding new things, which is what "no longer gated on this defect"
means. D42 is a parser defect with nothing to do with D38 and is not fixed here.

**Correction to an earlier draft of this spec**, which said the target now produces "no timeout
artifact of any kind". That was true of one run and is too strong. The fourth row is a **dense
aggregate scan, not a materialisation** — 29,153 ms at **0.00 MB**, so the allocation family really
is closed:

| | Time | Allocated |
|---|---|---|
| `SUM(B:E)` — bare range | 0 ms | 0.00 MB |
| `SUM(B:E/3)` — an operator in the argument | 235 ms | 0.00 MB |
| `SUM(1,TB:U/3,2,B2,)` — ~500 columns, 524M cells | 29,153 ms | 0.00 MB |

A bare range reaches the sparse iterator in `CalcContext.GetNonBlankValues` and skips blanks; wrap
the operand in an operator and the argument arrives as an `Array`, so the aggregate falls to
`array.Where(…)` and walks every cell, blanks included.

**This is the same missing piece as gap 1 below.** Under Excel's rule the operand intersects to a
single cell, so `=SUM(@TB:U/3)` is instant — closing the function-argument gap would close the last
performance cliff as well as the correctness one. That is an argument for closing it which the
correctness case alone did not make, and it is the strongest reason to revisit the three tests named
there.

Also verified against Excel, all at `C3` with `A1=42`, `B1=100`, `B2=7`, `B3=5`:

| Formula | Excel | XLibur before | XLibur after |
|---|---|---|---|
| `A1+B1:B3` | 47 | 142 | **47** |
| `A1*B:B` | 210 | 4200 | **210** |
| `A1&B1:B3` | `425` | `42100` | **`425`** |
| `A1>B1:B3` | TRUE | FALSE | **TRUE** |
| `-B1:B3` | -5 | -100 | **-5** |
| `B1:B3%` | 0.05 | 1 | **0.05** |
| `A1+E1:H3` | `#VALUE!` | 42 | **`#VALUE!`** |
| `B1:B3` (bare) | 5 | 5 | 5 |

## Known gaps, deliberate

1. **Inside a function argument, nothing intersects.** Excel intersects inside SUM, MIN, AVERAGE and
   every other range-accepting function; only array-typed parameters stop it. Closing this needs a
   per-parameter array/range classification and would change what
   `TTest_MatchesTheTDistributionOfEachTestStatistic` and the two `ChiSqTest_*` tests assert.

   Concretely, at `C3` with `B1:B3` = 100, 7, 5: `=SUM(B1:B3+0)` is **5** in Excel (stored as
   `=SUM(@B1:B3+0)`) and **112** in XLibur. `ImplicitIntersection_DoesNotApplyWithoutAScalarContext`
   pins the current answer so the gap is a decision on the record rather than a surprise.
2. **Unary plus does not intersect.** `AnyValue.UnaryPlus()` returns its operand unchanged and takes
   no `CalcContext`; Excel stores `=+B1:B3` as `=+@B1:B3`. Fixing it means giving `UnaryPlus` a
   context parameter.
3. **`Evaluate(expression)` with no address keeps array semantics.** Unavoidable, and now the
   documented answer to the investigation's question about what
   `IXLWorksheet.Evaluate(expression, formulaAddress)` is: with an address it is a legacy formula;
   without one it cannot be, because there is nothing to intersect against.
4. **A 2D operand that contains the formula cell** returns `#VALUE!` where Excel reports a circular
   reference. Both are errors and the case is degenerate.
