# D38 — investigation: implicit intersection is not applied to operator operands

**Status:** Investigation complete. **Not fixed**, deliberately. Wants its own spec.
**Defect:** [D38 in `DEFECTS.md`](../DEFECTS.md)
**Found by:** the `formula` fuzz target (spec 52), 2026-08-31, as a libFuzzer **timeout** rather than
a crash — a 16-byte input that takes 366 seconds.
**Investigated:** 2026-09-01. Every number below was measured on this machine, on
`task/52-fixes` at `720bae1a`, not estimated.
**Verified against Excel:** 2026-09-01, Excel 16.0 build 20228, on `task/d38-implicit-intersection`.
All three of the investigation's open questions are now answered; see *Excel's actual rule* and
*Open questions*. The verdict below is unchanged, and the scope is wider than it stated.

---

## Verdict

XLibur supports two formula kinds, and the calc engine is **correct for one of them and wrong for
the other in exactly one place**.

> A **legacy** formula does not pass a reference *operand* through implicit intersection, though it
> does pass a bare reference and a scalar function's arguments through it. So `=A1+B1:B3` at `C3`
> answers with the array's first element instead of the intersection — and builds the entire array
> to get there.

The **dynamic-array** path is correct throughout and needs no change. This is not a "which Excel are
we" question after all; the engine's model is coherent, and one branch of it is unfinished.

There are **two separable defects** behind one symptom, and they can be fixed independently:

| | |
|---|---|
| **Correctness** | The legacy path returns the wrong cell. Needs a context flag threaded through the visitor — a real design change |
| **Performance** | Every binary operation over a range **eagerly materialises the whole array**, ~24 MB and ~611 ms per million cells, regardless of which element is wanted. Fixable on its own, with **no change in any answer** |

**The performance half is the dangerous one and the cheaper one. Do it first.**

---

## Reproduction

```csharp
using var wb = new XLWorkbook();
var ws = wb.AddWorksheet("Data");
ws.Cell("A1").Value = 42;
ws.Cell("B1").Value = 100;
ws.Cell("B3").Value = 5;

ws.Evaluate("A1+B:B", "C3");   // 142, after ~640 ms. Excel 2016 says 47.
```

The fuzzer's own input is 16 bytes — `VLO` + five invalid UTF-8 bytes + `+A.+UU:B` — where `UU:B`
spans 566 columns. It takes **366 seconds**.

A `[Skip]`-marked gate is already in the suite:
`ArithmeticOperatorsTests.ImplicitIntersection_AppliesToAnOperandOfABinaryOperator`.

---

## What the engine actually does

`A1 = 42`, `B1 = 100`, `B2 = 7`, `B3 = 5`, formula anchored at `C3`.

| Formula | Legacy (`FormulaA1`) | Dynamic (`SetDynamicFormulaA1`) |
|---|---|---|
| `B1:B3` | **5** ✅ | 100, 7, 5 ✅ |
| `A1+B1:B3` | **142** ❌ *(should be 47)* | 142, 49, 47 ✅ |
| `SEQUENCE(3)` | 1 ✅ | 1, 2, 3 ✅ |

**Read the first two rows together.** The same range, in the same position, in the same formula
kind: `5` when it stands alone, `100` when it is an operand. One of those two is wrong whichever
semantics you target, so this needs no appeal to Excel's behaviour to be a defect — the engine
contradicts itself.

Which one is wrong *is* settled by the flag the engine already sets:
`CalcContext.UseImplicitIntersection => true`, whose own documentation says *"all arguments for
scalar functions were passed through implicit intersection before calling the function"*. That is
Excel 2016 semantics, under which the bare reference is right and the operand is wrong.

> ✅ **Verified against Excel, 2026-09-01.** Excel 16.0 build 20228. `=A1+B1:B3` at `C3` shows **47**,
> and Excel stores it as `=A1+@B1:B3` — the implicit-intersection operator sits on the operand. The
> check was also run on **a workbook XLibur wrote**: XLibur emits a bare `<f>A1+B1:B3</f>` with no
> cached value, Excel reads it as `=A1+@B1:B3` and shows 47, both on load and after a
> `CalculateFullRebuild`. `=B1:B3` at `C3` shows `5`, matching the legacy row above.
>
> The "should be 47" column stands as written. See *Excel's actual rule* below for the full
> characterisation, which is wider than this table assumed.

---

## Excel's actual rule, measured

Added 2026-09-01, from Excel 16.0 build 20228 driven over COM. Formulas were set through the legacy
`Range.Formula` property, which is what a legacy formula is; `Range.Formula2` is shown alongside
because it is where Excel reveals the `@` it inserted.

**Every operator kind intersects its operand, not just binary arithmetic.** Formula at `C3`;
`A1=42`, `B1=100`, `B2=7`, `B3=5`.

| Formula at `C3` | Excel stores | Excel shows | XLibur legacy |
|---|---|---|---|
| `=B1:B3` | `=@B1:B3` | `5` | `5` ✅ |
| `=A1+B1:B3` | `=A1+@B1:B3` | `47` | `142` ❌ |
| `=A1*B1:B3` | `=A1*@B1:B3` | `210` | `4200` ❌ |
| `=A1&B1:B3` | `=A1&@B1:B3` | `425` | `42100` ❌ |
| `=A1>B1:B3` | `=A1>@B1:B3` | `TRUE` | `FALSE` ❌ |
| `=-B1:B3` | `=-@B1:B3` | `-5` | `-100` ❌ |
| `=B1:B3%` | `=@B1:B3%` | `0.05` | `1` ❌ |

So **unary prefix, postfix `%`, comparison and concatenation share the defect** with binary
arithmetic. The gate test covers only `+` and `*`; a fix should be gated on the others too.

**The intersection itself is two-dimensional, and it can fail.** This is the part the "intersect on
the formula's row" shorthand does not capture. Measured with `A1=42` and a block at `E1:F3`:

| Operand shape | Formula cell | Excel |
|---|---|---|
| column vector `E1:E3` | `C2` — row inside | `49` (row 2) |
| column vector `E1:E3` | `C10` — row outside | `#VALUE!` |
| row vector `E3:H3` | `F10` — column inside | `48` (column F) |
| row vector `E3:H3` | `C3` — column outside | `#VALUE!` |
| 2D `E1:H3` | `F10` — column inside, row outside | `#VALUE!` |
| 2D `E1:H3` | `C3` — column outside, row inside | `#VALUE!` |
| 2D `B:E` | `C3` — both inside | circular reference (resolves to `C3` itself) |

The rule: a column vector intersects on the formula's **row**, a row vector on its **column**, a 2D
range on **both**, and any dimension whose span does not contain the formula gives `#VALUE!`.
XLibur currently returns element `[0, 0]` in every one of these cases — `A1+E1:H3` at `C3` answers
`42`, where Excel answers `#VALUE!`. **A fix must produce `#VALUE!`, not merely the right cell.**

## Measurements

Time and allocation are both **linear in the operand's cell count**, at ~611 ms and ~24 MB per
million cells. Measured with `GC.GetTotalAllocatedBytes(precise: true)` around a single evaluation.

| Columns | Cells | Time | Allocated | Result |
|---|---|---|---|---|
| 1 | 1,048,576 | 692 ms | 24.1 MB | 142 |
| 4 | 4,194,304 | 2,572 ms | 96.0 MB | 142 |
| 16 | 16,777,216 | 10,246 ms | 384.0 MB | 142 |
| 40 | 41,943,040 | 25,622 ms | 960.3 MB | 142 |

**Extrapolating to the fuzzer's input** (`UU:B`, 566 columns, 593M cells): **~362 s and ~13.6 GB**.
The measured time was 366 s, so the extrapolation holds.

That is the severity. **A 25-character formula in an untrusted workbook allocates about 13.6 GB and
holds a core for six minutes**, and a workbook can carry it: `EvaluateFormulasBeforeSaving` and
`RecalculateAllFormulas` both walk into it. .NET Core allows arrays over 2 GB by default, so this
does not even fail fast.

Bounded ranges are unaffected — `A1+B1:Q1000` is 10 ms — so nothing in ordinary use is slow. The
cost needs a full-height reference, which is exactly what a hostile or careless file supplies.

Also measured, for contrast: `SUM(B:Q)` is **5 ms**. Aggregates over the same range are fine; only
the operator path materialises.

---

## Where it is, in the code

```
CalculationVisitor.Visit(CalcContext, BinaryNode)
  └─ AnyValue.BinaryPlus / BinaryMinus / … (12 operators)
       └─ AnyValue.BinaryOperation                     ← the single funnel
            ├─ left.TryPickSingleOrMultiValue(…)       ← turns a Reference into a ReferenceArray
            ├─ right.TryPickSingleOrMultiValue(…)
            └─ Array.Apply(Array, BinaryFunc, ctx)     ← eagerly fills ScalarValue[height, width]
XLCalcEngine.EvaluateAndReduce
  └─ array => array[0, 0]                              ← keeps one element, discards the rest
```

Two facts make the shape clear:

1. **`AnyValue.ImplicitIntersection(CalcContext)` exists and has exactly one caller** —
   `FunctionDefinition.IntersectArguments`, reached from `CallFunction`. Operators never call it.
   Its own inline comment reads *"Array is unaffected by implicit intersection **for operands**"*,
   so it was written with operators in mind and never wired to them.

2. **`Array.Apply` is eager**:

   ```csharp
   var data = new ScalarValue[height, width];
   for (var y = 0; y < height; ++y)
       for (var x = 0; x < width; ++x)
           data[y, x] = func(leftItem, rightItem, ctx);
   return new ConstArray(data);
   ```

   `ReferenceArray` itself is lazy — it computes `this[y, x]` on demand — and that laziness is
   thrown away by the first operator applied to it.

---

## Why the obvious fix does not work

Intersecting both operands at the top of `BinaryOperation` when
`CalcContext.UseImplicitIntersection && !ctx.IsArrayCalculation` — mirroring exactly what
`CalculationVisitor` does when choosing `CallFunction` over `CallAsArray` — **fails six tests of
14,376**. This was run, not guessed.

| Test | Formula | Why it must keep the array |
|---|---|---|
| `ArraysOperation_BinaryOperationBetweenAreaReferenceAndSingleCellReferenceShouldWork` | `MIN(A1:A2-B1)` | `MIN` accepts ranges, so `A1:A2` must stay an array |
| `Filter_NoMatch_ReturnsIfEmpty` | `FILTER(A1:A2, A1:A2>9, "none")` | `FILTER` needs the whole boolean array |
| `ChiSqTest_IsTheRightTailOfTheStatisticItComputes` | `SUMPRODUCT((A1:C2 - E1:G2) ^ 2 / E1:G2)` | Nested operators inside a range-accepting parameter |
| `ChiSqTest_UsesOneLessThanTheLengthForAVector` | as above | as above |
| `TTest_MatchesTheTDistributionOfEachTestStatistic` | as above | as above |
| `ArraysOperation_MultiAreaReferencesArgumentResultsInScalarError` | `(A1:A1,A1:A2)+1` | Asserts the *scalar* `#VALUE!`, which the intersection changes |

**They are all the same shape, and they are all right.** An operator inside an argument to a
range-accepting parameter is in array context and must not intersect — in Excel 2016 as much as in
365, because the enclosing function is what supplies the context.

So the rule is not "operators intersect" or "operators don't". It is:

> **Whether an operand intersects depends on the context the operator's *result* flows into, not on
> the operator, and not on the formula's calculation mode.**

`CalcContext.IsArrayCalculation` is a **formula-level** flag. What is missing is a **per-argument**
one, and that is why gating on `IsArrayCalculation` does not work: all six failures are legacy
formulas, where the flag is false, containing a range-accepting function, where the context is array
anyway.

The information already exists — `FunctionDefinition._allowRanges` and `_markedParams`, the
`AllowRange` enum — but it is consulted *after* an argument subtree has been evaluated, so the
operators inside it have already committed.

---

## What a fix has to do

**Thread a scalar/array context through `CalculationVisitor` as it descends**, rather than deciding
once at the end:

| Position | Context |
|---|---|
| Top level of a legacy formula | scalar → operands intersect |
| Top level of a dynamic-array formula | array → no intersection *(already correct)* |
| Inside an argument to a parameter that allows ranges | array → no intersection |
| Inside an argument to a scalar parameter | scalar → operands intersect |

`BinaryOperation` then intersects only in scalar context. `IntersectArguments` becomes redundant
for the sub-expression case and should probably remain for the top-level argument case; that wants
checking rather than assuming.

**Decide explicitly what `IXLWorksheet.Evaluate(expression, formulaAddress)` is.** It currently
behaves as a legacy formula (`142`) and cannot spill, since it returns one `XLCellValue`. Legacy is
almost certainly right, but it is a public API and the choice should be written down rather than
inherited.

### The performance half can go first, and should

**Making `Array.Apply` lazy fixes the resource exhaustion without changing a single answer.** A
`BinaryArray` view computing `this[y, x]` on demand — the shape `ReferenceArray` already has —
makes `array[0, 0]` O(1), and the 366-second case becomes instant while still returning `142`.

This is worth doing on its own merits:

- It is a **denial-of-service fix**, and it lands without waiting on a semantics decision.
- It is **behaviour-preserving**, so the 28,774-test suite is a complete gate for it.
- It **does not become wasted work**: if the correctness fix later intersects the operand, the
  array is never built at all and the lazy view costs nothing.

Watch for: a consumer that iterates the result more than once now recomputes, so a lazy view wants
either cheap element access (it has that) or memoisation; and `func` captures `ctx`, so the view
must not outlive the evaluation. Both are ordinary, and `ReferenceArray` already lives with the
second.

---

## Acceptance

A fix is done when:

1. `ArithmeticOperatorsTests.ImplicitIntersection_AppliesToAnOperandOfABinaryOperator` passes with
   its `[Skip]` removed — `A1+B:B`, `A1+B1:B10` and `A1*B:B` at `C3` give 47, 47 and 210.
2. All six tests listed above still pass, unmodified. **If a fix needs one of them changed, that is
   a finding to argue in the PR, not an edit to make quietly.**
3. `ws.Evaluate("A1+B:B", "C3")` completes in single-digit milliseconds and allocates megabytes, not
   gigabytes.
4. The dynamic-array results are unchanged: `SetDynamicFormulaA1("A1+B1:B3")` still spills 142, 49,
   47.
5. The fuzzer's input is **added to the seed corpus** — `XLibur.Fuzz/corpus/formula/`, base64
   `VkxPsbSq/9crQS4rVVU6Qg==`. It is deliberately absent today because it would make every run spend
   six minutes on a known defect; once fixed it is a good regression seed.
6. A short `formula`-target run completes without a timeout artifact. **That target is currently
   gated on this defect** and cannot get a clean run until it is fixed.

---

## Open questions, not settled here

- ~~**Does Excel actually return 47?**~~ **Answered 2026-09-01: yes.** See *Excel's actual rule*
  above. Verified twice — typed fresh into Excel, and on a workbook XLibur itself wrote.
- ~~**A legacy formula in a cell whose operand spans multiple columns throws.**~~ **Answered
  2026-09-01: not a defect.** `C3` holding `=A1+B:E` throws `InvalidOperationException: Formula in a
  cell '$Data'!$C3 is part of a cycle.` — ordinary cycle detection, correctly triggered, because
  `C3` lies inside `B:E`. `C3` holding `=C3+1` throws the identical exception from the identical
  frame (`XLCalcEngine.RecalculateCurrentCell`, via `TryEvaluateSingleCell`), and Excel calls the
  same formula a circular reference. `G10` holding `=A1+B:E` answers `142` with no throw. The
  original note's "different path from `Evaluate`" was right — it is `XLCell.Evaluate`, which runs
  the calculation chain — but the path is doing its job. What remains is a pre-existing question
  unrelated to D38: whether a cycle should throw at all, rather than surface a circular-reference
  result the way Excel does.
- ~~**Whether the same missing intersection affects unary operators.**~~ **Answered 2026-09-01:
  yes** — and so do postfix `%`, comparison and concatenation. Table above.
- Whether `IntersectArguments` should survive the change, or be replaced by the descending context
  entirely. **Still open.**
- **New, still open:** whether the fix should also produce Excel's `#VALUE!` when the formula lies
  outside the operand's span. It should, on the evidence above, but it is a second behaviour the
  naive "take the intersecting cell" fix would miss.

---

## Provenance

Found by the `formula` fuzz target as a timeout, not a crash — the first finding in spec 52 that
came from the time budget rather than from an exception, and an argument for keeping a time budget
in the oracle.

It also recurred unprompted after being recorded, as `VLOOKUP(A1,BB:qX-1,FAESE)` — 396 columns, and
about four minutes. Withholding the seed does not keep the target away from this shape; only fixing
it will.
