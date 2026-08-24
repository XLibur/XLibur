# Spec 30 — Array application gets an interface, and the defect it hides gets fixed

**Area:** Architecture · Refactor · **Correctness (defect)**
**Effort:** S–M (~3 days)
**Dependencies:** None hard. **Overlaps spec 32** in the calc-engine function layer — see Conflicts.
**Status:** Proposed.

## Goal

Fix a dead-store defect that makes every scalar function return the same value for every element of
an array formula, and remove the shape that allowed it: element-wise array application becomes a
module that takes its per-element arguments **as a parameter** instead of reaching for an enclosing
buffer.

## Why this spec exists

Element-wise array application is a private path inside `FunctionDefinition`. It has no interface of
its own, no direct caller, and no test that can reach it without building a whole workbook. A
dead-store bug has lived there since December 2024.

### The defect

`XLibur/Excel/CalcEngine/FunctionDefinition.cs:106-118`:

```csharp
    private ScalarValue EvaluateSingleElement(CalcContext ctx, Span<AnyValue> args, int row, int column)
    {
        var itemArg = new AnyValue[args.Length];
        for (var i = 0; i < itemArg.Length; ++i)
        {
            ref var arg = ref args[i];
            itemArg[i] = IsParameterSingleValue(i)
                ? arg.GetArray()[row, column].ToAnyValue()
                : arg;
        }

        var itemResult = _function(ctx, args);
```

`itemArg` is allocated and filled once per result cell, and **never read**. `_function` is called
with `args` — the broadcast arrays that `NormalizeArguments` (`:66-87`) just produced — not with the
per-element scalars. Every element of the result therefore re-evaluates the function against
identical input, and the result array is the same value repeated.

The scalar coercion downstream is what makes the wrong answer look plausible.
`SignatureAdapter.ToNumber` (`XLibur/Excel/CalcEngine/Functions/SignatureAdapter.cs:1237-1250`) is
handed the whole array and takes its top-left corner:

```csharp
        // When user specifies array as an argument in an array formula for a scalar function, use [0,0]
        if (collection.TryPickT0(out var array, out var reference))
            return array[0, 0].ToNumber(ctx.Culture);
```

That comment at `:1242` documents the *symptom* as if it were the design.

### It reproduces

Measured on the tree at `1b41cadd` with a throwaway probe test, since deleted. Every row is the
value XLibur produces today; Excel's answer is in the last column.

| Formula | Entered over | XLibur today | Excel |
|---|---|---|---|
| `SIGN({-1,2,0})` | `B1:D1` | `-1,-1,-1` | `-1,1,0` |
| `SIGN({-1;2;0})` | `B1:B3` | `-1,-1,-1` | `-1,1,0` |
| `ABS({-5,-6,-7})` | `B1:D1` | `5,5,5` | `5,6,7` |
| `POWER({2,3,4},{1,2,3})` | `B1:D1` | `2,2,2` | `2,9,64` |
| `SIGN(A1:A3)`, `A1:A3` = `-1,2,0` | `B1:B3` | `-1,-1,-1` | `-1,1,0` |
| `UPPER({"a","b","c"})` | `B1:D1` | **throws** | `A,B,C` |

Four things this establishes beyond the prompt for this spec:

1. **It is not orientation-specific.** Row and column arrays are both wrong. The one shape where the
   bug is invisible is a 1×3 array entered over a 3×1 range, because broadcasting then makes every
   element genuinely equal to `[0,0]` — and that is precisely the shape the existing test uses.
2. **It is not literal-specific.** `SIGN(A1:A3)` over real worksheet cells is wrong the same way.
   This is production data, not just array constants.
3. **It is not arity-specific.** `POWER` with two array arguments is wrong in both.
4. **It has a second presentation.** Text-coercing adapters do not silently mis-answer — they
   *throw*. `SignatureAdapter.ToText:1252-1264` has no `[0,0]` branch where `ToNumber` does:

   ```csharp
        if (collection.TryPickT0(out _, out var reference))
            throw new NotImplementedException("Array formulas not implemented.");
   ```

   So `{=UPPER({"a","b","c"})}` throws `NotImplementedException` today. Fixing the mis-call makes
   `ToText` receive a scalar and the formula work. **Expect task 2 to turn some currently-throwing
   behaviour into working behaviour**, not only to change values.

### Provenance — and a correction

The premise handed to this spec was that fork commit `81043e71` ("fix: resolve SonarQube blockers
(#12)", 2026-03-15) introduced the mis-call while extracting `EvaluateSingleElement`. **That is
wrong, and the truth is worse.** Verified by reading the file at three commits:

| Commit | Date | Author | The element call |
|---|---|---|---|
| `819528c9` "Implement array formulas" | 2023-05-04 | jahav (upstream) | `EvaluateFunction(ctx, itemArg)` — **correct** |
| `2124d7cd` | 2024-01-20 | jahav (upstream) | `EvaluateFunction(ctx, itemArg)` — **correct** |
| `fc08037c` "Remove legacy Expression infrastructure" | 2024-12-12 | jahav (upstream) | `_function(ctx, args)` — **broken** |
| `81043e71` "fix: resolve SonarQube blockers (#12)" | 2026-03-15 | fork | `_function(ctx, args)` — carried along |

`fc08037c` deleted the `EvaluateFunction` wrapper and inlined it at both call sites. Its diff on this
file is two adjacent replacements:

```
-            return EvaluateFunction(ctx, args);
+            return _function(ctx, args);
...
-                    var itemResult = EvaluateFunction(ctx, itemArg);
+                    var itemResult = _function(ctx, args);
```

The first is correct — that site really was passing `args`. The second applied the same replacement
text to a site whose argument was `itemArg`. One inlining, two sites, one of them wrong.

So the bug **is** inherited from upstream ClosedXML rather than introduced by the fork, as the
premise said, but by `fc08037c` and not by `81043e71`. What `81043e71` did is the part this spec is
named for: it extracted the inner loop into a private `EvaluateSingleElement` **for testability**,
under a static-analysis remit, and gained no test surface at all — the extracted method is private,
has one caller, and nothing can call it with a chosen argument buffer. The extraction preserved the
defect exactly, and made it harder to see by putting the dead store and its non-use in a method
small enough to read as obviously fine.

**That is this spec's argument: code extracted for testability that gains no test surface preserves
its bugs exactly.** A private method with one caller is not a seam. It is the same code with a name.

### Blast radius

Counted by grep over `XLibur/Excel/CalcEngine/Functions/*.cs`:

| Measure | Count |
|---|---:|
| Registrations carrying `FunctionFlags.Scalar` | **265** |
| …of those, registered through an `Adapt*` wrapper | **261** |
| …registered without one (`RATE`, `NA`, `IFS`, `SWITCH`) | 4 |
| `Adapt*` overloads in `SignatureAdapter.cs` | 61 |

**Correction to the premise.** This spec was briefed with 245 and 241. The real numbers are 265 and
261. 241 is what `grep 'FunctionFlags.Scalar' | grep Adapt` returns, which undercounts by 20: in
`Text.cs` most registrations wrap across two lines with `Adapt(...)` on the first and
`FunctionFlags.Scalar);` on the second, e.g. `Text.cs:46-47`. The delta of 4 in the briefing was
right; the base was 20 low in both terms.

Every one of the 265 is affected inside an array formula or a dynamic-array formula. The exemption
is narrow: `CallAsArray:54` short-circuits the whole element-wise path when
`ReturnsArray && _allowRanges == AllowRange.All`, which is **12 functions** — `ROW`, `TRANSPOSE`,
`MINVERSE`, `MMULT`, `FREQUENCY`, `GROWTH`, `LINEST`, `LOGEST`, `TREND`, `MODE.MULT`, `T`,
`TEXTSPLIT`. `INDEX` is *not* exempt: it is `ReturnsArray` but `AllowRange.Only`.

### The one test that covers this pins the wrong answer

`XLibur.Tests/Excel/CalcEngine/ArrayFormulaCalculationTests.cs:120-131`:

```csharp
    [Test]
    public async Task Array_argument_for_scalar_function_in_array_formula_uses_only_first_value_of_array()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Range("B1:B3").FormulaArrayA1 = "SIGN({-1,2,0})";

        // Uses only -1 for all values
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(-1);
        await Assert.That(ws.Cell("B2").Value).IsEqualTo(-1);
        await Assert.That(ws.Cell("B3").Value).IsEqualTo(-1);
    }
```

A 1×3 array entered over a 3×1 range — the single shape in which the defect produces the same answer
as correct broadcasting would. The test name asserts the buggy rule as a rule, and the comment
explains it.

**Its name and its assertions both have to change, and that is not weakening a test.** It is
correcting a test that pinned a defect. A test whose name states an incorrect behaviour is worse than
no test: it stops the next reader from filing the bug.

## Non-goals

- **Not touching `SignatureAdapter.cs`.** That is spec 32, written in parallel today. Its `ToNumber`
  `[0,0]` branch and its `ToText` throw are *reached* because of this defect, but once the correct
  per-element scalars arrive neither branch is on the array path any more. Leave both in place —
  they are still reachable from a non-array formula with a range argument.
- **Not touching the dependency tree or recalculation policy.** Spec 04 owns those.
- **Not implementing LET/LAMBDA.** Spec 08.
- **Not implementing grid spilling.** `DynamicArray.cs:14` records that as separate work.
- **Not a performance spec.** The fix removes an allocation, and task 6 measures that it did not cost
  anything — but no performance claim goes in a PR description without BenchmarkDotNet.
- **No public API change.** `FunctionDefinition` is `internal sealed`.

## Current state

Verified against the tree at `1b41cadd` (2026-08-24). Every line number below was read, not inherited.

- `XLibur/Excel/CalcEngine/FunctionDefinition.cs` — 173 lines, `internal sealed class`
  - `CallFunction` — `:41-47`, public entry point, scalar semantics
  - `CallAsArray` — `:52-60`, public entry point, array semantics; short-circuit at `:54`
  - `NormalizeArguments` — `:66-87`, private; broadcasts single-valued params to the result shape
  - `EvaluateArrayElements` — `:92-104`, private; the `row`/`column` double loop
  - `EvaluateSingleElement` — `:106-124`, private; **the defect, at `:117`**
  - `IntersectArguments` — `:126-141`, private
  - `GetScalarArgsMaxSize` — `:143-159`, private
  - `IsParameterSingleValue` — `:161-172`, private
- `XLibur/Excel/CalcEngine/CalculationVisitor.cs`
  - `_argsPool` — `:13`, built at `:18` as `ArrayPool<AnyValue>.Create(XLConstants.MaxFunctionArguments, 100)`
  - `Visit(CalcContext, FunctionNode)` — `:74-95`; rents at `:80`, spans at `:81`, returns in
    `finally` at `:93`
  - The mode pick — `:87-89`:
    ```csharp
            return !context.IsArrayCalculation
                ? fn!.CallFunction(context, args)
                : fn!.CallAsArray(context, args);
    ```
- `XLibur/Excel/CalcEngine/XLFunctionLibrary.cs:105` — a **third** caller of `CallFunction`, from
  spec 13's public grid-free function library. Easy to miss; the entry-point collapse must handle it.
- `XLibur/Excel/CalcEngine/Functions/DynamicArray.cs:11` — a doc comment naming
  `FunctionDefinition.CallAsArray` in prose. Renaming the entry point makes it stale.
- `XLibur/Excel/CalcEngine/XLCalcEngine.cs:477` — the **only** place `IsArrayCalculation` is set
  true, inside `EvaluateArrayFormula` (`:473-485`). Reached from three sites: `:286` and `:408`
  (`FormulaType.Array`, i.e. CSE array formulas) and `:506` (`SpillDynamicArray`). This is why both
  array *and* dynamic-array formulas are affected.
- `XLibur/Excel/CalcEngine/XLCalcEngine.cs:752` —
  `internal delegate AnyValue CalcEngineFunction(CalcContext ctx, Span<AnyValue> arg);`. The
  parameter is `Span`, not `ReadOnlySpan`, which constrains the new module's signature.
- `XLibur.Tests/Excel/CalcEngine/ArrayFormulaCalculationTests.cs` — 132 lines, 8 tests; `:120-131` is
  the one that pins the defect. The other 7 exercise result *shaping* (broadcast, clipping,
  `TRANSPOSE`) and should be unaffected.
- Array-formula tests live in 12 files. Triage surface for task 2:
  `ArrayFormulaCalculationTests.cs` (8 tests), `ArrayFormulaTests.cs` (23),
  `ArrayShapingTests.cs` (35), `DynamicArrayFunctionTests.cs` (11), plus `HypothesisTestTests.cs`,
  `LookupTests.cs`, `ModernTextTests.cs`, `RegressionTests.cs`, `FormulaTests.cs`,
  `BatchRowDeleteTests.cs`, `FormulaShiftFilterTests.cs`, `XLRangeBaseTests.cs`.

## File structure

```
XLibur/Excel/CalcEngine/ElementApplication.cs          new — the value-in/value-out module
XLibur/Excel/CalcEngine/FunctionDefinition.cs          modified — defect fixed; two entry points -> one
XLibur/Excel/CalcEngine/CalculationVisitor.cs          modified — one call site instead of a ternary
XLibur/Excel/CalcEngine/XLFunctionLibrary.cs           modified — third caller follows the rename
XLibur/Excel/CalcEngine/Functions/DynamicArray.cs      modified — stale doc reference
XLibur.Tests/Excel/CalcEngine/ArrayElementApplicationTests.cs   new — the failing test, then the gate
XLibur.Tests/Excel/CalcEngine/ArrayFormulaCalculationTests.cs   modified — corrected, not weakened
```

## The design

Three changes, and only the first one changes behaviour.

**1. Pass the per-element buffer.** `_function(ctx, args)` becomes `_function(ctx, itemArg)`. One
line.

**2. Make the wrong buffer inexpressible.** The reason `args` was reachable is that
`EvaluateSingleElement` is a method on a class that also holds `args` in scope from its caller. A
module that receives only the per-element arguments cannot pick up the wrong ones, because the wrong
ones are not there:

```csharp
namespace XLibur.Excel.CalcEngine;

/// <summary>
/// Applies one function to one element of an array-formula result.
/// </summary>
/// <remarks>
/// Value in, value out: the per-element arguments arrive as a parameter and nothing else is in
/// scope. That is the point of the type. Between 2024-12 and spec 30 the equivalent code was a
/// private method beside the broadcast argument buffer, and it called the function with that buffer
/// instead of the per-element one — every cell of every array formula over a scalar function got the
/// same answer. A module that cannot see the broadcast buffer cannot pass it.
/// </remarks>
internal readonly struct ElementApplication
{
    private readonly CalcEngineFunction _function;

    internal ElementApplication(CalcEngineFunction function) => _function = function;

    /// <summary>
    /// Calls the function with the arguments for a single element.
    /// </summary>
    /// <param name="ctx">Evaluation context.</param>
    /// <param name="argsForElement">
    /// Exactly the arguments for this element. Single-valued parameters are the scalar at this
    /// element's position; multi-valued parameters are the whole argument.
    /// </param>
    /// <returns>
    /// The element's value. If the function returns an array, only its top-left value is used —
    /// per the FILTERXML tests.
    /// </returns>
    internal ScalarValue Apply(CalcContext ctx, Span<AnyValue> argsForElement)
    {
        var itemResult = _function(ctx, argsForElement);

        return itemResult.TryPickSingleOrMultiValue(out var scalarResult, out var arrayResult, ctx)
            ? scalarResult
            : arrayResult![0, 0];
    }
}
```

`Span<AnyValue>` rather than `ReadOnlySpan<AnyValue>` because `CalcEngineFunction`
(`XLCalcEngine.cs:752`) takes a `Span`. That is a constraint from the existing delegate, not a
choice; note it in the remarks so nobody "tightens" it and finds it does not compile.

**3. One entry point instead of two.** `CallFunction` and `CallAsArray` differ only in which
semantics they apply, and `CalculationVisitor:87-89` picks between them with a ternary on a boolean.
Make the mode the parameter it already is:

```csharp
/// <summary>How a function call is evaluated.</summary>
internal enum CallMode
{
    /// <summary>Ordinary cell formula. Implicit intersection applies; one call, one result.</summary>
    Scalar,

    /// <summary>
    /// Array-formula semantics. Arguments are broadcast to a common shape and the function is
    /// applied once per element, unless it declares that it consumes and returns whole arrays.
    /// </summary>
    Array,
}
```

```csharp
    public AnyValue Call(CalcContext ctx, Span<AnyValue> args, CallMode mode)
```

and the caller becomes one line:

```csharp
            return fn!.Call(context, args,
                context.IsArrayCalculation ? CallMode.Array : CallMode.Scalar);
```

**Buffer reuse.** `EvaluateSingleElement` allocates `new AnyValue[args.Length]` per result cell
today, which is pure waste — the array is filled and discarded without being read. The fix hoists one
buffer to `EvaluateArrayElements` and reuses it across all `totalRows × totalColumns` elements. This
**removes** an allocation rather than adding one; it does not introduce a new one where the current
code was allocation-free. Do not size it from `_argsPool` — `CalculationVisitor` owns that pool and
returns it in a `finally`, and a second renter inside the element loop would be a second lifetime to
reason about for no gain.

## Global constraints

- Warnings are errors (`TreatWarningsAsErrors=true`); nullable enabled. New code must be
  null-annotated.
- Branch per spec; never commit to main. Commit prefixes `refactor:` / `fix:` / `test:` / `perf:`.
- No compound shell commands (`&&`, `||`, `;`) in agent tool calls.
- **Do not use `sed -i` on tracked files.** `.gitattributes` checks out CRLF and Git Bash's `sed -i`
  rewrites the file as LF, turning a one-line change into a whole-file diff. Use the Edit/Write
  tools; verify with `git diff --numstat` — a file whose changed-line count is near its total line
  count was rewritten, not edited.
- Test filtering uses `--treenode-filter`, never `--filter`. Exit 5 = invalid option; exit 8 = zero
  tests matched. Never filter at solution level — name the `.csproj`.
- Pass `-f net10.0` for iteration; run without it before opening the PR.
- Build: `dotnet build XLibur/XLibur.csproj -c Release -v q`
- Tests: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
- Tests are TUnit and assertions are awaitable: `await Assert.That(actual).IsEqualTo(expected)`. A
  missing `await` silently passes. `[Test]`, `[Arguments(...)]`, `[MethodDataSource(...)]`. The suite
  is serial (`[assembly: NotInParallel]`).
- **TUnit analyzer gotcha:** `Assert.That(true).IsTrue()` fails the build with
  `TUnitAssertions0005: Assert.That(...) should not be used with a constant value`. Probe tests that
  only want to print a value must assert on the value itself, not on a constant.

## Work plan

| # | Task | Size | Gate |
|---|---|---|---|
| 1 | Failing test across several array shapes | S | **Red on purpose** — 6 of 6 cases fail |
| 2 | Fix the mis-call; triage the whole suite | **M — the risky one** | Task 1 green; every other failure triaged |
| 3 | Correct the tests that pinned the old behaviour | S | Suite green |
| 4 | Extract `ElementApplication` | S | Suite green; `args` out of scope at the call |
| 5 | Collapse the two entry points into `Call(ctx, args, mode)` | S | Suite green; 3 call sites updated |
| 6 | Confirm no allocation or time regression | S | Within BenchmarkDotNet noise |

Task 2 is where the risk is. Task 1 exists to size it before taking it.

---

### Task 1 — The failing test

Write the test that proves the defect, and land it **red**. It is not a mistake in the branch
history; it is the record that the defect existed and what it looked like.

**Files:**
- Create: `XLibur.Tests/Excel/CalcEngine/ArrayElementApplicationTests.cs`

**Interfaces:**
- Produces: `Each_element_of_an_array_formula_is_evaluated_against_its_own_argument`, the gate for
  tasks 2 through 5.

- [ ] **Step 1: Write the test**

```csharp
using XLibur.Excel;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.CalcEngine;

/// <summary>
/// A scalar function inside an array formula must be applied element by element. Between upstream
/// commit fc08037c (2024-12-12) and spec 30 it was not: the broadcast argument buffer was passed to
/// the function instead of the per-element one, so every result cell re-evaluated against [0,0].
/// These cases cover both orientations, one and two array arguments, worksheet references as well as
/// array literals, and a text-coercing function — which did not mis-answer but threw.
/// </summary>
public class ArrayElementApplicationTests
{
    [Test]
    [Arguments("SIGN({-1,2,0})", "B1:D1", "-1|1|0")]
    [Arguments("SIGN({-1;2;0})", "B1:B3", "-1|1|0")]
    [Arguments("ABS({-5,-6,-7})", "B1:D1", "5|6|7")]
    [Arguments("POWER({2,3,4},{1,2,3})", "B1:D1", "2|9|64")]
    public async Task Each_element_of_an_array_formula_is_evaluated_against_its_own_argument(
        string formula, string target, string expected)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var range = ws.Range(target);

        range.FormulaArrayA1 = formula;

        var actual = string.Join("|", range.Cells().Select(c => c.Value.ToString()));
        await Assert.That(actual).IsEqualTo(expected);
    }

    /// <summary>
    /// The same defect, reached through worksheet cells rather than an array literal — which is what
    /// makes it a production defect and not a curiosity about array constants.
    /// </summary>
    [Test]
    public async Task A_range_argument_is_evaluated_element_by_element()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").Value = -1;
        ws.Cell("A2").Value = 2;
        ws.Cell("A3").Value = 0;

        ws.Range("B1:B3").FormulaArrayA1 = "SIGN(A1:A3)";

        await Assert.That(ws.Cell("B1").Value).IsEqualTo(-1);
        await Assert.That(ws.Cell("B2").Value).IsEqualTo(1);
        await Assert.That(ws.Cell("B3").Value).IsEqualTo(0);
    }

    /// <summary>
    /// SignatureAdapter.ToText has no [0,0] fallback where ToNumber does, so a text function handed
    /// the whole broadcast array threw NotImplementedException instead of returning a wrong value.
    /// Once each element gets its own scalar, ToText never sees an array on this path.
    /// </summary>
    [Test]
    public async Task A_text_function_in_an_array_formula_does_not_throw()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        ws.Range("B1:D1").FormulaArrayA1 = "UPPER({\"a\",\"b\",\"c\"})";

        await Assert.That(ws.Cell("B1").Value).IsEqualTo("A");
        await Assert.That(ws.Cell("C1").Value).IsEqualTo("B");
        await Assert.That(ws.Cell("D1").Value).IsEqualTo("C");
    }
}
```

Add `using System.Linq;` for `Select`. If `c.Value.ToString()` formats a double as `-1` on one
framework and `-1.0` on another, switch the parameterised test to explicit per-cell assertions like
the other two — do not add a culture shim, the suite is already pinned to en-US by
`TestDefaults.ApplyCulture`.

- [ ] **Step 2: Run it and confirm it is red for the right reason**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/ArrayElementApplicationTests/*"`

Expected: **FAIL, 6 of 6.** The four parameterised cases report `-1|-1|-1`, `-1|-1|-1`, `5|5|5`,
`2|2|2`. `A_range_argument_is_evaluated_element_by_element` reports `-1` for all three cells.
`A_text_function_in_an_array_formula_does_not_throw` reports
`NotImplementedException: Array formulas not implemented.`

**If any case passes**, the defect is narrower than this spec claims for that shape. Record which and
why before continuing — the blast-radius table above is a premise, and a case that already passes is
evidence against it.

- [ ] **Step 3: Commit red, and say so**

```bash
git add XLibur.Tests/Excel/CalcEngine/ArrayElementApplicationTests.cs
git commit -m 'test(calc): failing test for element-wise array application - RED, fixed in the next commit (spec 30 task 1)'
```

---

### Task 2 — Fix the mis-call, then triage everything it moves

**This is the risky task.** 265 scalar functions change behaviour under array semantics. Some of the
suite has been written against the broken behaviour without anyone noticing, exactly as
`ArrayFormulaCalculationTests:120` was. Every failure has to be triaged, and the triage is the
deliverable.

**Files:**
- Modify: `XLibur/Excel/CalcEngine/FunctionDefinition.cs:117`

- [ ] **Step 1: The one-line fix**

At `:117`, replace `_function(ctx, args)` with `_function(ctx, itemArg)`:

```csharp
        // itemArg, not args: args holds the broadcast whole-array arguments that NormalizeArguments
        // produced. Calling with those re-evaluates every element against [0,0] and fills the result
        // with one repeated value. Upstream fc08037c inlined EvaluateFunction at two call sites and
        // applied the same replacement text to both; this one took itemArg and got args.
        var itemResult = _function(ctx, itemArg);
```

Verify the edit did not rewrite the file:

Run: `git diff --numstat XLibur/Excel/CalcEngine/FunctionDefinition.cs`
Expected: a small changed-line count against a 173-line file — not ~173.

- [ ] **Step 2: Task 1's test goes green**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/ArrayElementApplicationTests/*"`
Expected: PASS, 6 of 6.

- [ ] **Step 3: Run the array-formula suites first, then the whole thing**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/ArrayFormulaCalculationTests/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/ArrayShapingTests/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/ArrayFormulaTests/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/DynamicArrayFunctionTests/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`

**Expect failures here. That is the point of the task.** At minimum
`Array_argument_for_scalar_function_in_array_formula_uses_only_first_value_of_array` fails — it is
the known one, and task 3 fixes it.

- [ ] **Step 4: Triage every failure into one of three buckets**

For each failing test, decide and write down which it is:

| Bucket | Meaning | Action |
|---|---|---|
| **A — pinned the bug** | The test asserts a value that is only correct if elements share `[0,0]`. Check the expected value against Excel. | Task 3 corrects it. Record the old and new expectation. |
| **B — newly reachable** | The test previously threw or errored and now produces a value, or vice versa. `ToText` cases land here. | Verify against Excel, then assert the correct value. |
| **C — something else** | Anything that is not A or B. | **Stop.** |

**Bucket C is a finding and it ends the task.** If a test fails for a reason that is *not* "it pinned
the bug" and *not* "it now reaches code it could not reach before", the fix has a genuine downstream
dependency this spec did not anticipate. Do not adjust the test to make it pass. Record the test, the
formula, the before and after values, and the reasoning, and report before continuing. A real
dependency on the broken semantics is a more valuable result than the rest of this spec.

Write the triage table into this spec file under a `## Results` heading as you go — not afterwards
from memory.

- [ ] **Step 5: Commit the fix on its own, with the triage**

Commit the source fix separately from the test corrections, so `git log` shows the one line that
changed behaviour.

```bash
git add XLibur/Excel/CalcEngine/FunctionDefinition.cs
git commit -m 'fix(calc): apply array-formula elements to their own arguments, not the broadcast buffer (spec 30 task 2)'
```

```bash
git add docs/specs/30-array-application-seam.md
git commit -m 'docs(specs): record the task 2 triage for spec 30'
```

---

### Task 3 — Correct the tests that pinned the old behaviour

**Files:**
- Modify: `XLibur.Tests/Excel/CalcEngine/ArrayFormulaCalculationTests.cs:120-131`
- Modify: every bucket-A and bucket-B test from task 2 step 4

**Interfaces:**
- None. Test-only.

- [ ] **Step 1: Rewrite the known one**

Both the name and the assertions change. The name is the more important half — it currently states
the defect as a rule.

```csharp
    /// <summary>
    /// An array argument to a scalar function is applied element by element, not collapsed to its
    /// first value. Until spec 30 this test asserted the opposite and was named for it: the
    /// broadcast argument buffer was passed to the function instead of the per-element one. The
    /// shape below is the one shape where the two agree — a 1x3 array over a 3x1 range broadcasts to
    /// -1 everywhere — which is why the defect survived here for 20 months. The second case is the
    /// same formula over a 1x3 range, where they do not agree.
    /// </summary>
    [Test]
    public async Task Array_argument_for_scalar_function_is_applied_element_by_element()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();

        // 1x3 array broadcast down a 3x1 range: every element really is -1.
        ws.Range("B1:B3").FormulaArrayA1 = "SIGN({-1,2,0})";
        await Assert.That(ws.Cell("B1").Value).IsEqualTo(-1);
        await Assert.That(ws.Cell("B2").Value).IsEqualTo(-1);
        await Assert.That(ws.Cell("B3").Value).IsEqualTo(-1);

        // Same array across a 1x3 range: each element gets its own value.
        ws.Range("D1:F1").FormulaArrayA1 = "SIGN({-1,2,0})";
        await Assert.That(ws.Cell("D1").Value).IsEqualTo(-1);
        await Assert.That(ws.Cell("E1").Value).IsEqualTo(1);
        await Assert.That(ws.Cell("F1").Value).IsEqualTo(0);
    }
```

Keeping the original 3×1 case matters: it pins that broadcasting a 1×3 down three rows still yields
`-1` three times, which is correct and is *not* what the fix changed.

- [ ] **Step 2: Correct the rest of the bucket-A and bucket-B tests**

For each: check the expected value against Excel, change the assertion to the correct value, and
update the summary to say what changed and why. Do not delete a test to make the suite green.

- [ ] **Step 3: Full suite, both frameworks**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Expected: PASS on net8.0 and net10.0.

- [ ] **Step 4: Confirm no test name still asserts the defect**

Run: `grep -rn 'uses_only_first_value\|only the first value\|Uses only' XLibur.Tests/Excel/CalcEngine`
Expected: no hit that refers to array-formula element application. `Median_uses_only_numbers` in
`StatisticalTests.cs:790` is unrelated and stays.

- [ ] **Step 5: Commit**

```bash
git add XLibur.Tests/Excel/CalcEngine
git commit -m 'test(calc): correct the tests that pinned the array-element defect (spec 30 task 3)'
```

---

### Task 4 — Extract `ElementApplication`

The defect is fixed. This task removes the shape that allowed it.

**Files:**
- Create: `XLibur/Excel/CalcEngine/ElementApplication.cs`
- Modify: `XLibur/Excel/CalcEngine/FunctionDefinition.cs:92-124`

**Interfaces:**
- Produces: `ElementApplication.Apply(CalcContext, Span<AnyValue>) → ScalarValue`.

- [ ] **Step 1: Create the module**

Use the type from "The design" above verbatim.

- [ ] **Step 2: Rewrite `EvaluateArrayElements` to own the buffer and build the per-element args**

`EvaluateSingleElement` disappears. Its argument-building loop moves into `EvaluateArrayElements`,
which is the only place that can see `args`, and the buffer is allocated once instead of once per
cell:

```csharp
    /// <summary>
    /// Applies the function to every element of the result shape.
    /// </summary>
    /// <remarks>
    /// The per-element buffer is allocated once and refilled, not once per element. It is
    /// deliberately not rented from <c>CalculationVisitor</c>'s pool: that pool's lifetime is the
    /// enclosing <c>Visit</c> call and a second renter inside this loop would be a second lifetime to
    /// reason about for no gain.
    /// </remarks>
    private AnyValue EvaluateArrayElements(CalcContext ctx, Span<AnyValue> args, int totalRows, int totalColumns)
    {
        var application = new ElementApplication(_function);
        var result = new ScalarValue[totalRows, totalColumns];
        var itemArgs = new AnyValue[args.Length];

        for (var row = 0; row < totalRows; ++row)
        {
            for (var column = 0; column < totalColumns; ++column)
            {
                for (var i = 0; i < itemArgs.Length; ++i)
                {
                    ref var arg = ref args[i];
                    itemArgs[i] = IsParameterSingleValue(i)
                        ? arg.GetArray()[row, column].ToAnyValue()
                        : arg;
                }

                result[row, column] = application.Apply(ctx, itemArgs);
            }
        }

        return new ConstArray(result);
    }
```

`application.Apply(ctx, itemArgs)` — `args` is in scope here, so this is not yet the full guarantee.
The guarantee is inside `ElementApplication`: `Apply` has no access to anything but what it is given,
so the *function call itself* can no longer receive the wrong buffer. That is the line the defect
was on.

- [ ] **Step 3: Confirm `EvaluateSingleElement` is gone**

Run: `grep -n 'EvaluateSingleElement' XLibur/Excel/CalcEngine/FunctionDefinition.cs`
Expected: no output.

- [ ] **Step 4: Build and run the suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/CalcEngine/ElementApplication.cs XLibur/Excel/CalcEngine/FunctionDefinition.cs
git commit -m 'refactor(calc): extract ElementApplication so the per-element buffer is a parameter (spec 30 task 4)'
```

---

### Task 5 — One entry point

**Files:**
- Modify: `XLibur/Excel/CalcEngine/FunctionDefinition.cs:41-60`
- Modify: `XLibur/Excel/CalcEngine/CalculationVisitor.cs:87-89`
- Modify: `XLibur/Excel/CalcEngine/XLFunctionLibrary.cs:105`
- Modify: `XLibur/Excel/CalcEngine/Functions/DynamicArray.cs:11`

**Interfaces:**
- Produces: `FunctionDefinition.Call(CalcContext, Span<AnyValue>, CallMode) → AnyValue`.
- Removes: `CallFunction`, `CallAsArray`.

- [ ] **Step 1: Add `CallMode`**

Put the enum in `ElementApplication.cs` — it exists to describe this dispatch and does not need a
file of its own. Use the definition from "The design".

- [ ] **Step 2: Collapse the two entry points**

```csharp
    /// <summary>
    /// Evaluates the function.
    /// </summary>
    /// <param name="ctx">Evaluation context.</param>
    /// <param name="args">The evaluated arguments. Modified in place by both modes.</param>
    /// <param name="mode">Scalar or array-formula semantics.</param>
    public AnyValue Call(CalcContext ctx, Span<AnyValue> args, CallMode mode)
    {
        if (mode == CallMode.Scalar)
        {
            if (CalcContext.UseImplicitIntersection)
                IntersectArguments(ctx, args);

            return _function(ctx, args);
        }

        // Functions that both consume and return whole arrays are called once with everything.
        // Twelve functions qualify: ROW, TRANSPOSE, MINVERSE, MMULT, FREQUENCY, GROWTH, LINEST,
        // LOGEST, TREND, MODE.MULT, T, TEXTSPLIT. INDEX is ReturnsArray but AllowRange.Only, so it
        // goes down the element-wise path like everything else.
        if (_flags.HasFlag(FunctionFlags.ReturnsArray) && _allowRanges == AllowRange.All)
            return _function(ctx, args);

        var (totalRows, totalColumns) = GetScalarArgsMaxSize(args);
        NormalizeArguments(ctx, args, totalRows, totalColumns);
        return EvaluateArrayElements(ctx, args, totalRows, totalColumns);
    }
```

- [ ] **Step 3: Update all three call sites**

`CalculationVisitor.cs:87-89`:

```csharp
            return fn!.Call(context, args,
                context.IsArrayCalculation ? CallMode.Array : CallMode.Scalar);
```

`XLFunctionLibrary.cs:105` — the grid-free library always evaluates with scalar semantics:

```csharp
            value = definition.Call(context, args.AsSpan(), CallMode.Scalar);
```

`DynamicArray.cs:11` — the prose reference goes stale:

```csharp
/// engine (see <c>FunctionDefinition.Call</c> with <c>CallMode.Array</c>) uses their whole array
```

Confirm nothing else names either method:

Run: `grep -rn 'CallAsArray\|CallFunction' XLibur XLibur.Tests --include=*.cs`
Expected: no output.

- [ ] **Step 4: Build and run the full suite on both frameworks**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Expected: PASS on net8.0 and net10.0.

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/CalcEngine XLibur/Excel/CalcEngine/Functions/DynamicArray.cs
git commit -m 'refactor(calc): one Call entry point with an explicit mode (spec 30 task 5)'
```

---

### Task 6 — Confirm no regression

The fix calls the function `totalRows × totalColumns` times where it previously called it the same
number of times with the same arguments, so the call count is unchanged. What changes: one fewer
`AnyValue[]` allocation per result cell, and the coercion inside each `Adapt*` wrapper now sees a
scalar rather than an array, which is the cheaper branch of `ToNumber`.

Both of those point the same way, but **the repo's ground rules do not allow a performance claim
without BenchmarkDotNet**, and the benchmark machine has ~40% run-to-run timing variance. Take three
runs and compare medians. Allocation numbers are stable and are the more trustworthy signal here.

- [ ] **Step 1: Find or add a formula-heavy benchmark**

Run: `dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --list flat`

Pick the benchmark that exercises formula evaluation. If none covers array formulas specifically,
this task's honest answer is "the general formula benchmark is unchanged, and array formulas are not
benchmarked" — say that rather than inventing a number.

- [ ] **Step 2: Measure the merge-base, three runs**

- [ ] **Step 3: Measure the branch, three runs**

- [ ] **Step 4: Compare medians and record**

Expected: allocated bytes down or flat; time within noise.

**Decision rule.** A median time regression above the machine's ~40% noise floor must be explained
before this spec lands. An allocation *increase* of any size must be explained — the change is
supposed to remove one. Record the numbers in a Results section either way, including "not
measurable" if that is the outcome.

- [ ] **Step 5: Commit the Results section**

```bash
git add docs/specs/30-array-application-seam.md
git commit -m 'docs(specs): record the benchmark numbers for spec 30'
```

---

## Acceptance criteria

1. `_function(ctx, args)` appears **exactly twice** in `FunctionDefinition.cs` — the scalar path and
   the whole-array short-circuit — and neither is inside a per-element loop. Gate:
   `grep -c '_function(ctx, args)' XLibur/Excel/CalcEngine/FunctionDefinition.cs` returns `2`.
2. `EvaluateSingleElement` no longer exists. Gate:
   `grep -rn 'EvaluateSingleElement' XLibur --include=*.cs` returns nothing.
3. No `AnyValue[]` is allocated inside the element loop. Gate: `new AnyValue[` appears at most once
   in `EvaluateArrayElements`, outside both `for` loops.
4. `ElementApplication.Apply` receives the per-element arguments as a parameter and closes over no
   argument buffer. Gate: `ElementApplication.cs` contains no field other than `_function`.
5. `CallFunction` and `CallAsArray` no longer exist; one `Call` takes a `CallMode`. Gate:
   `grep -rn 'CallAsArray\|CallFunction' XLibur XLibur.Tests --include=*.cs` returns nothing.
6. All three former call sites — `CalculationVisitor.cs:87`, `XLFunctionLibrary.cs:105`, and the
   `DynamicArray.cs:11` doc comment — name the new entry point.
7. Task 1's six cases pass, with no assertion weakened from the values in the reproduction table.
8. No test name or comment in `XLibur.Tests/Excel/CalcEngine/` asserts that a scalar function uses
   only the first value of an array argument.
9. Every test that failed at task 2 is triaged into bucket A, B or C in a Results section, with its
   old and new expected values. Any bucket-C finding is recorded whether or not it is resolved here.
10. Full suite green on net8.0 and net10.0.
11. No public API change. `FunctionDefinition`, `ElementApplication` and `CallMode` are all
    `internal`.
12. Benchmark medians recorded from three runs per side, or "not measurable" stated explicitly.

## Conflicts

Assessed against `docs/specs/README.md` and the open specs listed there.

- **Spec 32 (collapse the 61-overload registration interface) — a real overlap, not a file
  collision.** 32 touches `SignatureAdapter.cs` and `FunctionRegistry.cs`; 30 touches
  `FunctionDefinition.cs`, `CalculationVisitor.cs`, `ElementApplication.cs` and
  `XLFunctionLibrary.cs`. **No file is shared.** The overlap is semantic: `FunctionDefinition` reads
  `_allowRanges` and `_markedParams` in `IsParameterSingleValue` (`:161-172`), `GetScalarArgsMaxSize`
  (`:143-159`) and `IntersectArguments` (`:126-141`), and those two fields are exactly what 32
  replaces. If 32 changes how a parameter declares itself single- or multi-valued, all three methods
  change with it — and `IsParameterSingleValue` is what decides which arguments get indexed per
  element, so it sits directly on the path this spec fixes.

  **Recommended order: 30 first.** It is three days, it is a correctness fix, and it is small enough
  to rebase. 32 is a 441-call-site sweep, and a sweep that large should land on corrected semantics —
  otherwise every one of those call sites is re-verified against a function-application path that is
  about to change meaning. Running 32 first would also mean 30's task 2 triage runs against a
  registration layer that just moved, making bucket C impossible to distinguish from 32 fallout.

  Verified independently: `SignatureAdapter.cs` does contain **61** `static CalcEngineFunction
  Adapt*` overloads (31 `Adapt`, 14 `AdaptLastOptional`, 7 `AdaptLastTwoOptional`, and 9 one-off
  named variants), which corroborates 32's headline number.

- **Spec 04 (demand-driven formula evaluation) — check before starting, low risk.** 04 owns
  `CalcContext` and the evaluation stack. 30 touches `CalculationVisitor.cs` at exactly one place,
  `:87-89`, replacing a two-branch ternary with a one-line call, and it reads `CalcContext` for
  `IsArrayCalculation` (`:78`) and the static `UseImplicitIntersection` without changing either. If
  04 rewrites `Visit(CalcContext, FunctionNode)` — which is plausible, since that is where the
  argument pool is rented — the two collide in a ~20-line method. **30 first**, for the same reason:
  it is days and 04 is an L.

- **Spec 08 (LET / LAMBDA)** shares 04's territory in the engine core and is listed against it in the
  README's conflict map (`04↔08`). It does not touch `FunctionDefinition`. No conflict with 30.

- **Specs 22, 23, 24, 25** are all in the IO, chart and style layers. File-disjoint from 30.

- **Spec 13** contributed `XLFunctionLibrary.cs` and is done. 30 changes one line of it, `:105`, from
  `CallFunction` to `Call(..., CallMode.Scalar)`. This is an internal call behind 13's public
  surface; 13's public API is unaffected.
