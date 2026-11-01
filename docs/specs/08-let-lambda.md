# Spec 08 — LET, LAMBDA, and Lambda-Helper Functions

**Area:** Feature (calc engine core)
**Effort:** L (2–4 weeks; parser/scoping work, not bulk registration)
**Dependencies:** Independent of Spec 07, but land LET before LAMBDA. Interacts with the spill engine (`docs/dynamic-array-spill-phase-b.md`).
**Status:** In progress — tasks 1–5 done (parser spike, LET, `_xlpm.` on save, LAMBDA); the helpers (task 6) next. See [Results](#results).

## Summary

`LET` (named intermediate values) and `LAMBDA` (user-defined functions) plus the helper set (`MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY, ISOMITTED`) are the most-requested modern Excel capabilities not coverable by simple function registration: they introduce **lexical scoping and function values** into a calc engine that currently has neither. Files using LET/LAMBDA today evaluate to errors in XLibur.

## Current state

- Evaluation: `XLibur/Excel/CalcEngine/CalculationVisitor.cs` (315 lines) walks the AST from ClosedXML.Parser; values are `AnyValue`/`ScalarValue` discriminated unions (`AnyValue.cs` 680 lines — cases: logical, number, text, error, Array, Reference).
- Context: `CalcContext` carries workbook/sheet/address state per evaluation.
- Functions dispatch through `FunctionRegistry` + `SignatureAdapter` — args are eagerly evaluated before the function body runs. **LET/LAMBDA cannot use this path**: LET binds names sequentially with later value-exprs referencing earlier names; LAMBDA's body must not evaluate at definition time.
- Parser: ClosedXML.Parser 2.0 — **task 1 is verifying** it parses `LET(x, 1, x+1)` and `LAMBDA(x, x*2)(3)` (nested-call syntax where a function call's target is itself an expression) and how `_xlpm.`-prefixed parameter names surface in the AST. Excel stores lambda parameters as `_xlpm.x` and wraps names in `_xlfn.LET`/`_xlfn.LAMBDA` — check `FormulaConverter`/name-handling for how `_xlfn.` prefixes are already handled for other modern functions.

## Design

### Phase 1 — LET

1. **Special-form evaluation.** Add a special-cases hook in `CalculationVisitor` for LET (bypass `SignatureAdapter` eager evaluation): evaluate value-exprs left to right, binding each name in an environment, then evaluate the body expression in that environment.
2. **Environment.** A small immutable-ish scope chain on `CalcContext`: `sealed class XLNameEnvironment { XLNameEnvironment? Parent; Dictionary<string, AnyValue> Bindings; }` — lookups walk the chain; name resolution in the visitor consults the environment **before** defined-names/sheet names (Excel gives LET names precedence within scope). Case-insensitive comparer matching Excel name rules.
3. **Name references in the AST**: LET names arrive as name/identifier nodes (likely with `_xlpm.` prefix from files, bare when typed via API). Normalize both.
4. Round-trip: formulas containing LET must save/load verbatim (they already round-trip as strings — verify `_xlfn./_xlpm.` prefixes are preserved on save and hidden in the user-facing `Formula.A1` the same way existing `_xlfn` functions are handled).

### Phase 2 — LAMBDA + invocation

1. **Function values.** Add a `Lambda` case to `AnyValue` (parameters: `IReadOnlyList<string>`, body AST node reference, captured `XLNameEnvironment`). This is the deep cut — audit every `Match`/`TryPick` switch over `AnyValue` (compiler will find them once the case is added; the union is hand-rolled so follow the existing pattern for adding a case, see how `Array` is threaded through).
2. **Definition:** LAMBDA as a special form returns the function value without evaluating the body. A LAMBDA evaluated as a cell's final result yields `#CALC!` (Excel behavior) unless invoked.
3. **Invocation:** support call-on-expression `LAMBDA(...)(args)` and — the common storage form — lambdas bound via LET or **defined names** (a defined name whose refersTo is `=LAMBDA(...)` acts as a UDF). Defined-name lambda invocation is the highest-value entry point: check how `DefinedNames` resolution feeds the evaluator and allow a name to resolve to a lambda value that a call node can apply.
4. **Recursion:** named lambdas may self-reference; guard with the calc engine's existing cycle/depth protections (coordinate with Spec 04 if it lands first — its evaluation stack is the natural place for the depth limit).

### Phase 3 — Helpers

`MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY, ISOMITTED` — regular registered functions that take a lambda argument (now representable as `AnyValue.Lambda`) and apply it element-/row-/column-wise, producing arrays that spill via the existing engine. `ISOMITTED` requires optional-parameter plumbing in lambda invocation (omitted args bind a distinct "omitted" marker).

## Work plan

| # | Task | Size |
|---|------|------|
| 1 | Parser capability spike: LET/LAMBDA/call-on-expression/`_xlpm.` through ClosedXML.Parser; write findings in PR (if parser lacks support, file upstream issue — parser is a sibling ClosedXML project — and stop) | S |
| 2 | `XLNameEnvironment` + LET special form + tests | M |
| 3 | LET round-trip (`_xlfn`/`_xlpm` save/load) + tests with a real Excel-authored file in `XLibur.Tests/Resource/` | S |
| 4 | `AnyValue.Lambda` case + exhaustive-switch audit | M |
| 5 | LAMBDA definition/invocation incl. defined-name UDFs + `#CALC!` semantics | L |
| 6 | Helper functions + spill integration + `ISOMITTED` | M |
| 7 | Test corpus: Excel-authored workbook exercising LET/LAMBDA/helpers with cached values — assert XLibur recomputes identical values | M |

## Acceptance criteria

1. `LET(x, 2, y, x*3, x+y)` = 8; sequential binding, shadowing, and name-precedence tests pass.
2. A workbook authored in Excel using LET/LAMBDA/MAP/BYROW loads, recalculates to the same values Excel cached, and saves back openable in Excel with formulas intact.
3. LAMBDA via defined name is invocable from cell formulas; bare LAMBDA in a cell yields `#CALC!`.
4. Recursion terminates with Excel-like error (not stack overflow) at depth limit.
5. No regression across the calc-engine suite; spill tests for MAP/BYROW/MAKEARRAY.

## Risks

- Parser support is the go/no-go gate — that's why task 1 is a spike with a stop condition.
- Adding a case to a hand-rolled discriminated union touches many switches; lean on the compiler (the union's Match methods) and the nullable/warnings-as-errors build to find them all.
- Volatile semantics inside lambdas and dirty-tracking of lambda-invoked references: ensure `DependenciesVisitor` treats a lambda body's references as precedents of the calling cell (add tests: edit a cell referenced only inside a lambda body → dependent recalculates).

## Results

**Task 1, parser spike: done.** Parser 5.1.0 (#863) reads every form: LET, `LAMBDA(…)(args)` and chained calls, `_xlpm.` names, and helpers with a lambda argument. A call of an expression is a `NotSupportedNode` that keeps its operands, so the cells they read are precedents.

**Task 3, `_xlpm.` on save: done** (#864). A typed formula's parameters are written as Excel 365 writes them. Excel refused to open a workbook whose parameters had no prefix.

**Task 2, LET: done.** `LET` is a special form in `CalculationVisitor`. A `LetScope` chain on `CalcContext` holds the bindings, a `NameNode` looks there before the defined names, and names match in any case, with or without `_xlpm.`. A defined name and another cell are evaluated in contexts of their own, so a LET name never reaches their formulas. The dependency walk reads LET with a scope of its names too (PR #865 review, below).

The fixture is `Resource/Other/ExcelAuthored/LetFormulas.xlsx`, typed and calculated by Excel 365 over COM. Facts from it:

- Excel refuses duplicate names, too few arguments and an even argument count as they are typed. XLibur gives `#VALUE!` for such a formula from a file.
- An error bound to a name reaches the result only where the name is used: `LET(x,1/0,5)` is 5.
- A range stays a reference: `INDEX(r,2)`, `ROWS(r)` and `ISREF(r)` all see the range.
- Excel accepts a name that reads as a cell, `LET(A1,1,A1)`, written `_xlpm.A1`. Typed, the parser reads `A1` as a reference, so XLibur does not bind it: a known gap, and a parser question.
- Excel writes each use of a name in the case it was declared in (`_xlpm.X+1`). XLibur keeps the case it was typed in, and Excel reads both.
- **A LET value is intersected in a legacy formula.** Excel gives 2 for `LET(x,A1:A3*2,SUM(x))` in C1 entered as a legacy formula, which it shows as `@A1:A3*2`, and 12 entered as a dynamic one. XLibur gives the same for each (PR #865 review).
- **A name can be an operand of the range operator:** Excel recalculates `LET(r,A1,SUM(r:A3))` when A2 changes. The parse-time precedents walk meets a name before the LET around it, so a formula with LET takes the AST path. There, `DependenciesVisitor` keeps a scope of names bound to reference areas and puts it aside while it reads a defined name's formula.
- **The parser reads `r:A3` as the column range `r:A` followed by `3`.** A LET name that can be a column cannot be the operand of `:`, which Excel allows. This is a parser question.
- **A bound computed array is memoized** (`MemoizedArray`, up to 65,536 elements), so a chain of names, each read twice by the next, computes each element once rather than 2^n times. A bigger array stays lazy, as `LazyArrayEvaluationTests` requires.
- A LET that returns an array is a dynamic-array formula in Excel. Through `FormulaA1`, XLibur saves it as a legacy formula, which Excel evaluates with implicit intersection. `SetDynamicFormulaA1` spills it.

**Tasks 4–5, LAMBDA: done.** `AnyValue` has a seventh case, a `LambdaFunction`: the parameter names, the body node, and the `LetScope` where the `LAMBDA` was written, so a function is lexically scoped. The case needed no switch audit beyond `AnyValue` itself: `TryPickScalar` reads a function as `#VALUE!`, so every operator and function that reduces a value sees that error, as Excel does (`SUM(LAMBDA(x,x))`, `LAMBDA(x,x)&"a"`). Only a call, a `LET` binding, a parameter and a defined name keep it. `LAMBDA` is a special form in `CalculationVisitor`, and a call is a `CallNode` (`LAMBDA(x,x*2)(3)`) or a function node whose name is no function, found in the `LET` scope and then among the defined names (`f(3)`, `Double(4)`). A registered function wins over a name, as #864 found Excel does on save. The arguments are evaluated where the call is and bound as a LET binds; the body is evaluated in the function's scope. A cell whose result is a function is `#CALC!`.

The fixture is `Resource/Other/ExcelAuthored/LambdaFormulas.xlsx`, typed and calculated by Excel 365 over COM, with the defined names `Double`, `Factorial`, `Loop`, `SumTo`, `AddB` and `Apply`. Facts from it, and from probes beside it:

- **IF must be lazy for recursion.** XLibur evaluated every argument of `IF` before calling it, so a function that calls itself never reached its base case. `IF` now evaluates only the branch its condition picks, read as `IF` reads it; an array condition in an array formula still evaluates both. A branch that is a function is the result as it is, as the array path would read it as `#VALUE!` (`IF(TRUE,LAMBDA(x,x*2))(3)` is 6 in a dynamic formula).
- **Depth.** Excel's limit depends on the stack a formula takes: it calculated `SumTo(2000)` and not `SumTo(2500)`, and a busier body failed by 2000. Past it Excel saved a strange 0 with `ca="1"`, while a function that never stops (`Loop`) is `#NUM!`, which `IFERROR` does not catch. XLibur's limit is a fixed 1,024 calls, and past it the cell is `#NUM!`, thrown as `LambdaDepthException` so `IFERROR` cannot catch it. A call takes about 2.4 KB of stack, so a 1 MB thread ran out near 420: a call that finds the stack low goes on on a thread with a 16 MB stack while the caller waits, so every thread reaches 1,024. WebAssembly has no threads, and there the stack is the limit.
- **Arguments.** More or fewer arguments than parameters is `#VALUE!`; an empty argument, `(1,)`, is blank; an error argument is the value of its parameter (`LAMBDA(x,5)(1/0)` is 5); a range stays a reference (`ROWS(LAMBDA(x,x)(A1:A3))` is 3).
- **`#CALC!` in a file.** Excel writes a cell whose result is a function as `#VALUE!` with a vm index whose rich value is an `_error` naming the error by its `errorType`, as it writes every error newer than Excel 2007, and refuses to open a file whose cell holds `#CALC!` or `#SPILL!` itself. XLibur wrote them, so a blocked spill made a workbook Excel could not open, a bug older than this phase. The owner chose to write the error as Excel does rather than as a plain `#VALUE!`, so its kind survives a save. Excel opens an `_error` structure whose only key is `errorType`, with no types part, and shows the error. The save (`XLWorkbook.RichValuePlan`) gives the in-cell images their rich values first, then one per kind of error, shared by its cells, as Excel shares it. An error's index is given as the writer writes the cell, since a formula's result is known only then; before any sheet is written, every cell that has a rich value's index takes its error's, so a sheet the save copies unchanged (#702, #719) never names a moved rich value: a changed index changes the sheet. Only `XLRICHVALUE` indexes are touched; a cube function's `XLMDX` index stays (`CubeTests`), and so does its record: the save gives a rich value an old rich value's record where it is, rather than removing the old ones and appending, which moved the records after them (PR #866 review). Rich data XLibur does not know is still kept as loaded; a new error's rich value is added after its own, with an `_error` structure, and Excel opened such a file and showed the errors (PR #866 review). A parameter as an operand of `:` (`LAMBDA(first,SUM(first:A3))(A1)`) makes the formula's precedents unknown, as the cells the range takes in are known only at the call (PR #866 review). The load reads an `_error` rich value, and rich data of images and errors is no longer kept verbatim. Excel 365 opened a workbook XLibur wrote with `#SPILL!`, `#CALC!` and `#FIELD!` cells, a `LAMBDA` result, a blocked spill, an in-cell image and a dynamic array with no repair, and showed each error and the image.
- **Defined names get prefixes.** A `LAMBDA` typed into a defined name is written `_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)`, as Excel writes it: `FormulaText.AddFuturePrefixes` now runs for a typed name as for a typed cell formula. Excel opens a workbook XLibur wrote with typed `LAMBDA` cells and names, and calculates what XLibur did.
- **Dependencies.** The body of a `LAMBDA` is walked with its parameters bound to no areas, so the cells it reads are precedents wherever it is called. A call of a defined name adds the precedents of the name's formula (`AddB(1)` recalculates when B1 changes). A call is of a name when the name is neither a registered function nor one of Excel's (`FormulaParser.IsCallOfName`), so `CHOOSE`, which XLibur lacks, stays `#NAME?` and keeps the parse-time walk.
- **Known gaps.** `ISERROR`, `ISERR`, `ERROR.TYPE` and `IFERROR` see a function as the `#VALUE!` it reads as; Excel sees no error (FALSE, FALSE, `#N/A`, and the function itself, `#CALC!`). `ISREF(A1)` is FALSE in a dynamic formula, LAMBDA or not, so `LAMBDA(x,ISREF(x))(A1)` is too, where Excel gives TRUE. A typed `LAMBDA` cell formula is saved as a legacy formula, which Excel shows with `@`, and calculates alike. A sheet before a call of a name, `Sheet1!Double(4)`, is dropped by the parser for a name that is no function.
