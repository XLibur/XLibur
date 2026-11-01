# Spec 08 — LET, LAMBDA, and Lambda-Helper Functions

**Area:** Feature (calc engine core)
**Effort:** L (2–4 weeks; parser/scoping work, not bulk registration)
**Dependencies:** Independent of Spec 07, but land LET before LAMBDA. Interacts with the spill engine (`docs/dynamic-array-spill-phase-b.md`).
**Status:** In progress — tasks 1–3 done (parser spike, LET, `_xlpm.` on save); LAMBDA (tasks 4–5) next. See [Results](#results).

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

**Task 2, LET: done.** `LET` is a special form in `CalculationVisitor`. A `LetScope` chain on `CalcContext` holds the bindings, a `NameNode` looks there before the defined names, and names match in any case, with or without `_xlpm.`. A defined name and another cell are evaluated in contexts of their own, so a LET name never reaches their formulas. Dependencies needed no change, because LET reaches the dependency walk as an unknown function whose reference arguments are precedents.

The fixture is `Resource/Other/ExcelAuthored/LetFormulas.xlsx`, typed and calculated by Excel 365 over COM. Facts from it:

- Excel refuses duplicate names, too few arguments and an even argument count as they are typed. XLibur gives `#VALUE!` for such a formula from a file.
- An error bound to a name reaches the result only where the name is used: `LET(x,1/0,5)` is 5.
- A range stays a reference: `INDEX(r,2)`, `ROWS(r)` and `ISREF(r)` all see the range.
- Excel accepts a name that reads as a cell, `LET(A1,1,A1)`, written `_xlpm.A1`. Typed, the parser reads `A1` as a reference, so XLibur does not bind it: a known gap, and a parser question.
- Excel writes each use of a name in the case it was declared in (`_xlpm.X+1`). XLibur keeps the case it was typed in, and Excel reads both.
- A LET that returns an array is a dynamic-array formula in Excel. Through `FormulaA1`, XLibur saves it as a legacy formula, which Excel evaluates with implicit intersection. `SetDynamicFormulaA1` spills it.
