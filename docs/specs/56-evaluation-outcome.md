# Spec 56 — An evaluation failure gets one outcome, and each entry point reads one policy table

**Area:** Architecture · **Defects (5 executed, 1 refuted)** · **API (additive)** · **Behaviour change (`fix!:`)**
**Effort:** M (~4 days)
**Dependencies:** None hard. **Must land before spec 32**, which rewrites `SignatureAdapter.cs`. Soft
file overlaps with specs 42, 43, 04, 30 and 54 — see *Conflicts*.
**Status:** Proposed. From the 2026-09-13 architecture review (round 4). Every design decision below
was taken by the owner in a design interview (see *Decisions*). Also recorded:
`docs/adr/0001-save-policy-on-evaluation-failure.md`, and the terms *error value*, *unsupported
feature*, *defect*, *circular reference*, *cached value* and *dirty* in `CONTEXT.md`.

Line numbers are from `d68dffdc` (2026-09-13). Verify them before editing.

---

## Goal

The calc engine sorts every evaluation failure once, into one internal outcome. Each public entry
point decides what its caller sees by reading one policy table. Public exception types are raised in
one place.

## Why this spec exists

### Seven places answer one question

"What does this failure mean to the caller?" is answered at each of these, read at `d68dffdc`:

| Site | What it does |
|---|---|
| `Cells/XLCell.cs:315`, `:343` | `catch when EvaluationFailure.IsExpected` — false, or the cached value |
| `Ranges/XLRangeBase.cs:588` | same, for `Search` |
| `CalcEngine/XLCalcEngine.cs:305`, `:376` | `catch (GettingDataException)` — evaluate the precedent, retry |
| `CalcEngine/XLCalcEngine.cs:495` | `catch (MissingContextException)` → public `XLNoWorksheetContextException` |
| `CalcEngine/XLFunctionLibrary.cs:107` | the same translation, written a second time |
| `IO/SheetDataWriter.cs:394` | **a bare `catch`** — every failure, defects included |
| `XLibur.Report/Functions/ExcelFunctionBridge.cs:144`, `Ranges/CellEvaluator.cs:60`, `:74` | Report translates a third time |

The classifier, `EvaluationFailure.IsExpected` (`Exceptions/EvaluationFailure.cs:17-25`), works by
exception type: `NotImplementedException`, `NotSupportedException`, `ExpressionParseException`, the two
missing-context types and `CircularReferenceException`. So `NotImplementedException` means two things:

- **"XLibur does not evaluate this"** — `CalculationVisitor.cs:123` and `XLCalcEngine.cs:428`;
- **"an argument converter has no branch for this"** — `SignatureAdapter.cs:1249`, `:1258` and `:1263`,
  all "Array formulas not implemented."

There are no other `NotImplementedException` throws in the calc engine, so every one of them means
*unsupported*. The type cannot tell them apart from a `NotImplementedException` thrown anywhere else in
the library, which is a defect.

### The five defects, executed

All executed against a scratch build at `d68dffdc` before this spec was written.

- **D57 — `wb.Evaluate` on a dirty cell throws an internal exception.** `Sheet1!A1.FormulaA1 = "1+1"`,
  not read. `wb.Evaluate("Sheet1!A1")` throws `GettingDataException` — `internal`, not an
  `InvalidOperationException`, with the default message "Exception of type … was thrown." So do
  `wb.Evaluate("Sheet1!A1+0")` and a cell whose input was edited after it was read.
  `ws.Evaluate("A1")` on the same dirty cell returns 2. The cause is `recursive: false` at
  `XLWorkbook.cs:1192`, which leaves the precedent's `GettingDataException` (`CalcContext.cs:189`) with
  no one to catch it. **This is D37's leak** — an internal type escaping a public `Evaluate` — which
  D37 closed for `MissingContextException` only.
- **D58 — a circular reference escapes as an internal type.** `A1.FormulaA1 = "A1+1"`; `A1.Value`
  throws `CircularReferenceException`, which is an `InvalidOperationException` and `IsPublic == false`.
  A caller can catch it only as `InvalidOperationException`, which also covers misuse and defects.
- **D59 — `=LEN({"ab","c"})` throws `NotImplementedException`.** Excel returns 2. The fix is spec 30's
  (array application). This spec only classifies it: an **unsupported feature**, not a defect.
- **D60 — a defined name that calls `ROW()` fails inside a cell.** Name `MyRow = ROW()`; `A6 = MyRow`;
  `A6.Value` throws `XLNoWorksheetContextException`, whose message advises "Use it in a cell formula" —
  it already is one. Excel gives 6. `EvaluateName` (`XLCalcEngine.cs:~723`) builds a context without
  the calling cell's address. The failure is "no context", raised where a context exists.
- **D61 — save swallows every evaluation failure.** A workbook with `A1 = A1+1` (a cycle) and
  `B1 = LEN({"ab","c"})` saves without a word. Both cells are written as `<c><f>…</f></c>` with no
  `<v>`. A defect in any function would be written exactly the same way, and #459's rule — a defect
  reaches the caller — does not hold on save.

### One refuted

**P12 — "a throw inside `Recalculate` leaves the engine broken".** The missing `finally` around
`_chain.Reset()` (`XLCalcEngine.cs:333-341`) looked dangerous. Executed: the first
`RecalculateAllFormulas` throws on the cycle; after the cycle is fixed, the second succeeds and a
dependent reads 20, as it should. **Not a defect; not in scope.**

### Read, not executed

With `LoadOptions.RecalculateAllFormulas = true`, the `XLWorkbook` constructor calls
`RecalculateAllFormulas` (`XLWorkbook.cs:944`, `:993`), which throws on the first cycle
(`XLCalcEngine.cs:361`). So a file Excel opens cannot be opened with that option set. Task 1 executes
this.

## Decisions (owner, 2026-09-13)

| # | Decision |
|---|---|
| Q13 | On save, an expected failure writes no cached value, and Excel recalculates on open; a defect throws. Recorded as ADR 0001. |
| Q14 | A circular reference reaches a caller as a **public** type that still derives from `InvalidOperationException`. |
| Q15 | An internal type marks an unsupported feature. Any `NotImplementedException` left elsewhere is a defect. |
| Q16 | The outcome type stays internal. `XLibur.Report` maps the public exceptions itself. |
| Q17 | `wb.Evaluate` calculates a dirty formula first, as `cell.Value` does. |
| Q21 | The argument converter's array branch raises the unsupported-feature type: passing an array to a scalar function is a missing feature, and spec 30 is its fix. |
| Q22 | Expected failures are a circular reference, an unsupported feature and a refused formula. Everything else is a defect. |
| Q23 | `RecalculateAllFormulas`, including recalculate-on-load, skips the cells in a cycle and carries on. The cycle surfaces when one of those cells is read. |
| Q24 | The public type is `XLCircularReferenceException`, in `XLibur.Excel.CalcEngine.Exceptions`, next to `XLNoWorksheetContextException`. It derives from `InvalidOperationException`, not `XLiburException`. |
| Q25 | An unsupported feature reaches a caller as `NotImplementedException`, as it does today. No API change. |
| Q26 | The `ROW()`-in-a-name fix is part of this spec, as a defects-first task. |
| Q38 | Where these decisions say nothing, the policy table records today's behaviour and changes nothing. |
| Q39 | Report maps `XLCircularReferenceException` to a template error; the gate runs all four test projects. |

## Non-goals

- **Fixing the converter's array branch** so `LEN({"ab","c"})` gives 2. Spec 30.
- **Any cell of the policy table not decided above.** Q38: recorded, not changed. The table is where a
  later change goes.
- **The missing `finally` around `_chain.Reset()`.** P12 refuted the defect.
- **Error-value representation** — the `XLError` ↔ text tables. Round-4 candidate 05, not chosen.
- **The load path's error mode** (`PartStructureException`). A separate backlog item.
- **Demand-driven evaluation.** Spec 04.

## File structure

| File | Change |
|---|---|
| `XLibur/Excel/CalcEngine/Exceptions/EvaluationFailure.cs` | Classifies into a kind rather than answering "expected?" |
| `XLibur/Excel/CalcEngine/Exceptions/UnsupportedFeatureException.cs` | **New.** `internal sealed`, derives from `NotImplementedException` |
| `XLibur/Excel/CalcEngine/Exceptions/CircularReferenceException.cs` | **Replaced** by the public `XLCircularReferenceException.cs` |
| `XLibur/Excel/CalcEngine/EvaluationPolicy.cs` | **New.** The policy table |
| `XLibur/Excel/CalcEngine/CalculationVisitor.cs`, `XLCalcEngine.cs`, `Functions/SignatureAdapter.cs` | The five throw sites raise the new type; `EvaluateName` passes the address; recalculation skips cycles; one missing-context translation |
| `XLibur/Excel/CalcEngine/XLFunctionLibrary.cs` | Reads the table |
| `XLibur/Excel/Cells/XLCell.cs`, `Ranges/XLRangeBase.cs` | Read the table |
| `XLibur/Excel/XLWorkbook.cs` | `Evaluate` evaluates recursively |
| `XLibur/Excel/IO/SheetDataWriter.cs` | `EvaluateFormulaForSave` reads the table |
| `XLibur/PublicAPI.Unshipped.txt` | Gains `XLCircularReferenceException` |
| `XLibur.Tests/PublicSurfaceTests.cs` | `CircularReferenceException` leaves any `MustStayInternal` list; the public type is expected |
| `XLibur.Report/Ranges/CellEvaluator.cs` | Maps the public cycle type to a template error |
| `XLibur.Tests/Excel/CalcEngine/EvaluationOutcomeTests.cs` | **New.** The matrix |

## The design

### 1. Kinds

| Kind | Raised as, after this spec | Classified from |
|---|---|---|
| **Cycle** | `XLCircularReferenceException` (public) | that type |
| **Unsupported** | `UnsupportedFeatureException` (internal, derives from `NotImplementedException`) | that type, and `NotSupportedException` (recorded as today) |
| **Refused** | `ExpressionParseException` | that type |
| **No context** | `MissingContextException` inside; `XLNoWorksheetContextException` at the public edge | either |
| **Pending** | `GettingDataException` (internal) | that type — must never reach a public caller |
| **Defect** | anything else, including a plain `NotImplementedException` | everything else |

Because `UnsupportedFeatureException` derives from `NotImplementedException`, a caller's existing
`catch (NotImplementedException)` keeps working and the public edge needs no translation (Q25). The
classifier can still tell the two apart.

### 2. The policy table

One internal module, `EvaluationPolicy`, maps entry point × kind to what the caller sees. **Bold**
cells are decided by this spec; the rest record today's behaviour, which task 1 measures and writes in.

| Entry point | Cycle | Unsupported | Refused | No context | Defect |
|---|---|---|---|---|---|
| `IXLCell.Value` | **throw `XLCircularReferenceException`** | throw `NotImplementedException` | throw `ExpressionParseException` | today *(D60 removes the case where a context exists)* | throw |
| `TryGetValue`, `GetFormattedString`, `Search` | today | today | today | today | throw (#459) |
| `Evaluate`, `EvaluateExpr` | **throw `XLCircularReferenceException`** | today | today | throw `XLNoWorksheetContextException` (D37) | throw |
| `XLFunctionLibrary.TryInvoke` | today | today | today | throw `XLNoWorksheetContextException` | throw |
| `RecalculateAllFormulas`, recalculate-on-load | **skip the cycle's cells, carry on** | today | today | today | throw |
| Save | **no cached value** | **no cached value** | **no cached value** | **throw** | **throw** |
| Report `CellEvaluator` | **template error** | today | today | template error (today) | throw |

The translation from `MissingContextException` to the public type is written once, in this module.

### 3. `Evaluate` calculates a dirty formula first

`XLWorkbook.Evaluate` (`XLWorkbook.cs:1192`) stops passing `recursive: false`, so a dirty precedent is
evaluated first, as `cell.Value` and `ws.Evaluate` already do. D57 turns green. "Pending" then cannot
reach any public entry point, and a test asserts that for each one.

### 4. A defined name gets its cell

`EvaluateName` passes the calling formula's address into the context it builds. D60 turns green.

**Interaction with spec 53.** Spec 53 turns on implicit intersection wherever a legacy formula has a
cell to intersect against. Once a name has its cell, operators at the top level of a name's formula
intersect their range operands too. That is what Excel does, and the test for this task pins one case.
Spec 53's trick for asking Excel — set the formula through `Range.Formula` and read where Excel adds
`@` — can confirm it if the answer is in doubt.

### 5. Save (ADR 0001)

`SheetDataWriter.EvaluateFormulaForSave` loses its bare `catch` and reads the table. A cycle, an
unsupported feature or a refused formula leaves the cell with no `<v>`. A defect throws out of
`SaveAs`.

### 6. Recalculation skips cycles

When `RecalculateCurrentCell` meets a cycle, it leaves the cycle's cells dirty with no value and moves
on. Reading one of those cells later throws `XLCircularReferenceException`. A workbook containing a
cycle opens with `LoadOptions.RecalculateAllFormulas = true`.

## Global constraints

- **Additive public API only.** `XLCircularReferenceException` and its constructor are added to
  `PublicAPI.Unshipped.txt`. It derives from `InvalidOperationException`, so no existing `catch` stops
  working — **check the base type before writing "still works"**, the lesson D37 recorded.
- **Two behaviour changes a caller can see** — save can now throw, and recalculation no longer throws
  on a cycle. Both are `fix!:`.
- **Q38 holds.** No cell of the table outside the decisions changes. If the matrix test shows a
  "today" cell that looks wrong, record it in *Results*; do not fix it here.
- **Run all four test projects.**
- The ground rules in `TASKLIST-architecture-deepening-4.md` §6 apply.

## Work plan

One owner, one branch: `fix/56-evaluation-outcome`.

### Task 1 — The matrix, and the defects land red

`EvaluationOutcomeTests`: every entry point × every kind, asserting **today's** behaviour. To make a
defect happen on purpose, use the prior art in `EvaluationFailureTests`, which has done it since
#459. Separate tests assert the **correct** behaviour for D57, D58, D60 and D61; they fail, and the
commit names them. D59 gets a test asserting that it classifies as unsupported. Execute the
recalculate-on-load lead and record the result.

### Task 2 — Kinds

`UnsupportedFeatureException` at the five throw sites. `XLCircularReferenceException` replaces the
internal type. The classifier returns a kind. `PublicAPI.Unshipped.txt` and `PublicSurfaceTests` are
updated. **D58 turns green.**

### Task 3 — The policy table

Route `XLCell`, `XLRangeBase`, `XLCalcEngine` (`:495`) and `XLFunctionLibrary` through the table. One
missing-context translation remains.

### Task 4 — Save

ADR 0001. **D61 turns green.** Assert that a defect thrown from a test function makes `SaveAs` throw,
and that the three expected kinds write no `<v>`.

### Task 5 — Recalculation, `Evaluate`, names

Recalculation skips cycles, and recalculate-on-load opens a workbook containing one. `Evaluate` is
recursive (**D57**). Names get their cell (**D60**), with spec 53's interaction pinned.

### Task 6 — Report

`CellEvaluator` maps the public cycle type to a template error. Run all four test projects on both
TFMs.

### Task 7 — Changelog

`feat:` — `XLCircularReferenceException` is public. `fix!:` — save throws on a defect; recalculation
skips cycles instead of throwing. `fix:` — D57 and D60.

## Acceptance criteria

| # | Criterion |
|---|---|
| 1 | `grep -rn "catch (MissingContextException" XLibur/` matches one line |
| 2 | `SheetDataWriter.cs` has no untyped `catch` |
| 3 | `grep -rn "new NotImplementedException" XLibur/Excel/CalcEngine/` matches nothing |
| 4 | `CircularReferenceException` is gone; `XLCircularReferenceException` is in `PublicAPI.Unshipped.txt` |
| 5 | D57, D58, D60 and D61 are green; D59 classifies as unsupported |
| 6 | A workbook with a cycle opens with `LoadOptions.RecalculateAllFormulas = true` |
| 7 | The matrix test is green, and every cell that changed is one this spec decided |
| 8 | All four test projects are green on net8.0 and net10.0 |

## Conflicts

| Spec | Shared ground | Resolution |
|---|---|---|
| **32** | `SignatureAdapter.cs` — 56 changes three throw sites, 32 rewrites the file | **56 first** |
| **30** | `CalculationVisitor.cs` — 30 edits the function-call dispatch (`:87-89`), 56 the throw at `:123`. 30 also replaces the converter branch that 56 reclassifies | Soft. Either order; if 30 lands first, D59 stops reproducing and its test becomes a regression test |
| **42, 43** | `XLCalcEngine.cs` — 43 owns `TryEvaluateSingleCell` and `ApplyFormula`; 56 owns `EvaluateName`, `:495` and the recalculation loop | Soft, adjacent regions. Whichever lands second rebases |
| **04** | the evaluation stack | Soft. Assess at dispatch |
| **54** | `XLCell.cs` — different regions | Soft |
| **45** | `SheetDataWriter.cs` — 45 edits the cell value writers, 56 `EvaluateFormulaForSave` | Soft, different methods |
| **55** | `XLWorkbook.cs` — 55 removes `NotifyWorksheetDeleting` (`:1370`), 56 edits `Evaluate` (`:1192`) and the constructor path (`:944`) | Soft |
