# Dynamic Arrays — Phase B: the spill engine

Status: **implemented** (PR [XLibur#169](https://github.com/XLibur/XLibur/pull/169)). This document
was the implementation plan; it has been updated to record what actually shipped, the deviations
from the plan, and the remaining limitations.

**Phase A** (PR #168) added the modern functions (`SEQUENCE`, `UNIQUE`, `SORT`, `SORTBY`, `FILTER`,
`XLOOKUP`, `XMATCH`) as `ReturnsArray` functions. **Phase B** is the actual *spilling*: a
dynamic-array formula in a single cell auto-fills its computed footprint into neighbouring empty
cells, growing/shrinking with the result, participates in dependency-driven recalculation, and
round-trips through save/load.

Only the **anchor** cell holds the `XLCellFormula`; spilled cells stay formula-less (matching
Excel), which preserves the "only the anchor has `<f>`" serialization model.

---

## 1. What already existed (reused)

- **Array evaluation.** `CalcContext.IsArrayCalculation` flips `CalculationVisitor.Visit(FunctionNode)`
  to `FunctionDefinition.CallAsArray`, returning the whole array for a `ReturnsArray | AllowRange.All`
  function. `XLCalcEngine.EvaluateArrayFormula` evaluates a formula to a full `Array` without the
  scalar collapse. **Reused as-is** by the spill path.
- **Array types.** `Array` (`Width`, `Height`, `this[y,x]`, `Broadcast`), `ConstArray`,
  `AnyValue.GetArraySize()`.
- **Multi-cell writeback loop.** The `FormulaType.Array` branches in `TryEvaluateSingleCell` and
  `ApplyFormula` were the template for the new `SpillDynamicArray` path.
- **Formula model.** `XLCellFormula` stores no cached value; cleanliness is an epoch. A dynamic array
  is `Type = Normal` **plus** `IsDynamicArray = true` (`DynamicArrayA1` factory), with a settable
  `Range` (`XLSheetRange`) used to hold the computed footprint.
- **Public entry point.** `XLCell.SetDynamicFormulaA1` → `DynamicArrayA1`.
- **Save metadata plumbing.** `XLWorkbook_Save.EnsureDynamicArrayMetadata`,
  `SaveContext.DynamicArrayMetaIndex`, and the `cm` attribute in `SheetDataWriter`.
- **Parser.** `A1#` parses to `UnaryOp.SpillRange`.

## 2. The gaps — and how each was closed

1. **`UnaryOp.SpillRange` evaluation** — implemented in `CalculationVisitor` (B3). `ImplicitIntersection`
   and `BinaryOp.Intersection` remain `NotImplementedException` (out of scope).
2. **Dynamic sizing** — `SpillDynamicArray` derives the footprint from `array.Width/Height` (B1).
3. **`IsDynamicArray` formulas spill** — a dedicated `SpillDynamicArray` path in both
   `TryEvaluateSingleCell` and `ApplyFormula` (taken before the `Normal` branch) evaluates via the
   array path and never applies the `[0,0]` scalar collapse (B1).
4. **Spill range stored / persisted** — the footprint is stored on `formula.Range`; save writes
   `<f t="array" ref="…">` + `cm`, and load reconstructs `IsDynamicArray` from the `cm`→XLDAPR
   metadata (B4).
5. **Dependency tree covers the spill region** — `DependencyTree.CreateFrom` registers the whole
   footprint as the `FormulaArea`; `UpdateSpillFootprint` re-registers on grow/shrink (B2).
6. **`#SPILL!` error + collision detection** — `XLError.SpillRange` added; collision and
   out-of-bounds detection in `SpillDynamicArray` (B1).
7. **Stale-footprint clearing** — `ClearSpillFootprint` blanks cells the new footprint no longer
   covers (B1).
8. **Structural shifts** — a dynamic array is now shifted **in place** (text + footprint) like an
   array formula, instead of being rebuilt (which reset `Range`). `Purge` + re-spill remains the
   safety net (B5).

## 3. Design decisions (as resolved)

- **Where the footprint is stored.** On `formula.Range`, keeping `Type = Normal,
  IsDynamicArray = true`. The anchor is `Range.FirstPoint`.
- **Anchor vs spilled cells.** Spilled cells are **formula-less**; only the anchor holds the
  formula. A per-workbook spill-owner map (`point → anchor formula`) lets a read of a spilled cell
  find its anchor.
- **Array semantics.** Dynamic arrays evaluate with `IsArrayCalculation = true` (via
  `EvaluateArrayFormula`); the scalar `[0,0]` collapse is bypassed.
- **`#SPILL!` representation.** `XLError.SpillRange = 7`. **Note the deviation:** `ERROR.TYPE(#SPILL!)`
  is **9**, not `value + 1`, because the modern errors are non-contiguous — so `ERROR.TYPE`
  special-cases it, and `XLCellValue`'s bounds check was widened to include it.

## 4. Implementation order (as shipped)

Each step landed as a separate commit on `feat/dynamic-array-spill-b1`.

### B1 — Spill an in-memory dynamic-array formula
`SpillDynamicArray` in `XLCalcEngine`: evaluate → size → collision-check (every non-anchor footprint
cell not owned by the previous footprint that holds a formula or non-blank value blocks the spill) →
clear the stale footprint → write the array → store `formula.Range`. Blocked footprint or
out-of-bounds → `#SPILL!` on the anchor only. Added the `XLError.SpillRange` member and its four
touch-points (enum, display/parse, `XLCellValue` bound, `ERROR.TYPE`).

### B2 — Dependency tree covers the spill region
`DependencyTree.CreateFrom` registers `IsDynamicArray` formulas with `formula.Range` as the
`FormulaArea` (1×1 anchor before the first spill). `UpdateSpillFootprint` re-registers when the
footprint changes, so dependents of any spilled cell are invalidated.

### B3 — The `A1#` spill-range operator
`CalculationVisitor.EvaluateSpillRange` resolves the operand to a spill anchor, forces the anchor to
evaluate first (so its footprint is current), and returns a `Reference` to `formula.Range`. A
non-anchor cell yields `#REF!`. Because the operator's range **includes the anchor**, `SUM(A1#)`
orders correctly even when the referencing formula precedes the anchor.

### B4 — Save/load round-trip
Save: a dynamic-array anchor is written as `<f t="array" ref="footprint">` + the `cm` XLDAPR
metadata; spilled cells round-trip as plain cached values. Load: `LoadContext.LoadDynamicArrayMetadata`
parses the cell-metadata part to collect the `cm` indexes referencing XLDAPR; an array formula whose
`cm` is in that set is reconstructed as a dynamic array (`Range = ref`), while other array formulas
stay CSE arrays. Storing the footprint on `Range` marks the cached child values as owned so the first
re-spill after load doesn't misfire a `#SPILL!` collision.

### B5 — Robustness & polish
- **Recalc ordering.** A per-workbook spill-owner map (rebuilt when the dependency tree is (re)built,
  kept in sync by `SpillDynamicArray`, cleared on `Purge`) lets `CalcContext.GetCellValue` throw
  `GettingDataException` at the owning anchor when a formula-less spilled cell of a dirty anchor is
  read — so the chain evaluates the anchor first. Gated on `HasSpillOwners`.
- **Shift preserves the footprint.** `ShiftDynamicArrayFormula` updates the text in place and
  relocates `Range`, fixing a spurious `#SPILL!` on re-spill after a row/column insert.
- **Out-of-bounds `#SPILL!`** — landed in B1.

## 5. Deviations & gotchas discovered during implementation

- **`ERROR.TYPE(#SPILL!)` is 9, not 8.** The naive `(int)error + 1` breaks for the new member;
  `ERROR.TYPE` special-cases it and the `XLCellValue` error-bounds check was widened.
- **The `#SPILL!` *literal* can't be parsed.** The external `ClosedXML.Parser` doesn't tokenize
  `#SPILL!` (it errors with *Unexpected token SPILL*). It can only appear as a computed value, so
  `ERROR.TYPE(#SPILL!)` is tested against a real spilled cell rather than a literal.
- **`SEQUENCE(cell-reference)` returns `#VALUE!`** in array mode (literal arguments work). This is an
  orthogonal argument-coercion issue, not a spill bug; tests use `UNIQUE(range)` where a
  reference-driven size is needed.
- **The old "split-into-`@` on shift" behaviour was already gone** — the shift path kept the dynamic
  flag but reset `Range`, which is what B5 fixed.

## 6. Known limitations (documented, with tests)

- **First-evaluation ordering.** When a dependent is positioned *before* a not-yet-spilled anchor and
  reads a spilled non-anchor cell, the footprint is unknown until the anchor runs — a genuine circular
  dependency that would need a calc-chain pre-pass to size arrays before evaluation. Captured by the
  `[Ignore]`d `Spill_DependentBeforeAnchor_FirstEvaluationOrdering` test. Post-first-spill ordering
  works; the `A1#` operator is unaffected and is the recommended way to reference a spill.
- **`#SPILL!` auto-recovery.** Clearing a cell that blocks a spill does not by itself re-spill the
  anchor (the desired footprint of a `#SPILL!` anchor isn't tracked); recovery happens on the next
  anchor re-evaluation. Editing a live spilled cell is transient (overwritten on next recalc),
  consistent with Excel treating spill cells as read-only.

## 7. Files touched

- `XLibur/Excel/CalcEngine/XLCalcEngine.cs` — spill path, footprint sizing/collision/clearing,
  spill-owner map.
- `XLibur/Excel/CalcEngine/CalculationVisitor.cs` — `UnaryOp.SpillRange`.
- `XLibur/Excel/CalcEngine/DependenciesVisitor.cs` — `SpillRange` precedent propagation (comment).
- `XLibur/Excel/CalcEngine/DependencyTree.cs` — footprint `FormulaArea` + `UpdateSpillFootprint`.
- `XLibur/Excel/CalcEngine/CalcContext.cs` — spill-owner lookup on cell reads.
- `XLibur/Excel/CalcEngine/XLError.cs` (+ `XLErrorExtensions`, `Information.ErrorType`, `XLCellValue`)
  — `#SPILL!`.
- `XLibur/Excel/Cells/XLCell.cs` — in-place dynamic-array shift.
- `XLibur/Excel/IO/SheetDataWriter.cs` — write the spill `ref`.
- `XLibur/Excel/IO/WorksheetSheetDataReader.cs` + `LoadContext.cs` + `XLWorkbook_Load.cs` —
  reconstruct `IsDynamicArray` from `cm`/XLDAPR.
- Tests: `SpillEvaluationTests`, plus `ArrayFormulaTests`, `FormulaParserTests`, `InformationTests`.

## 8. Verification

```
dotnet build XLibur/XLibur.csproj -c Release --no-restore -v q
dotnet test XLibur.Tests/XLibur.Tests.csproj -c Release -f net10.0 --filter "FullyQualifiedName~Spill|FullyQualifiedName~ArrayFormula|FullyQualifiedName~DynamicArray"
dotnet test XLibur.Tests/XLibur.Tests.csproj -c Release -f net10.0   # full suite
```

Result after Phase B: full suite **6275 passed / 0 failed** on net8.0 and net10.0.
