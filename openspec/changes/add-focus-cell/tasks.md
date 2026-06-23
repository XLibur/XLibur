## 1. Low-level primitive: PaneTopLeftCellAddress

- [x] 1.1 Add `IXLAddress? PaneTopLeftCellAddress { get; set; }` to `IXLSheetView` with XML docs (maps to `<pane topLeftCell>`; null = anchor to split+1)
- [x] 1.2 Implement the property on `XLSheetView` with same-worksheet validation as `TopLeftCellAddress` (throw `ArgumentException` on mismatch)
- [x] 1.3 Handle the new field in the `XLSheetView` copy-constructor
- [x] 1.4 Update `SheetViewWriter.SetupPane` to honor a set `PaneTopLeftCellAddress`, falling back to the `split+1` recompute when null
- [x] 1.5 Confirm the reader (`WorksheetElementReader.LoadSheetViewPane`) is intentionally left unchanged (does NOT populate the property); add a code comment noting the normalize-to-top default decision

## 2. High-level convenience: SetActiveCell + FocusCell

- [x] 2.1 Add `IXLWorksheet.SetActiveCell(string address)` / `SetActiveCell(IXLCell cell)` (active cell + `sqref`, no scroll) and `IXLWorksheet.FocusCell(string address)` / `FocusCell(IXLCell cell)` (active + selection + scroll) to the interface
- [x] 2.2 Implement `SetActiveCell` on `XLWorksheet`: set active cell + single-cell selection (`sqref`) only; implement `FocusCell` as `SetActiveCell` plus scroll anchoring
- [x] 2.3 Implement scroll logic in `FocusCell`: frozen target below/right of split → set `PaneTopLeftCellAddress` to target; unfrozen → set `SheetView.TopLeftCellAddress` to target
- [x] 2.4 Implement orthogonal-axis reset (clear residual scroll on the non-target axis to origin)
- [x] 2.5 Implement frozen-region handling: reset pane to `split+1` and emit `<selection pane=>` naming the pane that owns the active cell (Decision 4)
- [x] 2.6 Add fluent `IXLCell Focus()` (focus + scroll into view), leaving the existing `SetActive(bool value = true)` untouched

## 3. Public API surface

- [x] 3.1 Add new members to `PublicAPI.Unshipped.txt`
- [x] 3.2 Verify build with warnings-as-errors (nullable annotations correct on the new nullable property)

## 4. Tests

- [x] 4.1 Pane primitive: set `PaneTopLeftCellAddress` on a frozen sheet → `<pane topLeftCell>` equals the value
- [x] 4.2 Null default: unset property on a frozen sheet → `<pane topLeftCell>` equals `split+1`
- [x] 4.3 No-pane: set property on an unfrozen sheet → no `<pane>`, value not emitted
- [x] 4.4 Frozen scroll-into-view: `FocusCell("A3")` with a residual `<sheetView topLeftCell="G1">` → `selection activeCell/sqref="A3"`, `<pane topLeftCell>="A3"`, `<sheetView topLeftCell>` absent
- [x] 4.5 Non-frozen scroll-into-view: `FocusCell("M50")` → `<sheetView topLeftCell>="M50"`, no `<pane>`
- [x] 4.6 Backward compat: `cell.SetActive()` and `SetActiveCell("A3")` change no scroll attributes
- [x] 4.7 Frozen-region target: `FocusCell("A1")` on a row-frozen sheet → pane reset to `split+1`, `<selection pane>` names the owning pane
- [x] 4.8 Freeze-shape matrix: row-only, column-only, and both-axis freezes covered (split windows are not XLibur-producible — see Decision 5)
- [x] 4.9 Same-worksheet validation: setting `PaneTopLeftCellAddress` from another worksheet throws `ArgumentException`

## 5. Docs / examples

- [x] 5.1 Add or extend an example in `XLibur.Examples/Misc` demonstrating `worksheet.FocusCell(...)`
- [x] 5.2 Note the normalize-to-top vs. honored-when-set behaviour in the `PaneTopLeftCellAddress` XML docs
