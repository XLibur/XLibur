# worksheet-cell-focus Specification

## Purpose

Provide explicit, opt-in entry points for setting a worksheet's active cell and scrolling it into view, while keeping low-level control over the scrollable pane's top-left anchor. Focus operations adjust scroll position; existing active-cell APIs do not.

## Requirements

### Requirement: Pane top-left scroll primitive

`IXLSheetView` SHALL expose a nullable `PaneTopLeftCellAddress` that controls the `<pane topLeftCell>` attribute of the scrollable region on a frozen sheet.

- When `PaneTopLeftCellAddress` is `null`, XLibur SHALL anchor the pane to the first non-frozen cell (`split+1`), preserving existing behaviour.
- When set, XLibur SHALL emit that address as `<pane topLeftCell>` on save.
- Setting an address whose worksheet differs from the sheet view's worksheet SHALL throw `ArgumentException`, consistent with `TopLeftCellAddress`.
- When the sheet has no frozen pane (no `<pane>` element is emitted), `PaneTopLeftCellAddress` SHALL have no effect on output.

#### Scenario: Low-level pane control on a frozen sheet
- **WHEN** a worksheet has the top 2 rows frozen and the consumer sets `SheetView.PaneTopLeftCellAddress = Cell("A3").Address` and saves
- **THEN** the saved `<pane>` element has `topLeftCell="A3"`

#### Scenario: Null pane address preserves normalize-to-top default
- **WHEN** a worksheet has the top 2 rows frozen and `SheetView.PaneTopLeftCellAddress` is never set, and the worksheet is saved
- **THEN** the saved `<pane>` element has `topLeftCell="A3"` (the `split+1` anchor)

#### Scenario: Pane address ignored without a pane
- **WHEN** a worksheet has no frozen panes and `SheetView.PaneTopLeftCellAddress = Cell("M50").Address` is set, and the worksheet is saved
- **THEN** no `<pane>` element is emitted and `topLeftCell="M50"` does NOT appear in the sheet view

### Requirement: Scroll a frozen-sheet active cell into view

`IXLWorksheet.FocusCell(address)` SHALL set the active cell and selection, and on a frozen sheet whose target lies below/right of the split, SHALL set `<pane topLeftCell>` so the target is at the top-left of the scrollable region and reset the orthogonal scroll axis so the target is not parked off-screen sideways.

#### Scenario: Active cell on a frozen sheet is scrolled into view
- **GIVEN** a worksheet with the top 2 rows frozen
- **AND** the sheet was previously loaded with `<sheetView topLeftCell="G1">` (a residual horizontal scroll)
- **WHEN** the consumer calls `worksheet.FocusCell("A3")` and saves
- **THEN** the `<selection>` has `activeCell="A3"` and `sqref="A3"`
- **AND** the `<pane>` element has `topLeftCell="A3"`
- **AND** `<sheetView topLeftCell>` is absent (reset to the `A1` origin, clearing the residual `G1`)

### Requirement: Scroll a non-frozen active cell into view

On a worksheet with no frozen panes, `FocusCell` SHALL set `<sheetView topLeftCell>` to the target address (top-left anchor) and SHALL NOT emit a `<pane>` element.

#### Scenario: Active cell on a non-frozen sheet is scrolled into view
- **GIVEN** a worksheet with no frozen panes
- **WHEN** the consumer calls `worksheet.FocusCell("M50")` and saves
- **THEN** `<sheetView topLeftCell>` is `"M50"`
- **AND** the `<selection>` has `activeCell="M50"` and `sqref="M50"`
- **AND** no `<pane>` element is emitted

### Requirement: Focus is opt-in and backward compatible

Scroll position SHALL change only through the explicit focus entry points (`FocusCell` / `IXLCell.Focus`). Existing active-cell APIs and `SetActiveCell` SHALL not move scroll.

#### Scenario: Default preserves current behaviour
- **WHEN** the consumer calls `cell.SetActive()`
- **THEN** only the selection/active cell changes and no scroll attribute is altered

#### Scenario: SetActiveCell does not move scroll
- **WHEN** the consumer calls `worksheet.SetActiveCell("A3")`
- **THEN** the active cell and `sqref` are set to `A3` and neither `<pane topLeftCell>` nor `<sheetView topLeftCell>` is changed by the call

### Requirement: Focus target inside the frozen region

When `FocusCell` targets a cell that lies within the frozen rows/columns (already always visible), it SHALL reset the scrollable pane to its origin (`split+1`) and SHALL emit a `<selection>` whose `pane` attribute matches the frozen pane that actually owns the active cell, so the produced `<selection>`/`<pane>` pairing is internally consistent.

#### Scenario: Focusing a cell in the frozen top rows
- **GIVEN** a worksheet with the top 2 rows frozen
- **WHEN** the consumer calls `worksheet.FocusCell("A1")` and saves
- **THEN** the `<selection>` `activeCell` is `"A1"`
- **AND** the `<pane>` element has `topLeftCell="A3"` (scrollable region reset to origin)
- **AND** the `<selection>` `pane` attribute names the pane that owns `A1`, not a pane the cell is absent from
