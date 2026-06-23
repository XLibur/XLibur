## Why

Consumers can set a worksheet's active cell, but there is no supported way to ensure that cell is actually *visible* when the file opens. A workbook can be saved with the cursor at `A3` while the scroll is parked far away (e.g. `<pane topLeftCell="A966">`), so the user opens the file staring at empty rows. XLibur today exposes only `SheetView.TopLeftCellAddress`, which maps to `<sheetView topLeftCell>` — the *frozen-header* scroll. It offers **no** control over `<pane topLeftCell>`, the scroll of the scrollable region, which is the attribute that actually parks the active cell off-screen on a frozen sheet. There is also no one-call "put the cursor here and make sure it's on screen" convenience.

## What Changes

- **Add a low-level primitive** `IXLSheetView.PaneTopLeftCellAddress` (nullable) mapping to `<pane topLeftCell>`. `null` preserves today's behaviour (XLibur anchors the pane to the first non-frozen cell, `split+1`); a set value is honored on write.
- **Add a high-level convenience** on `IXLWorksheet`: `FocusCell(address)` sets the active cell + selection (`sqref`) **and** scrolls the sheet so the cell is visible at the top-left of the scrollable region — choosing `<pane topLeftCell>` vs `<sheetView topLeftCell>` based on whether the sheet is frozen, and resetting the orthogonal axis so the cell is not left scrolled away sideways. The companion `SetActiveCell(address)` sets the active cell + selection **without** moving scroll.
- **Add a fluent cell entry point** `IXLCell.Focus()` (focus + scroll into view), leaving the existing `SetActive(bool value = true)` signature untouched.
- **Intent-named methods, not a boolean flag** — `FocusCell` always scrolls, `SetActiveCell` never does; no `scrollIntoView` control flag.
- **"Make visible" is defined as top-left anchoring of the scrollable region** — deterministic, no viewport/centring math (XLibur has no rendered viewport).
- Backward compatible: existing `cell.SetActive()` / `ActiveCell` set behaviour is unchanged; scroll only moves through the explicit `FocusCell`/`Focus` entry points.

## Capabilities

### New Capabilities

- `worksheet-cell-focus`: Setting the active cell and (optionally) scrolling the sheet so that cell is visible, including the low-level `<pane topLeftCell>` primitive and the high-level scroll-into-view convenience, across frozen and unfrozen sheets.

### Modified Capabilities

<!-- None. openspec/specs/ is empty; all behaviour here is new public surface. -->

## Impact

- **Public API (additive):** `IXLSheetView.PaneTopLeftCellAddress`, `IXLWorksheet.FocusCell(...)` + `IXLWorksheet.SetActiveCell(...)`, `IXLCell.Focus()`. New entries in `PublicAPI.Unshipped.txt`.
- **IO write path:** `SheetViewWriter.SetupPane` must honor a set `PaneTopLeftCellAddress` instead of always recomputing `split+1`; `WriteSheetViews` orthogonal-axis reset.
- **Model:** `XLSheetView` gains the pane top-left field and copy-constructor handling.
- **Reader:** intentionally **unchanged** — `<pane topLeftCell>` stays unread so bare round-trips keep self-healing to top (see design.md for the normalize-vs-faithful decision).
- **Tests/Examples:** new tests under `XLibur.Tests/Excel/Worksheets`; optional example in `XLibur.Examples/Misc`.
