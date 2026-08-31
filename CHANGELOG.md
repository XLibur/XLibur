# Changelog

## Contents

- [Unreleased](#unreleased)
- [v0.311.1](#v03111---2026-08-11)
- [v0.311.0](#v03110---2026-08-08)
- [v0.310.0](#v03100---2026-08-06)
- [v0.301.0](#v03010---2026-08-04)
- [v0.300.0](#v03000---2026-08-02)
- [v0.200.0](#v02000---2026-08-01)
- [XLibur.Report — Unreleased](#xliburreport--unreleased)
- [v0.106.0](#v01060---2026-07-25)

## Unreleased

### ⚠️ Breaking Changes

#### Sheet views

- **`SheetView.SplitRow` and `SheetView.SplitColumn` assigned without a `Freeze` call now produce a split bar rather than a freeze, and their values are read in different units.** These have always been public, and assigning them used to write `<pane state="frozen">`, because the pane writer never asked whether a freeze was meant. It now asks, so the same code produces `<pane state="split">` — and OOXML measures an unfrozen split's position in twentieths of a point, not in lines. `SheetView.SplitRow = 3` has therefore gone from freezing three rows to placing a split bar 3/20 of a point from the top, which is visually nothing. **There is no compile-time signal for this**: the code still compiles and silently does something else. To keep the old behaviour, call `Freeze(3, 0)` — or `FreezeRows(3)` — instead of assigning the property. To get a real split bar, scale the value: 900 is three default 15pt rows, 2880 three default 48pt columns. XLibur carries whichever number you give it verbatim rather than converting between the two units, so a file round-trips exactly; converting would need every column width and row height on the sheet and is not reversible.

- **`IXLSheetView` gains a `FreezePanes` property**, which is source-breaking for anyone implementing that interface outside the library. It is the value that distinguishes a freeze from a split; the declaration already existed in the file, commented out.

#### Pivot tables

- **`IXLPivotTable` gains a `ShowLastColumn` property and its two `SetShowLastColumn` overloads**, which is source- and binary-breaking for anyone implementing that interface outside the library. The setting already existed and already round-tripped; it was simply the one of the five table-style emphasis flags (`ShowRowHeaders`, `ShowColumnHeaders`, `ShowRowStripes`, `ShowColumnStripes`) that had never been put on the interface, so it could not be reached through `IXLPivotTable`.

### ✨ New Features

- **Slicers loaded from a workbook can now be read.** `IXLWorksheet.Slicers` lists the slicers drawn on a sheet, and `IXLPivotTable.Slicers` shows which of them filter a given pivot table. A slicer reports its name, caption, style, column count, row height, the field it filters and the items currently selected — for both kinds Excel writes: pivot slicers, which read their selection from the slicer cache, and table slicers, which read it from the bound column's auto filter. Styling XLibur has no model for, such as a custom slicer style name, is reported rather than dropped.

- **Slicers can now be created.** `IXLWorksheet.Slicers.Add(pivotTable, fieldName)` adds a slicer that filters a pivot table; `Add(table, columnName)` adds one that filters a table column, matching what EPPlus offers. A created slicer is placed to the right of whatever it filters, and `IXLSlicer.Position` moves it — it is anchored to the grid like any other drawing, so it travels when rows or columns are inserted above it. Moving a slicer read from a file shifts both of its corners together, so it keeps its size. Adding a slicer to a sheet that already carries one puts it in a slicers part of its own, exactly as Excel does, so the part holding the existing slicer is never opened and passes through the save byte for byte.

- **A loaded slicer's caption, style, column count, row height and caption visibility can now be changed.** The change is patched into the part the slicer was read from rather than the part being regenerated, so everything alongside the edited attribute survives — the slicer's `xr10:uid`, its `startItem`, its extension list. A slicer nobody assigns to is not written to at all: its part is not even opened, and comes through a save byte for byte. Changing a slicer's *selection* is not supported yet, because it has to move the pivot table's item visibility with it.

### 🐛 Bug Fixes

- **A pivot table's "show last column" emphasis no longer follows its column-stripes setting.** The reader read `ShowLastColumn` from the `showColStripes` attribute instead of `showLastColumn`, so the two settings could not vary independently and a round trip could switch the emphasis on or off on its own.

- **A pivot table's `Title` and `Description` now survive a save and reload.** Both are public and settable, but were persisted nowhere; a reload always came back with them empty.

- **`IXLPivotTable.CopyTo` no longer silently resets thirty-five settings** — including compact/outline form, visual totals, the grand-total caption and the data caption — that both the reader and the writer already carry. The hand-written copy list carried 29 of the definition's 65 attributes and dropped the other 36; the copy now uses the same attribute-by-attribute description the reader and writer are driven from, so a setting can no longer be present in the round trip and dropped by copy. One of the 36 is deliberately still not copied: `chartFormat` is the next free id for the pivot charts pointing at the source table, and a copy has none of them.

- **A pivot table whose row or column axis references the values field now keeps the `dataPosition` attribute across a reload.** The loader never set it — a documented crash condition the subsystem's own notes call out — so a file that needed it could be re-saved without it and fail to open in Excel without a repair prompt.

- **A pivot table XLibur creates is now stamped with a version that supports slicers.** `createdVersion` and `updatedVersion` were left at zero on a created pivot table, and the writer omits an attribute sitting at its default, so both were absent. Excel reads a pivot table stamped version 0 as one written before the features that came later and silently refuses to attach a slicer to it — the slicer, its cache, its registration and its drawing are all written correctly and the panel simply never appears, with no repair prompt and no validation error. A pivot table loaded from a file still keeps whatever version its file declares.

- **Deleting a pivot table now deletes its part.** Deleting a worksheet already took its pivot table parts along, but deleting a pivot table on its own left the part in the package, still holding a `cacheId` pointing into a `pivotCaches` element that the same save then rebuilt without it. Excel offered to repair the result. Found while building the slicer cascade, whose purpose is not to leave orphans of exactly that kind.

- **A dynamic-array or regression function no longer refuses a cell reference in a scalar argument.** `SEQUENCE(A1)`, `TAKE(range, B1)`, `SORTBY(range, by, C1)`, `XLOOKUP(k, a, b, , C1)`, `LINEST(y, x, D1)` and every other scalar slot across SEQUENCE, TAKE, DROP, SORT, SORTBY, UNIQUE, EXPAND, TOCOL, TOROW, WRAPCOLS, WRAPROWS, XLOOKUP, XMATCH, TEXTSPLIT, LINEST, LOGEST, TREND and GROWTH returned `#VALUE!` whenever the argument was a reference instead of a literal — the ordinary way to write these formulas — because each function's own copy of "reduce this to a scalar" stopped at the first cell reference it saw rather than reading it. These now share one reduction on the value model, which also implements the implicit intersection a whole-row or whole-column reference needs; a multi-area reference still reports `#VALUE!`, as Excel does. SORTBY's order argument had a second, independent defect that made a reference there be misread as a new sort key entirely.

  Sharing one reduction also settles two places where those copies disagreed with the rest of the library, both visible without a cell reference in sight. A scalar slot handed a **multi-element array constant** — `XLOOKUP({3;5}, a, b)`, `LINEST(y, x, {FALSE;TRUE})` — now reduces to that array's first element instead of reporting `#VALUE!`, which is the rule `DATE({2020;2021}, 1, 1)` and every other scalar parameter already followed, and the rule by which a cell stores a non-spilled `=SEQUENCE(3)` as `1`. XLibur does not spill one result per element the way Excel would, so the alternative was an error, not a better answer. Second, a **range** in SORTBY's order slot is now resolved by implicit intersection like any other scalar argument rather than rejected on its shape, so `SORTBY(A1:A2, B1:B2, C1:C3)` reads the order from whichever of `C1:C3` shares the formula's row, and reports `#VALUE!` only when the formula sits outside those rows.

- **A conditional format's font now keeps its name, its family numbering and its character set across a round trip.** Set a conditional format's font to Arial with `XLFontCharSet.Arabic`, save and reopen, and the font came back as the workbook default — Calibri, with no character set. The file itself was correct and Excel rendered it correctly; the three values were dropped on the way back in, so the next save wrote the degraded form and lost them from the file too. The cause was that differential formats (`<dxf>`, which is what a conditional format's styling is stored as) were decoded by a different implementation from ordinary cell styles, and that one searched for the three children by their *rich-text* names — `rFont`, and the `rPr` spellings of `family` and `charset` — which are unrelated types to the `name`, `family` and `charset` a `<font>` actually carries. The same applied to pivot-table formats.

- **A conditional format now keeps its alignment and its protection, and a pivot-table format its protection.** A `<dxf>` may carry six kinds of formatting; the three places that read one each read a different subset, so alignment survived a round trip for pivot-table formats but not for conditional formats, and protection survived for neither. All six are now read wherever a differential format is read.

- **A workbook whose pivot-table format specifies an indent no longer fails to load.** `<alignment horizontal="center" indent="2"/>` is legal in a differential format, but loading it went through the same setter the public API uses, which rejects an indent on anything but a left, right or distributed alignment — so opening the file threw `ArgumentException`. A related quieter case: an indent with no horizontal alignment stated came back as left-aligned, because that setter rewrites it. Both now load as written.

- **A `<border>` that states `diagonalUp` or `diagonalDown` without a `<diagonal>` child is now read the same way everywhere.** Two decoders disagreed: one read the two flags only when the border also carried a `<diagonal>` child, the other read them always. ECMA-376 §18.8.4 makes them attributes of `border` itself, siblings of that child rather than dependents of it, so the unconditional reading is the correct one and is now the only one. Previously the same element could decode to two different styles depending on which path reached it.

- **A workbook that declares the same `numFmtId` twice no longer fails to load.** The custom number formats were collected into a dictionary with `Add`, which throws on a duplicate key, and that collection ran on every load — so such a file could not be opened at all. The first declaration now wins, which is what the linear scan the other number-format lookup used already did.

- **An auto-filter's colour filter no longer changes colour when a workbook is loaded and saved.** The differential formats a workbook writes are rebuilt from scratch on every save, in a fixed order — conditional-format entries first, colour-filter entries last — so the index a colour filter was loaded with is generally not the index it will be written at. A filter column that had not been changed since loading was written back exactly as it came in, stale index and all, and so pointed at whichever differential format now occupied its old slot. In a workbook holding one colour filter and one conditional format the two swap places, and the filter silently took the conditional format's colour. The index is now re-pointed on save; everything else about an unchanged filter column is still written back verbatim.

- **A column range that runs to the last column no longer loses its hidden, collapsed and outline settings.** Excel writes `<col min="2" max="16384" hidden="1"/>` when a user hides or groups from a column rightwards. Every range ending at the last column was treated as the sheet's default-width declaration and skipped, so its width was applied but `hidden`, `collapsed` and `outlineLevel` were dropped — the columns loaded visible and ungrouped. Such a range is now expanded onto the columns it covers. A range that states only a width and a style is still treated as the sheet default and is not expanded, which matters because expanding one would materialise all 16,384 columns on every load; that is the shape XLibur's own writer emits, so its own files are unaffected either way.

- **A pivot table's cells are no longer blanked by a round trip.** Loading a workbook cleared the cells every pivot table had been rendered into, on the understanding that the output belonged to the pivot table rather than to the sheet. Nothing put them back — XLibur writes the definition of a pivot table and never its output — so loading a workbook and saving it, with no edit at all, produced a file whose pivot tables were empty in Excel until the user clicked one. The rendered cells are now left where Excel put them. The alternative, asking Excel to rebuild the table on open, was tried and rejected: it discards the filters applied through hidden items ([ClosedXML#2219](https://github.com/ClosedXML/ClosedXML/issues/2219)) and fails with *Reference isn't valid* when the cache's source is gone. **This reverses the behaviour introduced by [ClosedXML#856](https://github.com/ClosedXML/ClosedXML/pull/856)**, so cells inside a pivot table's range now read back as ordinary cell values rather than as blanks.

- **A sheet with grouped rows now reports its row outline depth, and stops inflating its column one.** Setting a row's outline level counted into the *column* outline counter — the row property was a verbatim copy of the column one and kept the column calls. Two things followed: `sheetFormatPr/@outlineLevelRow` was never written at all, so a sheet with grouped rows declared a maximum row outline of zero, and every row group instead raised `@outlineLevelCol`, so a sheet with no grouped columns could still claim column outlines. Because the load path also sets a row's outline level, this was a round-trip corruption rather than a create-path one: opening a file with row groups and saving it inflated that file's `@outlineLevelCol`. Files XLibur saved before this release carry the wrong attribute; opening and re-saving them corrects it. Per-row `row/@outlineLevel` was always written correctly and is unaffected — only the sheet-level summary attributes were wrong. Saving a sheet whose rows were all ungrouped also threw `InvalidOperationException`, because the counter's emptiness check looked at the whole tally rather than at the levels still in use; ungrouping leaves a zero behind. Separately, inserting below a grouped row or after a grouped column copies that line's outline level onto the new one, and the copy was written straight to the backing field, so the sheet's outline counter never saw it — group some rows, insert below one of them, then ungroup the originals, and the sheet declared no outline level while the inserted line still claimed one. **This half of the fix applies to columns too**, where it has been present for as long as the column counter has worked.

- **`IXLColumn.CellCount()` returns 1,048,576 instead of 1.** It was a verbatim copy of `IXLRow.CellCount()`, down to measuring the *column* span of its address. A column's address runs from row 1 to the last row of a single column, so both ends carried the same column number and the count was always `c - c + 1`. A column now reports one cell per row in the sheet, which is what the row half has always reported for its own axis. The signature is unchanged; only the value it returns moves — which means the compiler cannot warn you: **code that loops over a column with `for (var i = 1; i <= column.CellCount(); i++)`, or calls `column.Column(2, column.CellCount())`, silently goes from one iteration to 1,048,576 and materialises the whole column.** Check any such loop written against `IXLColumn` before upgrading. `IXLRangeColumn.CellCount()` is a different method and is unaffected.

- **Inserting rows or columns no longer throws when the sheet's last used line is near its edge.** Carrying line sizes forward walked from the last used row (or column) down to the insert point and wrote each height to `row + numberOfRows` without clamping, so a sheet with anything on a row within `numberOfRows` of row 1,048,576 addressed row 1,048,577 and threw `ArgumentOutOfRangeException`. Both axes were affected, identically. A line pushed past the edge now takes its height or width off the sheet with it, which is what happens to its contents.

- **A frozen pane is now written as `state="frozen"`, and the unsplit axis is left out.** `XLWorkbook.SaveAs` wrote `state="frozenSplit"` on every pane it emitted, and `xSplit="0"` or `ySplit="0"` for an axis that was not split; `XLStreamingWorkbook` wrote `state="frozen"` and omitted the unused axis. Same freeze, two different files. Excel writes `frozen` for a pure freeze and reserves `frozenSplit` for a pane frozen out of an existing manual split, which XLibur never produces — of the 30 `<pane>` tags across the test corpus, 27 are `frozen`, and the 3 that are not sit inside `<customSheetView>` elements whose main sheet view is `frozen`. Nothing caught it because the reader accepts both spellings and maps them to the same model, so a load-and-save comparison could not see the difference; the tests that pin this now read the bytes of both packages and compare them attribute by attribute. Files XLibur saved before this release keep loading exactly as they did — the reader is unchanged and still accepts both — and re-saving one rewrites the pane in the form Excel uses.

- **Chart and note anchors now move when rows or columns are inserted or deleted above or to the left of them.** A chart anchored at row 10 stayed at row 10 after an insert at row 1, and a note's callout box detached from the note: the note itself moved, because notes live in the misc slice and the slice shifts with the cells, while the box it points from did not — so the callout ended up hanging three rows above the cell it belonged to. Both anchors are a pair of raw integers, and nothing in the library told them an edit had happened. They now move exactly as a picture's anchor already did, including the cases a picture already handled: an edit *inside* a two-cell anchor grows or shrinks it rather than moving it, an edit outside the anchor's columns leaves it alone, and a delete that swallows the anchor entirely clamps it to the deletion point rather than destroying the chart. An absolutely anchored drawing is still pinned in place, because its position carries no cell reference for the grid to move. **This changes the file XLibur writes for an existing document that is edited structurally**: a chart or note below an insert now lands where the cells it annotates landed. ([spec 33](docs/specs/33-sheet-listener-seam.md))

- **Frozen and split panes, and pivot table positions, now move when rows or columns are inserted or deleted above or to the left of them.** A sheet frozen at row 5 still reported a freeze at row 5 after three rows were inserted at row 1, so the freeze cut through the middle of the header it had been placed to hold; a pivot table anchored at `D10` still claimed `D10` while the cells Excel had rendered it into moved to `D13`. `SplitRow` and `SplitColumn` count frozen lines rather than naming one, so they grow and shrink with an edit *inside* the frozen region and are untouched by an edit below it — and deleting every line above the split leaves no split, which is written as no pane at all. Inserting cells into part of a row and shifting them down is not a row insert and still does not move a freeze. A pivot table's *source* range is a separate thing, re-pointed through the pivot cache as before, and is unaffected. **This changes the file XLibur writes for an existing document that is edited structurally.** ([spec 33](docs/specs/33-sheet-listener-seam.md))

- **A hyperlink no longer detaches from its cell when rows or columns are inserted into a range more than one line tall.** Put a hyperlink on `A5` and call `ws.Range("A1:A5").InsertRowsAbove(3)`, and the cell's content arrived on `A8` while its hyperlink arrived on `A12`. The cells are moved by one rectangle — the edited range's leading edge extended by the number of lines inserted, `A1:A3` here — but the listeners that reposition hyperlinks were handed a different one, the *whole* edited range extended the same way, which is `A1:A7`. A hyperlink was therefore moved seven rows while the cell under it moved three. The two rectangles are identical whenever the edited range is exactly one line tall on the axis being shifted, which is what `ws.Row(n).InsertRowsAbove(k)` and `ws.Column(n).InsertColumnsBefore(k)` both produce, so the whole of the ordinary API was unaffected and the defect only ever showed through a multi-line range. Deletes were already correct and are unchanged: a delete removes the whole edited range and shifts by that range's own line count, so its two rectangles already agreed.

- **A note XLibur creates is now written as moving and sizing with its cell, which is what its own object model always said.** A note stated its anchoring mode twice and the two disagreed on every note created: `XLComment.Anchor` said move-and-size-with-cells, while `Style.Properties.Positioning`, inherited from the default note style, said absolute. The VML writer read the style, so every note XLibur wrote was pinned to the sheet rather than to its cell — which is not Excel's default and not what the library reported when you asked it. The two are now one stored value under two names, so they cannot diverge again, and the default is move-and-size-with-cells. A note's anchoring mode also survives a round trip now: it was written from the style but read back into the style only, leaving `Anchor` at its default on every loaded workbook. **This changes the bytes XLibur writes for every note it creates** — `<x:MoveWithCells>` and `<x:SizeWithCells>` both go from `True` to `False`, which in VML's inverted spelling means the note does move and does resize with its cells. A note whose positioning you set explicitly is written exactly as before.

- **An unfrozen split pane now survives a round trip, and a split is no longer written back as a freeze.** A sheet carrying `<pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="split"/>` — Excel's draggable split bar, what View → Split gives you without Freeze Panes — loaded with `SplitRow` and `SplitColumn` both at zero, because the reader took the two splits only when the state was `frozen` or `frozenSplit`. Saving then wrote no `<pane>` at all, so the split was gone from the file too. In the other direction, `SplitRow` and `SplitColumn` are public, so a caller can ask for a split without ever calling `Freeze`, and every pane XLibur emitted was `state="frozen"` — a frozen pane cannot be dragged and does not scroll independently, so the round trip silently changed how the sheet behaves. The model can now tell the two apart: **`IXLSheetView.FreezePanes` is a public property again**, set by the three `Freeze` methods, left alone when you assign the splits yourself, and — new — set by the loader, which never touched it before. A pane loaded as `frozenSplit` reads as frozen and saves as `frozen`, the same normalization introduced when XLibur stopped writing `frozenSplit` for every pane.

  The rest of the library follows the same distinction: an unfrozen split is not moved by row or column edits, is not treated as a frozen band by `FocusCell`, is dropped rather than reinterpreted when `FreezeRows`/`FreezeColumns` freezes the other axis, and anchors its `topLeftCell` at `A1` rather than at split + 1, which is a cell address only for a freeze. The caller-visible half of this — what `SplitRow` and `SplitColumn` now mean, and what to call instead — is under Breaking Changes above.

- **Editing a cell now recalculates every dependent, even when something else already marked one of them dirty first.** Calling `InvalidateFormula`, renaming a sheet, moving a range, or shifting references on a row/column insert could each leave a formula flagged dirty for a reason unrelated to the current edit; the recalculation walk mistook that pre-existing flag for "already visited by this walk," stopped there, and left everything downstream silently stale. `A1=1, B1=A1+1, C1=B1+1, D1=C1+1; C1.InvalidateFormula(); A1.Value = 10` used to leave `D1` at `4` instead of `13`.

  This costs something on one shape. Treating "already dirty" as "already visited" also stopped the walk early *between* edits, so writing many cells that feed the same model without reading anything in between used to mark that model dirty once and skip it thereafter. Each such write now walks the model again, because there is no longer any sound way to tell "dirty because a walk reached it" from "dirty for some unrelated reason" — the conflation that caused the bug. Writing 20,000 inputs of a 50-formula-deep shared model with no intervening read goes from roughly 50 ms to 190 ms; reading or saving between edits, or writing cells nothing depends on, is unaffected and now allocates less than before.

## v0.311.1 - 2026-08-11

A dependency release: `XLibur.Fonts.SkiaSharp` moves to SkiaSharp 4.151.1, a patch bump. No new features, no bug fixes and no breaking changes.

### 🔧 Dependencies

- **`XLibur.Fonts.SkiaSharp` moves to SkiaSharp 4.151.1** (from 4.151.0), with the Linux native asset package following. The font engine decides text metrics, so the thing to watch in a bump like this is column widths and row heights moving; the tests covering them pass unchanged. ([#384](https://github.com/XLibur/XLibur/pull/384), [#385](https://github.com/XLibur/XLibur/pull/385) by [@dependabot](https://github.com/apps/dependabot))

## v0.311.0 - 2026-08-08

A performance release for cell styles, with two breaking changes. Style data now uses about half the memory it used before, and XLibur compares and looks up styles faster. `XLColor.Color` no longer returns a named colour, and the two alignment style types are now internal; each needs a small code change if you use it. Four bug fixes: an edge with no border line now always reports the same colour, setting an edge's colour before its line style no longer loses the colour, and an out-of-range value cast into a border, colour or alignment enum is now rejected instead of silently aliasing a different one.

### ⚠️ Breaking Changes

#### Colors

- **`XLColor.Color` returns a plain colour, not a named one.** The alpha, red, green and blue values are the same as before. What changes is that the result is no longer a *known* colour such as `Color.Red`: its `Name` is now a string of hex digits, and `IsNamedColor` is false. `Color.Equals` checks that name as well as the four channels, so compare ARGB values instead. XLibur reads only the four channels, and a colour loaded from a spreadsheet has no name to keep. ([#379](https://github.com/XLibur/XLibur/pull/379) by [@jafin](https://github.com/jafin))

  ```csharp
  // Before
  Assert.AreEqual(Color.Red, xlColor.Color);

  // After
  Assert.AreEqual(Color.Red.ToArgb(), xlColor.Color.ToArgb());
  ```

#### Styles

- **`XLAlignmentKey` and `XLAlignmentValue` are now internal.** XLibur builds a style from eight parts: alignment, border, fill, font, number format, protection, colour, and the style itself. The other seven have always been internal. Alignment was public on its own, and by accident — no public member returns an `XLAlignmentValue`, and neither type appears in the list of types XLibur undertakes to keep. Read and set alignment through `IXLStyle.Alignment`, which does not change. ([#379](https://github.com/XLibur/XLibur/pull/379) by [@jafin](https://github.com/jafin))

  ```csharp
  // Before
  var alignment = XLAlignmentValue.FromKey(ref key);
  var wraps = alignment.WrapText;

  // After
  var wraps = cell.Style.Alignment.WrapText;
  ```

### ⚡ Performance

#### Styles

- **Style data uses about half the memory it used before.** XLibur identifies each distinct cell format by a style key, and copies that key whenever you change a format and again on every cache lookup. Colours filled most of the key, and most of each colour went unread: XLibur uses the four ARGB bytes, but each colour also carried a whole `System.Drawing.Color` beside separate palette, theme and tint fields. A colour is only ever one of those things at a time, so the four now share one field. A border key is a third of its former size, and a full style key about half. ([#379](https://github.com/XLibur/XLibur/pull/379) by [@jafin](https://github.com/jafin))

  Hashing a border key is about twice as fast, and fill and colour keys about 15% faster. Part of that comes from reading less memory rather than doing less work, so it helps most in a workbook that holds many different styles. Noise on the test machine hides the change in the end-to-end styling benchmarks, so this release makes no claim for them.

- **Saving a workbook compares styles faster.** Each part of a style worked out its hash every time something asked for one. Saving uses these parts as dictionary keys, so each lookup of a border hashed all five of its colours again. Every part now works out its hash once, when XLibur creates it. Two parts that are the same object also count as equal without a comparison of their contents, which is the usual case because XLibur keeps one shared copy of each. ([#379](https://github.com/XLibur/XLibur/pull/379) by [@jafin](https://github.com/jafin))

- **Setting a border on a range, row, column or worksheet does less work.** XLibur stored the new border in its cache, then built a style that stored the same border again. It also applied the change twice. It now takes the stored border from the finished style. ([#379](https://github.com/XLibur/XLibur/pull/379) by [@jafin](https://github.com/jafin))

### 🐛 Bug Fixes

#### Styles

- **An edge with no border line always reports the same colour.** Two borders that differ only in the colour of an edge with no line have always counted as equal, so XLibur kept whichever it saw first. Setting a colour on such an edge did nothing, which is correct — Excel draws no line there, and writes no colour — but the colour you read back depended on what the workbook had done earlier. These edges now always report black. ([#379](https://github.com/XLibur/XLibur/pull/379) by [@jafin](https://github.com/jafin))

  A file can give an edge a colour but no line style, because the two are independent. XLibur now treats such a border the same way, so it matches the border already in the file instead of saving a second copy of it.

- **Setting an edge's colour before its line style no longer loses the colour.** The fix above means a colour set on an edge with no line style is dropped there and then, so `border.TopBorderColor = Red; border.TopBorder = Thin;` used to end up black instead of red — the second line had nothing left to colour. XLibur now holds such a colour until the edge is given a line style, and applies it then, so the two lines give the same result in either order. Setting the colour on the outside or inside of a range in one call, before the matching line style, still loses the colour as before; set the line style first for those. ([#379](https://github.com/XLibur/XLibur/pull/379) by [@jafin](https://github.com/jafin))

- **An out-of-range enum value cast into a border or colour property is rejected instead of silently becoming a different one.** Border line style, colour kind, and theme colour are each stored in a single byte to keep a style small. Casting an invalid value into one of the public enums that back them — `(XLBorderStyleValues)999`, for example — used to wrap around into whichever defined member the value happened to reduce to, silently applying a style or colour you never asked for. XLibur now rejects such a value with `ArgumentOutOfRangeException` instead. ([#379](https://github.com/XLibur/XLibur/pull/379) by [@jafin](https://github.com/jafin))

- **An out-of-range enum value cast into an alignment property is rejected too, where XLibur can tell.** `XLAlignmentHorizontalValues`, `XLAlignmentVerticalValues` and `XLAlignmentReadingOrderValues` were already declared to fit in a byte, rather than narrowed into one the way the border and colour enums above are, so most invalid values now throw `ArgumentOutOfRangeException` the same way. One case cannot be caught by any check: a cast that wraps around onto a real member — `(XLAlignmentHorizontalValues)262`, which wraps to `Left` — is, by the time XLibur sees it, genuinely that member, indistinguishable from a caller who passed `Left` directly. That gap exists only for these three alignment enums, not the border or colour ones, because their narrowing happens inside the caller's own cast rather than inside XLibur. ([#379](https://github.com/XLibur/XLibur/pull/379) by [@jafin](https://github.com/jafin))

## v0.310.0 - 2026-08-06

### ⚡ Performance

#### Saving

- **Saving a workbook that was loaded from a file no longer re-materialises every stored row and cell.** `GetWorksheetDom` called `LoadCurrentElement` on `<worksheet>` whenever the stored part was non-empty, building an OpenXML DOM of the entire sheet — which `StreamToPart` then discarded, because cells are streamed from the value slices instead. The intent was already recorded in the code ("Sheet data is not updated in the Worksheet DOM here, because it is later being streamed directly to the file … especially problematic for large sheets") but it was defeated on every save of a loaded workbook, and on every save after the first. The part is now copied with `<sheetData>` reduced to an empty element and the DOM built from that, so stored rows are tokenised once by a raw reader rather than becoming objects nobody reads. Re-saving an untouched 20,000 × 21 workbook went from 1,729.6 ms and 333.68 MB to 991.7 ms and 88.83 MB; a 10-sheet template round trip from 10.42 ms and 3.89 MB to 8.42 ms and 3.17 MB. ([#373](https://github.com/XLibur/XLibur/pull/373) by [@jafin](https://github.com/jafin))

  Copying the part rather than assembling the root element by hand is deliberate, and costs about 8% of the win. A child parsed on its own records the namespace declarations it needs instead of inheriting the root's, so the faster version round-tripped `mc:AlternateContent` subtrees as `<controls xmlns="…">` where the file had `<x:controls>`. Copying keeps the SDK doing one parse of one document, so the emitted bytes are unchanged by construction.

- **Fixed per-workbook and per-worksheet save cost is lower**, worth ~4% of total allocation on an open-and-save round trip of a 10-sheet template. Reading the theme no longer materialises the entire theme part — font scheme, format scheme, gradient fills, effect styles — to pull twelve hex colours. The worksheet and sheet-view content managers make one pass over their child list instead of 39 filtered traversals of it per worksheet per save, and their slot lookup no longer allocates on every emitted element. The conditional-formatting writer now discovers that a worksheet has no conditional formatting before building a hash set, cast, concat, ordering and list rather than after. ([#373](https://github.com/XLibur/XLibur/pull/373) by [@jafin](https://github.com/jafin))

#### Loading

- **The loader skips `<sheetData>` in one operation instead of once per row.** Each worksheet part is read twice: once for structural elements with the SDK reader, once for `<sheetData>` with a raw reader, which is far cheaper per cell. The first pass has to get past `<sheetData>`, and it was doing so a row at a time — and `Skip` on a start element runs the element factory and builds an attribute list, so each row cost a throwaway `Row` object plus attributes. On a 20,000 × 20 sheet the skip alone was 117 ms and 4.9 MB, against 268 ms to read the cells for real. Loading a 20,000 × 21 workbook went from 460.2 ms and 60.14 MB to 405.6 ms and 55.40 MB. ([#373](https://github.com/XLibur/XLibur/pull/373) by [@jafin](https://github.com/jafin))

  Allocation is exact and consistently 5–12% lower across the load benchmarks. Times carry wide error bars on the benchmark machine; a probe sweeping sheet shapes put the load-time saving at 10–37%, largest on many-rows-few-columns sheets.

#### Styles

- **Setting a style property on a cell costs ~33% less CPU, at ~219 ns per styled cell rather than ~328 ns.** All six style facades — number format, font, fill, border, alignment, protection — interned each new component key in its repository and then handed it straight to the style, which hashed the same key again. The interning was waste whichever way the modification went: on a transition-cache hit the component value is never needed, and on a miss the component is interned anyway further down. Taking the interned value back off the resulting style leaves the facade just as correct for later reads, at no lookup. Allocation is unchanged, since a repository probe of an existing entry does not allocate. ([#373](https://github.com/XLibur/XLibur/pull/373) by [@jafin](https://github.com/jafin))

### 🐛 Bug Fixes

#### Styles

- **A cell keeps the number format, font or border it was given when two style keys happen to share a hash code.** The per-style transition cache matched an entry on a 32-bit hash and never compared the key itself, so two distinct component keys colliding in both hash and cache slot returned each other's style — a cell silently ending up with formatting it was never assigned, and nothing in the file to explain it. This is reachable through the public API rather than merely theoretical: a custom number format pins `NumberFormatId` to a constant, so the key's hash reduces to the hash of the format string, and a workbook applying k distinct custom formats to default-styled cells collides with odds near k²/2^33 — roughly 1 in 8,000 at k=1,000. Font names are exposed the same way. Because string hashing is randomised per process, an unchanged workbook round-trips correctly on most runs and corrupts on others. A hit is now confirmed by comparing the key, with the hash kept as the cheap reject so the miss path is unchanged. ([#374](https://github.com/XLibur/XLibur/pull/374) by [@jafin](https://github.com/jafin))

  The cache also held its hash, key and result in three parallel arrays, where two threads storing different keys into one slot could interleave their writes and leave an entry whose key and result came from different transitions — again a wrong style rather than the documented cache miss. The three now travel in one immutable entry published by a single reference write, so a reader sees an entry whole or not at all.

  Alignment and protection keys were never exposed to this: a brute force over 11.4M alignment keys spanning Excel's full permitted indent and rotation ranges found no collision at all. Only the string-bearing keys were affected.

## v0.301.0 - 2026-08-04

### 🐛 Bug Fixes

#### Formulas

- **A dynamic-array formula keeps its spill through a save.** A workbook holding `=UNIQUE(...)` or any other spilling formula came back with the anchor value intact and `#VALUE!` in every cell below it. Excel writes a spilled cell as a cached formula result (`t="str"`), and reads a shared-string or inline-string cell inside a spill footprint as content occupying the range. XLibur decided whether text was a formula result purely from the presence of an `<f>` element, and a dynamic array carries its formula only in the anchor. So the spilled cells were taken for text constants, interned into the shared string table, and written back as `t="s"`. Data tables lost their footprint the same way. ([#370](https://github.com/XLibur/XLibur/pull/370) by [@jafin](https://github.com/jafin))

  Classic array formulas were never affected: the formula slice holds one across its whole range, so every cell of one already saved as a formula cell.

## v0.300.0 - 2026-08-02

Promotes to public API the three core capabilities `XLibur.Report` previously reached through
`InternalsVisibleTo`: `XLFunctionLibrary`, for evaluating a worksheet function without a grid;
source inspection and re-pointing on `IXLPivotCache`; and `IXLConditionalFormat.SetRanges`.
Deleting a set of rows is up to 67× faster and allocates 670× less. Every package now ships a
Software Bill of Materials. Two breaking changes: adding members to `IXLPivotCache` and
`IXLConditionalFormat` is source-breaking for anyone implementing those interfaces outside the
library, and `IXLRange.Cell(string)` now throws for an address it cannot resolve instead of
returning null (see Breaking Changes). Plus 2 bug fixes, covering range-address validation and
rich-text equality.

### ⚠️ Breaking Changes

#### Pivot tables and conditional formatting

- **`IXLPivotCache` and `IXLConditionalFormat` gain members**, which is source-breaking for any type outside the library implementing them. Neither interface is designed to be implemented externally — each has a single implementation with an internal constructor — and XLibur is pre-1.0. Consumers that only *use* these interfaces are unaffected. ([#354](https://github.com/XLibur/XLibur/pull/354) by [@jafin](https://github.com/jafin))

#### Ranges

- **`IXLRange.Cell(string)` throws `ArgumentException` for an address it cannot resolve, where it previously returned `null`.** The interface has always been annotated non-null, so the null contradicted the signature and typically surfaced as a `NullReferenceException` further downstream. It now fails at the call site, matching `IXLWorksheet.Cell(string)`, which already behaved this way. ([#360](https://github.com/XLibur/XLibur/pull/360) by [@jafin](https://github.com/jafin))

  ```csharp
  // Before — a null return was the only signal that the address was unresolvable
  var cell = range.Cell(name);
  if (cell is null)
      return;

  // After
  try
  {
      var cell = range.Cell(name);
  }
  catch (ArgumentException)
  {
      return;
  }
  ```

  Not a mechanical swap: if you relied on the null to mean "not found", the equivalent is now catching `ArgumentException`. Code that trusted the non-null annotation needs no change. `IXLRange.Range(string)` gained the same guard, but nothing reaches it — a bad address already throws while being parsed.

### 🔒 Supply Chain

- **Every package now ships a Software Bill of Materials.** Each `.nupkg` embeds an SPDX 2.2 manifest under `_manifest/spdx_2.2/`, generated at pack time, so the bill of materials travels with the package to anyone installing it from NuGet. Each release additionally carries a CycloneDX 1.7 document per package (`XLibur.<version>.cdx.json`) as a GitHub Release asset. ([#353](https://github.com/XLibur/XLibur/pull/353), [#356](https://github.com/XLibur/XLibur/pull/356) by [@jafin](https://github.com/jafin))

  ```bash
  # Read the embedded manifest without installing
  unzip -p XLibur.0.300.0.nupkg '_manifest/spdx_2.2/manifest.spdx.json' | jq .
  ```

### ✨ New Features

#### Formula functions

- **`XLFunctionLibrary` evaluates one of Excel's ~400 worksheet functions without a workbook.** Previously the calculation engine was reachable only through a cell, so computing `SUM` over values you already hold in memory meant building a grid to put them in. An instance is safe to share across threads, which is worth doing: constructing one builds the whole function table, at ~12 MB. ([#354](https://github.com/XLibur/XLibur/pull/354) by [@jafin](https://github.com/jafin))

  ```csharp
  using XLibur.Excel.CalcEngine;

  var library = new XLFunctionLibrary();               // or pass a CultureInfo
  ReadOnlySpan<XLCellValue> args = [1.0, 2.0, 3.0];

  if (library.TryInvoke("SUM", args, out var result))
      Console.WriteLine(result);                       // 6
  ```

  This calls a function; it does not parse a formula — there are no cell references or ranges, so a function that takes a range in a cell formula takes those cells as a run of arguments here. `TryInvoke` returns `false` only when no function has that name; a call that was made but could not succeed (wrong arity, wrong argument type, division by zero) returns `true` with an `XLError` result, as Excel reports it. Functions that need a worksheet to be relative to — `ROW`, `OFFSET`, `INDIRECT` — throw `XLNoWorksheetContextException`, and a result that is an array or a reference comes back as `#VALUE!`. To evaluate a formula string against real data, use `IXLWorksheet.Evaluate` instead.

#### Pivot tables

- **A pivot cache now reports where its data comes from, and can be re-pointed at a new range.** `SourceKind` distinguishes the six source kinds Excel defines, so a source XLibur cannot resolve (a connection, a consolidation, an external workbook, a scenario) is distinguishable from a named range that no longer exists. `SourceRange`, `SourceName` and `SourceWorksheet` expose the source itself, and `SetSourceRange` moves the cache to a new one — useful after generated rows have grown past the range the template defined. ([#354](https://github.com/XLibur/XLibur/pull/354) by [@jafin](https://github.com/jafin))

  ```csharp
  var cache = pivotTable.PivotCache;

  if (cache.SourceKind == XLPivotSourceKind.Range)
      cache.SetSourceRange(worksheet.Range("A1:D500"));
  ```

  `SourceRange`, `SourceName` and `SourceWorksheet` return `null` for kinds that cannot resolve to a range, rather than throwing — check `SourceKind` first when the distinction matters.

#### Conditional formatting

- **`IXLConditionalFormat.SetRanges` replaces the ranges a rule covers.** `Ranges` returns a fresh projection each time it is read, so there was previously no way to widen an existing rule over a larger block — the only option was deleting the rule and rebuilding it. ([#354](https://github.com/XLibur/XLibur/pull/354) by [@jafin](https://github.com/jafin))

  ```csharp
  format.SetRanges([worksheet.Range("A1:D500")]);
  ```

### ⚡ Performance

#### Structural edits

- **Deleting a set of rows is up to 67× faster and allocates 670× less.** `IXLRows.Delete()` looped a single-row delete over its rows, so every consumer of a shift — most expensively the formula pass, which visits every formula in the workbook — ran once per row. The rows are now re-pointed against the whole deletion in one pass, and contiguous rows coalesce into one run. The cell compaction moves a row at a time rather than a cell at a time, and the calculation-engine purge that each run repeated is coalesced into one. Deleting every third row of an 8,000-row sheet with a `SUM` per row went from 19,595 ms and 17,481 MB to 294 ms and 26 MB. The growth rate dropped from 4× per doubling to 2×. Row and column *inserts* and single-row deletes get a smaller share of the same win, since formulas the edit cannot reach now skip the parse entirely. ([#347](https://github.com/XLibur/XLibur/pull/347), [#348](https://github.com/XLibur/XLibur/pull/348), [#350](https://github.com/XLibur/XLibur/pull/350) by [@jafin](https://github.com/jafin))

  A workbook holding an array, data-table or dynamic-array formula keeps the row-at-a-time path, because the stored range and spill footprint those carry are relocated by the per-cell shift rather than by rewriting formula text.

### 🐛 Bug Fixes

#### Ranges

- **A null, empty or half-written range address now says what is wrong instead of crashing.** `Range("")` and half-written addresses such as `"A1:"`, `":"` and `"$"` threw `IndexOutOfRangeException`, and a null address threw `NullReferenceException` — internal failures escaping a public API rather than anything a caller could act on. Null now throws `ArgumentNullException`, empty throws `ArgumentException`, and a half-written address throws `FormatException`, joining the malformed-address case that already did. ([#363](https://github.com/XLibur/XLibur/pull/363), [#365](https://github.com/XLibur/XLibur/pull/365) by [@jafin](https://github.com/jafin))

  Deliberately unchanged: `" "` still throws `ArgumentException`, `"not an address"` still `FormatException`, an address past the sheet limits still `OverflowException`, and an unknown name still `ArgumentOutOfRangeException`. The XML docs on `IXLRange.Range(string)` now list all of them. This affects every entry point that parses an address string, including `IXLWorksheet.Range(string)` and `IXLWorkbook.Range(string)`.

#### Rich text

- **Two phonetic runs with identical content now compare equal however they are held.** `XLPhonetics` carried typed `Equals` overloads but never overrode `object.Equals`, and had no `GetHashCode`. Value comparison therefore applied only where the static type was known. An `object`-typed reference, `object.Equals(a, b)` or a non-generic collection fell back to reference equality, and hashing was identity-based. Reachable through `IXLFormattedText.Phonetics`. ([#362](https://github.com/XLibur/XLibur/pull/362) by [@jafin](https://github.com/jafin))

  `GetHashCode` returns a constant, deliberately and consistently with `XLRichString`: every field feeding equality is mutable, so a computed hash would go stale on the first edit. That is correct but degenerate, so avoid using phonetics as dictionary keys where lookup cost matters.

### 🔧 Dependencies

- **`XLibur.Fonts.SkiaSharp` moves to SkiaSharp 4.151.0** (from 4.150.1), with the Linux, macOS and Win32 native asset packages following. The font engine decides text metrics, so the thing to watch in a bump like this is column widths and row heights moving; nothing shifted. ([#351](https://github.com/XLibur/XLibur/pull/351) by [@jafin](https://github.com/jafin))

- **`XLibur` moves to OpenMcdf 3.2.0** (from 3.1.4). ([#342](https://github.com/XLibur/XLibur/pull/342) by [@dependabot](https://github.com/apps/dependabot))

## v0.200.0 - 2026-08-01

### ⚠️ Breaking Changes

#### Colors and styles

- **`XLColorType` members are renumbered.** `Automatic` takes ordinal 0, so `Color` moves 0 → 1, `Theme` 1 → 2 and `Indexed` 2 → 3. This is source compatible but binary breaking, and breaks anything that persists the numeric value — stored settings or serialized styles written by an earlier version need remapping. ([#232](https://github.com/XLibur/XLibur/pull/232) by [@jafin](https://github.com/jafin))

- **`XLColor.NoColor.Color` (and `.Indexed`/`.ThemeColor`) now throw** instead of returning a meaningless all-zero ARGB. Test with `IsAutomatic` before reading a color component. ([#232](https://github.com/XLibur/XLibur/pull/232) by [@jafin](https://github.com/jafin))

  ```csharp
  // Before — returned Color.FromArgb(0, 0, 0, 0) for an automatic color
  var rgb = cell.Style.Font.FontColor.Color;

  // After
  var color = cell.Style.Font.FontColor;
  var rgb = color.IsAutomatic ? defaultRgb : color.Color;
  ```

  There is no ARGB that means "automatic", which is why the property now throws rather than
  inventing one — decide what your code should render for a color Excel resolves at display time.

### ✨ New Features

#### Writing large files

- **Streaming write API (`XLStreamingWorkbook`)**: a forward-only writer for exports too large to hold in memory. Rows are serialized straight into the file as they are appended, so nothing is retained per row — a million rows × ten columns costs ~108 MB of peak managed heap, where `XLWorkbook` needs roughly that much for a *tenth* as many rows. On the 50K-row benchmark it is also 1.6× faster than `XLWorkbook` and allocates a fifth as much.

  ```csharp
  using XLibur.Excel.Streaming;

  using var workbook = XLStreamingWorkbook.Create("Large.xlsx");
  var sheet = workbook.AddWorksheet("Data");
  sheet.FreezeRows(1);
  sheet.AppendRow("Name", "Amount");

  for (var i = 0; i < 1_000_000; i++)
      sheet.AppendRow($"Item {i}", i * 1.5);

  workbook.Finish();   // required — disposing without it abandons the write
  ```

  The trade is that it is append-only: rows go in ascending order, one worksheet at a time, nothing can be read back or revised, and formulas are stored verbatim rather than evaluated. Streamed sheets support column widths, freeze panes, an autofilter range, row and per-cell styles, and formulas with cached values. Anything beyond that — tables, merges, conditional formats, pivot tables, drawings — still needs `XLWorkbook`.

  Memory is flat in the number of rows, but not unconditionally flat: under the default `SharedStrings` mode each distinct string is held until `Finish()`, so cost tracks how many *distinct* text values there are rather than how many rows. A million rows of repeating labels is cheap; a million distinct ones is not. `XLStreamingStringStorage.Inline` removes that term entirely at the cost of a larger file, and takes the worst case — every row carrying a distinct string — from 108 MB to 14 MB.

  Because the writer assembles the package itself rather than going through `System.IO.Packaging`, the destination stream does not have to be seekable — a workbook can be written straight to an HTTP response, which `XLWorkbook.SaveAs` cannot do. ([#263](https://github.com/XLibur/XLibur/pull/263) by [@jafin](https://github.com/jafin))

- **`SaveOptions.CompressionLevel`**: choose how hard the package is compressed on an ordinary save. `CompressionLevel.Fastest` trades a larger file for a quicker save, `NoCompression` skips it entirely. Applies to parts the save creates; re-saving a workbook loaded from an existing file leaves its existing parts alone. `XLStreamingOptions.CompressionLevel` does the same for streamed writes, where `Fastest` is ~1.7× quicker than the default. ([#263](https://github.com/XLibur/XLibur/pull/263) by [@jafin](https://github.com/jafin))

#### Formula functions

- **Regression and descriptive statistics — 23 functions**: `LINEST`, `LOGEST`, `TREND`, `GROWTH`, `FREQUENCY`, `FORECAST`, `FORECAST.LINEAR`, `CORREL`, `PEARSON`, `COVARIANCE.P`, `COVARIANCE.S`, `COVAR`, `SLOPE`, `INTERCEPT`, `RSQ`, `STEYX`, `SKEW`, `SKEW.P`, `KURT`, `PROB`, `TRIMMEAN`, `HARMEAN` and `AVEDEV`. The five array-returning ones spill.

  `LINEST` handles several predictors, not just one, and reports the full five-row statistics block on request. Both orientations work — y down a column with each predictor in its own column, or y across a row with each in its own row. ([#257](https://github.com/XLibur/XLibur/pull/257) by [@jafin](https://github.com/jafin))

- **The statistical distributions — 69 functions**. The modern dotted set (`NORM.DIST`, `NORM.INV`, `NORM.S.DIST`, `NORM.S.INV`, `LOGNORM.*`, `CHISQ.*`, `F.*`, `T.DIST*`, `EXPON.DIST`, `POISSON.DIST`, `WEIBULL.DIST`, `GAMMA`, `GAMMA.DIST`, `GAMMA.INV`, `GAMMALN(.PRECISE)`, `BETA.DIST`, `BETA.INV`, `HYPGEOM.DIST`, `NEGBINOM.DIST`, `BINOM.INV`, `CONFIDENCE.NORM`, `CONFIDENCE.T`, `PERCENTILE.EXC`, `QUARTILE.EXC`, `RANK.AVG`, `MODE.MULT`, `PERCENTRANK(.INC/.EXC)`), the hypothesis tests (`CHISQ.TEST`, `F.TEST`, `T.TEST`, `Z.TEST`) and the 26 pre-2010 names (`NORMDIST`, `CHIDIST`, `FINV`, `TDIST`, `CRITBINOM`, …). The pre-2010 names are registered against the same implementations, not copies of them.

  `MODE.MULT` spills. Everything reduces to one of four special functions — the regularized incomplete gamma, the regularized incomplete beta, the error function, or elementary functions — so the inverses invert their own distributions to full double precision rather than to the accuracy of a rational approximation. `T.TEST` with unequal variances truncates the Welch degrees of freedom, as the rest of the `T.DIST` family truncates its own. ([#256](https://github.com/XLibur/XLibur/pull/256) by [@jafin](https://github.com/jafin))

- **`AGGREGATE`, `NETWORKDAYS.INTL` and `WORKDAY.INTL`**. `AGGREGATE` covers all nineteen function numbers — including `PERCENTILE.EXC` and `QUARTILE.EXC`, which have no standalone registration yet — and all eight option values, so ignoring hidden rows and error values both work rather than being documented limitations. The `.INTL` date functions take a weekend as either one of Excel's numbered codes or a seven-character Monday-to-Sunday mask. ([#255](https://github.com/XLibur/XLibur/pull/255) by [@jafin](https://github.com/jafin))

- **20 modern text and array-shaping functions**: `TEXTSPLIT`, `TEXTBEFORE`, `TEXTAFTER`, `VALUETOTEXT`, `ARRAYTOTEXT`, `UNICHAR`, `UNICODE`, `DBCS`, `ENCODEURL`, and the array-shaping set `VSTACK`, `HSTACK`, `TOROW`, `TOCOL`, `WRAPROWS`, `WRAPCOLS`, `CHOOSEROWS`, `CHOOSECOLS`, `TAKE`, `DROP`, `EXPAND`. The array-shaping functions and `TEXTSPLIT` spill.

  `DBCS` derives its mapping by inverting `ASC`'s, so the two are exact inverses. Where Excel would report `#CALC!` — a `DROP` that leaves nothing, a `TOCOL` that ignores every value — XLibur reports `#VALUE!` instead, because the value model has no `#CALC!`. ([#254](https://github.com/XLibur/XLibur/pull/254) by [@jafin](https://github.com/jafin))

- **42 engineering functions**: the complex-number family (`COMPLEX` and all 26 `IM*` functions), `CONVERT` with the full unit table, `BESSELI`/`BESSELJ`/`BESSELK`/`BESSELY`, `ERF`/`ERF.PRECISE`/`ERFC`/`ERFC.PRECISE`, `DELTA`, `GESTEP` and the bitwise set (`BITAND`, `BITOR`, `BITXOR`, `BITLSHIFT`, `BITRSHIFT`).

  A complex number in Excel is text, so the `IM*` functions parse `"3+4i"` and write their result back the same way — echoing whichever of `i` or `j` the input used, and refusing to mix the two. `CONVERT` unit names are case sensitive, as Excel's are: `Pica` is a point and `pica` is six to the inch. Prefixes are accepted on metric units, binary prefixes on `bit` and `byte` only, and on no temperature unit — a scale with an offset has no meaningful "milli". ([#253](https://github.com/XLibur/XLibur/pull/253) by [@jafin](https://github.com/jafin))

- **24 more financial functions**: depreciation (`SLN`, `SYD`, `DB`, `DDB`, `VDB`), rate conversion and growth (`EFFECT`, `NOMINAL`, `RRI`, `PDURATION`), fractional dollar notation (`DOLLARDE`, `DOLLARFR`), loan schedules (`ISPMT`, `CUMIPMT`, `CUMPRINC`), discount securities (`TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`, `DISC`, `INTRATE`, `RECEIVED`) and irregular cash flows (`FVSCHEDULE`, `MIRR`, `XNPV`, `XIRR`). `XIRR` solves with Newton–Raphson and falls back to bisection when that wanders off, so a poor guess still converges.

  Two deliberate limitations. The day-count-basis bond family (`PRICE`, `YIELD`, `DURATION`, `MDURATION`, `ACCRINT`, `COUP*`, `ODD*`, `AMOR*`) is not included — it needs a coupon-period engine rather than another day-count fraction. And `DISC`/`INTRATE`/`RECEIVED` take their year fraction from the same code as `YEARFRAC`, so basis 1 (actual/actual) uses `YEARFRAC`'s average-year-length rule. ([#252](https://github.com/XLibur/XLibur/pull/252) by [@jafin](https://github.com/jafin))

#### Security and encryption

- **Password-protected workbooks (ECMA-376 encryption)**: `LoadOptions.Password` opens an encrypted workbook and `SaveOptions.Password` writes one. Reading covers both schemes in the wild — agile encryption (Office 2010 and later) and standard encryption (Office 2007); writing uses agile encryption with the parameters Excel itself writes (AES-256-CBC, SHA-512, 100,000 spins and a fresh random key per save). Previously an encrypted file could not be opened at all and failed opaquely.

  A wrong or missing password throws `XLInvalidPasswordException`, which is what a caller re-prompts on; an altered package whose HMAC fails throws `XLEncryptionException`, because there the password was right. A legacy `.xls` is detected and named rather than mistaken for an encrypted workbook. ([#245](https://github.com/XLibur/XLibur/pull/245) by [@jafin](https://github.com/jafin))

  `Save` and `SaveAs` read a missing `SaveOptions.Password` differently, because they are asking for different things. **`Save` preserves the encryption of the file it came from** — a workbook opened with a password is written back to that file encrypted under the same password, so a load/edit/save round trip works in one call. Giving `Save` a password rotates it in place, or encrypts an origin that was plain; `Save` cannot remove encryption at all. **`SaveAs` states the encryption of the file it writes** — a password means encrypt with it, no password means plaintext, whatever the workbook was loaded as, so a plain `SaveAs` can never silently produce a file the caller cannot open, and is how encryption is removed.

  ```csharp
  using var workbook = new XLWorkbook("Confidential.xlsx", new LoadOptions { Password = "s3cret" });
  workbook.Worksheet("Data").Cell("A1").Value = "Updated";

  workbook.Save();                                          // still encrypted, same password
  workbook.Save(new SaveOptions { Password = "n3w" });      // rotated in place
  workbook.SaveAs("Public.xlsx");                           // plain — explicitly asked for
  ```

  The load password is therefore retained for the lifetime of the workbook, which is what lets `Save` re-encrypt; derived keys are still zeroed after use. ([#251](https://github.com/XLibur/XLibur/issues/251) by [@jafin](https://github.com/jafin))

#### Charts

- **Chart series formatting**: `IXLChartSeries` gained `FillColor`, `LineColor`, `LineWidthPt`, `MarkerStyle` (new `XLMarkerStyle` enum), `MarkerSize`, `MarkerFillColor` and `Smooth`, so a generated chart can be styled instead of relying on Excel's automatic theme colors. Leaving a property `null` omits its element, which keeps the automatic color — nothing is ever written as an explicit black. ([#220](https://github.com/XLibur/XLibur/pull/220) by [@jafin](https://github.com/jafin))

- **Secondary value axis per series**: `IXLChartSeries.UseSecondaryAxis` plots a series against a value axis on the right, so a percentage can share a chart with values in the thousands. It applies to series of the primary chart type as well as to a combo chart's `SecondarySeries`. ([#220](https://github.com/XLibur/XLibur/pull/220) by [@jafin](https://github.com/jafin))

- **Chart data labels**: `IXLDataLabels` on both `IXLChart.DataLabels` (chart-wide) and `IXLChartSeries.DataLabels` (per series, overriding the chart's), with `ShowValue`, `ShowCategoryName`, `ShowSeriesName`, `ShowPercentage`, `NumberFormat` and `Position`. `Position` is validated against the chart type — Excel refuses to open a file that uses a position it does not offer for that type, so the setter throws with the allowed values listed rather than producing a workbook Excel has to repair. ([#221](https://github.com/XLibur/XLibur/pull/221) by [@jafin](https://github.com/jafin))

- **Chart legend**: `IXLChart.Legend` with `Visible`, `Position` (right, bottom, left, top, top-right) and `Overlay`. Charts XLibur creates still have no legend unless one is asked for; setting `Visible = false` on a chart read from a file removes the legend it came with. ([#222](https://github.com/XLibur/XLibur/pull/222) by [@jafin](https://github.com/jafin))

- **Chart axes**: `IXLChart.CategoryAxis`, `ValueAxis` and `SecondaryValueAxis`, each with `Title`, `NumberFormat`, `Min`, `Max`, `MajorUnit`, `MinorUnit`, `Visible`, `MajorGridlines`, `Orientation` (reversed axes) and `LogScale`/`LogBase`. The unit and log-scale properties belong to a value axis in the file format and are skipped on a category axis — except on scatter and bubble charts, whose horizontal axis holds numbers. ([#222](https://github.com/XLibur/XLibur/pull/222) by [@jafin](https://github.com/jafin))

- **Charts loaded from a file can be restyled**: setting the series formatting, data labels, legend or axes on a loaded chart now writes back on save. Only the properties actually assigned are patched into the existing chart part, so trendlines, error bars, gradient fills, per-point colors and label overrides, label and axis fonts, tick marks and the chart's style/color parts are all preserved — and a chart nobody edited is left byte-identical. ([#220](https://github.com/XLibur/XLibur/pull/220), [#221](https://github.com/XLibur/XLibur/pull/221), [#222](https://github.com/XLibur/XLibur/pull/222) by [@jafin](https://github.com/jafin))

- **Chart anchoring**: `IXLChart.Anchor` (`MoveAndSizeWithCells`, `MoveWithCells`, `Absolute`) with `Width`, `Height`, `Left` and `Top` in pixels, so a chart can keep its size as rows are inserted or be pinned to a spot on the sheet. Two-cell anchoring via `Position`/`SecondPosition` remains the default. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

#### Comments

- **Threaded comments are modeled and round-trip**: `IXLCell.GetThreadedComment()`, `CreateThreadedComment(author, text)` and `HasThreadedComment`, with `IXLThreadedComment` (`Text`, `Author`, `CreatedUtc`, `Parent`, `Replies`, `Resolved`, `AddReply`, `Delete`) and the workbook's author list on `IXLWorkbook.Persons`. Previously a thread was read lossily: the whole conversation was flattened into the legacy note's text joined by newlines, discarding authors, timestamps and reply structure. There was no write path at all, so saving regenerated the fallback note without the `tc={rootId}` marker that ties it to the thread. That left Excel with threaded parts it no longer recognized a fallback for.

  ```csharp
  var author = workbook.Persons.Add("Dana Reed");
  var thread = sheet.Cell("B4").CreateThreadedComment(author, "Is this figure final?");
  thread.AddReply(workbook.Persons.Add("Sam Ali"), "Yes — signed off Friday.");
  thread.Resolved = true;
  ```

  A cell carries either a note or a thread, never both; creating one over the other throws rather than silently discarding it. Timestamps are pinned to UTC, since Excel writes no designator. The fallback note and its VML shape are regenerated from the thread on every save, so the two cannot drift after an edit. Mentions round-trip as raw XML, and are dropped when the text changes, because their offsets index into the text they were written against. Verified against Excel 365: threads created from scratch, an Excel-authored file round-tripped, that file edited through the API, and a sheet mixing a note with a thread all open with no repair prompt. ([#258](https://github.com/XLibur/XLibur/pull/258) by [@jafin](https://github.com/jafin))

#### Colors and styles

- **`XLColorType.Automatic`**, with `XLColor.Automatic` and `XLColor.IsAutomatic`. OOXML has four kinds of color and XLibur modeled three; the automatic color (ECMA-376 `CT_Color/@auto`, what Excel's font color picker labels "Automatic") was disguised as a fully transparent black. `auto="1"` is now read and written explicitly. ([#232](https://github.com/XLibur/XLibur/pull/232) by [@jafin](https://github.com/jafin))

### ⚡ Performance

#### Workbook creation

- **40% fewer allocations building a workbook** (create phase of the 50K × 10 benchmark, 305.5 → 183.6 MB; whole benchmark 423.4 → 301.3 MB). Every `cell.Value = x` ran a merged-range membership test that allocated ~250 bytes even on a sheet with no merged ranges at all, boxing the address and building a LINQ iterator chain over an empty list. That test is now skipped outright when the sheet has no merges, and made allocation-free when it does — a merged title row is a common layout and used to cost *more* than no merge at all, because the predicate then actually ran. A cell write drops 375.5 → 103.6 bytes on both paths. ([#185](https://github.com/XLibur/XLibur/pull/185) by [@jafin](https://github.com/jafin))

- **Repeated access to the same cell is cheaper**: `ws.Cell(r, c)` minted a fresh `XLCell` every call, so the per-object style cache never hit and a `.Style` access threw away an 80-byte `XLStyle` one statement later. A small direct-mapped cache now hands back the same wrapper for the same address. ([#185](https://github.com/XLibur/XLibur/pull/185) by [@jafin](https://github.com/jafin))

#### Formula engine

- **The worksheet function table is built once per process instead of once per calc engine** — ~70 KB and 0.12 ms back on every engine constructed, which is one per workbook that touches a formula plus a fresh one on some internal evaluation paths. Opening a workbook allocates 644 → 574 KB. The ~400-entry table is the same whatever the culture (a function takes its culture from the calculation context, not from the engine holding it) and nothing mutates it after it is built, so the per-engine copy only ever bought a private duplicate of a constant. ([#276](https://github.com/XLibur/XLibur/issues/276), [#306](https://github.com/XLibur/XLibur/pull/306) by [@jafin](https://github.com/jafin))

- **Evaluating a formula reuses the addresses it already holds, instead of re-parsing them.** A reference node built its range address by parsing an address string that the node itself had generated from an already-parsed reference. The parse only recovered what the node was already holding, and it cost a `Contains`, a `Split` allocating an array plus a substring per endpoint, and an address parse per endpoint. The address is now built from the parsed form and the result memoized on the node. Memoizing helps here because a formula's syntax tree is shared by every cell holding the same text. Recalculating 20,000 formula rows is 12–18% quicker and allocates 35% less. ([#286](https://github.com/XLibur/XLibur/pull/286) by [@jafin](https://github.com/jafin))

#### Structural edits

- **Inserting or deleting rows one at a time is roughly 3× quicker on a sheet full of formulas** (1,000 single-row inserts over 1,000 formula cells and 1,000 live ranges: 4,753 → 1,556 ms, 6,091 → 2,072 MB). Formula shifting was 68% of that cost, and its expense was not the regular expression it looked like: for every matched address it materialized a live range through the worksheet's range repository, once per reference, per formula, per shift. References are now shifted through `ClosedXML.Parser`, which hands each one back already decomposed, so no address is re-parsed and a formula the shift cannot reach is returned as the same string instance. Equivalence with the previous implementation is pinned by a 2,072-case corpus. Separately, a shift pass no longer visits ranges it provably cannot move — a range out of a shift's reach is now free rather than merely cheap. ([#264](https://github.com/XLibur/XLibur/pull/264) by [@jafin](https://github.com/jafin))

#### Styling

- **86% fewer allocations styling a range, row or column in bulk** (~234 → ~33 bytes per cell): setting a style on a container collected every child cell into a `HashSet` of wrappers before writing anything. Containers that can enumerate exactly those cells now walk the addresses and write the style slice directly. ([#185](https://github.com/XLibur/XLibur/pull/185) by [@jafin](https://github.com/jafin))

#### Copying and clearing ranges

- **`IXLRange.CopyTo` no longer gets slower as the sheet fills up** ([#271](https://github.com/XLibur/XLibur/issues/271), [#303](https://github.com/XLibur/XLibur/pull/303) by [@jafin](https://github.com/jafin)): copying a range in a loop was quadratic in the number of copies, so the natural way to expand a template — one `CopyTo` per generated row — degraded badly. On a 30,000-row sheet with no data validations at all, a 1×10 `CopyTo` cost ~420 µs against ~160 µs on the same sheet when nearly empty; it is now ~13 µs and flat. Three things were at fault, all fixed:

  - `Clear(XLClearOptions.All)`, which `CopyTo` runs over its target first, created a data validation covering the target and immediately deleted it *on every call*, even with no validations on the sheet. Going through `Add`/`Delete` is what splits an overlapping rule so only the cleared cells lose their validation, so it is still done when a rule actually overlaps — and skipped when none does. `XLCell.Clear` already guarded this way.
  - A range index promoted itself from a flat list to a QuadTree after 20 `Add` calls rather than 20 live entries, so a collection that only ever held one range at a time still built a tree.
  - QuadTree quadrants are created on demand and never torn down, so an index that had seen many add/remove cycles kept a skeleton of empty quadrants that every subsequent removal walked. Each quadrant now tracks how many ranges its subtree holds, and traversals skip subtrees holding none.

#### Saving

- **50% fewer allocations on save** (237.1 → 117.9 MB for the same benchmark), with wall time improving from 1187–1290 ms to 1051–1167 ms. Four save helpers materialized a cell wrapper for every used cell just to read one property and now read the underlying storage directly; `int`/`uint` cell values are formatted into a span instead of allocating a string each; style-key hashes are memoized, cutting `StyleKey.GetHashCode` by 95% over 100K styles. Saved output is byte-identical to before. ([#179](https://github.com/XLibur/XLibur/pull/179) by [@jafin](https://github.com/jafin))

### 🐛 Bug Fixes

#### Formulas and references

- **`SUBTOTAL` no longer counts a nested post-2007 function twice.** The check that stops a subtotal counting a subtotal inside its own range compared the function name as the formula stores it, and Excel stores every function added after 2007 under an `_xlfn.` namespace — so a nested `AGGREGATE` was never recognized. The namespace is now stripped before the comparison. ([#255](https://github.com/XLibur/XLibur/pull/255) by [@jafin](https://github.com/jafin))

- **A reference whose rows or columns are all deleted becomes `#REF!`** ([ClosedXML #880](https://github.com/ClosedXML/ClosedXML/issues/880)): endpoints were shifted by the deleted height and clamped to row 1, so deleting rows 1–5 turned `Sheet1!$A$1:$B$2` into `Sheet1!$A$1:$B$1` — a phantom one-row range over whatever data had moved up into it. The same shifter serves cell formulas, so `=SUM(A1:A2)` with those rows deleted now reads `SUM(#REF!)` rather than quietly summing the wrong cells. Deleting the sheet afterwards drops the stale sheet prefix instead of leaving a defined name pointing at a sheet that no longer exists. ([#243](https://github.com/XLibur/XLibur/pull/243) by [@jafin](https://github.com/jafin))

- **Row-only and column-only references are positioned correctly when rows or columns are inserted or deleted**: `3:5` with row 4 deleted became `2:4` — a reference that had walked onto row 2, which it never covered, while losing row 5, which survived. Because a multi-row delete is applied one row at a time, the reference kept drifting up instead of shrinking and never became small enough for the `#REF!` check to fire. Insertions had the mirror problem: inserting two rows at row 4 moved `3:5` to `5:7` rather than expanding it to `3:7`. Both axes now follow the same boundary rules as an equivalent cell range. ([#244](https://github.com/XLibur/XLibur/pull/244) by [@jafin](https://github.com/jafin))

- **A formula reading a table keeps up with the table's contents.** A structured reference such as `=SUM(Table1[Amount])` registered no precedents at all, so nothing invalidated the formula when those cells changed and it went on serving a stale cached value until something forced a full recalculation. Evaluation and dependency tracking now resolve a structured reference through the same code, so they cannot disagree about what one covers. A reference naming a table on another sheet is also resolved against *that* sheet rather than the calling one — it previously read the same coordinates on the wrong sheet and returned 0. ([#297](https://github.com/XLibur/XLibur/pull/297) by [@jafin](https://github.com/jafin))

- **A defined name holding a structured reference resolves to what the reference says.** `IXLDefinedName.Ranges` resolved every form as though it named a data column, so a name over `Sales[[#Headers],[Amount]]` pointed at the data instead of the header, a column span `Sales[[Amount]:[Tax]]` lost everything but its first column, and a whole-table `Sales[#All]` resolved to nothing. An unknown column threw `ArgumentOutOfRangeException` out of a property getter — reachable just by loading a workbook whose table column had since been renamed; it now contributes no range, as an unknown table already did. `Sales[Amount]`, the common form, resolves as before. ([#311](https://github.com/XLibur/XLibur/pull/311) by [@jafin](https://github.com/jafin))

- **A deletion that removes the tail of a reference no longer inverts it.** The new bottom was the last row plus the shift, with nothing clamping it to the row above the deletion. So `3:5` with rows 5–7 deleted came back as `3:2` — an inverted range, not a valid formula. And `A2:A8` with rows 5–9 deleted came back as `A2:A3`, dropping row 4, which survived. Both now shrink to the rows that are left. ([#264](https://github.com/XLibur/XLibur/pull/264) by [@jafin](https://github.com/jafin))

- **A reference written back to front evaluates instead of throwing.** `=SUM(B2:A1)` threw `ArgumentException("Range address must be normalized")` because the parser hands endpoints back in the order they were written and nothing ordered them before use. Each axis is now ordered independently, carrying its own fixed marker, and the formula evaluates as Excel's does. ([#286](https://github.com/XLibur/XLibur/pull/286) by [@jafin](https://github.com/jafin))

- **Copying a worksheet repoints the copy's self-references at the copy** ([ClosedXML #836](https://github.com/ClosedXML/ClosedXML/issues/836)): copying a sheet named `Original` holding `Original!A1 * 3` produced a sheet whose formula still pointed back at the original. References to *other* sheets are left alone, as they should be. ([#241](https://github.com/XLibur/XLibur/pull/241) by [@jafin](https://github.com/jafin))

#### Text and number parsing

- **Text-to-number and text-to-date coercion matches Excel much more closely.** The date-time patterns had no seconds component, so `8/22/2008 3:30:45 PM` failed to coerce at all. Beyond that:

  - Time components that overflow their range carry into the date. `11/30/2022 24:59` is one minute to one in the morning of December 1st.
  - Parenthesized, spaced and sign-separated numbers such as `(100%)`, `(   1,000.54  )` and `- 100 %` are read as negative.
  - A month is matched by any prefix from three letters up, as Excel does.
  - A shortened or dot-suffixed AM/PM designator is accepted.

  In the other direction, group-separator and currency placement are now enforced the way Excel enforces them. `1,00`, `1,00,000` and `1$` are rejected under en-US, where the underlying BCL parse accepted them. ([#241](https://github.com/XLibur/XLibur/pull/241) by [@jafin](https://github.com/jafin))

#### Rich text and shared strings

- **Rich-text runs keep their automatic or absent color through a round-trip**: a plain load → `SaveAs` wrote every color-less run back with an explicit `<color rgb="FF000000"/>`, and promoted a plain string carrying a phonetic guide (`<t>` + `<rPh>`, common in Japanese workbooks) into a synthetic run that inherited the same injection. An explicit black cannot be overridden by a theme color change or by conditional formatting, where the automatic color it replaced can. A run that was read with no `<rPr>` is now written back without one, so inheritance stays inheritance; splitting such a run keeps that property rather than materializing the cell font as explicit formatting. ([#225](https://github.com/XLibur/XLibur/pull/225), [#227](https://github.com/XLibur/XLibur/pull/227) by [@jafin](https://github.com/jafin))

- **A plain shared string that carries a phonetic guide saves cleanly.** Its text is decoded on the way in, so a decoded `_xHHHH_` escape is a literal control character — not valid XML content — and the writer emitted it raw, failing with an `ArgumentException` naming the invalid character. ([#225](https://github.com/XLibur/XLibur/pull/225), [#227](https://github.com/XLibur/XLibur/pull/227) by [@jafin](https://github.com/jafin))

#### Colors and conditional formatting

- **An automatic color is no longer written as an explicit `rgb="00000000"`.** Color writers switch on the color type, and the automatic color fell into the RGB arm — pinning down a color the source deliberately left for the application to resolve. The three conditional-format color converters gained explicit automatic arms too, where they would otherwise have dropped the color silently. ([#232](https://github.com/XLibur/XLibur/pull/232) by [@jafin](https://github.com/jafin))

#### Charts

- **Charts anchored with a one-cell or absolute anchor load into `IXLWorksheet.Charts`.** The reader only looked at `xdr:twoCellAnchor`, so a chart Excel had anchored either of the other two ways was missing from `IXLWorksheet.Charts` entirely (its XML survived a round trip, but the chart was invisible to the API). ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

- **3D and of-pie chart groups are read.** `c:pie3DChart`, `c:line3DChart`, `c:area3DChart`, `c:surface3DChart` and `c:ofPieChart` were not recognized, so an Excel-authored chart using one loaded with no series and the wrong chart type. Their series and series formatting now read the same as the 2D groups', and pie-of-pie and bar-of-pie are told apart. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

- **Chart XML now passes OpenXML schema validation.** Three long-standing violations in the chart writer are fixed: series names were written as a `c:strRef` with no required `c:f` (a literal name now uses `<c:tx><c:v>`, and both forms are read back), `c:doughnutChart` omitted the required `c:holeSize`, and `c:marker` was written after `c:cat`/`c:val` instead of before. Excel tolerated all three, but stricter readers and `SaveOptions.ValidatePackage` did not. ([#220](https://github.com/XLibur/XLibur/pull/220) by [@jafin](https://github.com/jafin))

- **A `Line` chart whose markers are switched off reads back as `Line`.** The reader treated the presence of a `c:marker` element as "has markers", even when it held `<c:symbol val="none"/>`. ([#220](https://github.com/XLibur/XLibur/pull/220) by [@jafin](https://github.com/jafin))

- **Charts with more than one plot group of the same type now read all of their series.** The reader took only the first `c:barChart` (or `c:lineChart`, …) of a plot area, so the series of a second group — which is how Excel stores a secondary axis — were dropped. ([#220](https://github.com/XLibur/XLibur/pull/220) by [@jafin](https://github.com/jafin))

- **Which of those groups is on the secondary axis no longer depends on the order they appear in the file.** The primary axis pair was taken from whichever group came first. So a file that wrote its secondary group ahead of the primary one read back with `UseSecondaryAxis` inverted on every series, and the two axis models swapped. The reader now passes over the group whose value axis crosses at the maximum, which is how a secondary axis comes to be drawn on the right. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

- **`Smooth` is honored on a new stock chart.** A stock chart's series are `CT_LineSer` and take `c:smooth`, but the writer never emitted it, so the property worked on a stock chart read from a file and was dropped on one XLibur created. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

- **`IXLChart.SetTitle` on a chart loaded from a file is applied on save** ([#272](https://github.com/XLibur/XLibur/issues/272)): the patcher rewrote the legend, axes, data labels and series of a loaded chart, but never its title. The model accepted the new title and `chart.Title` read it back, yet the saved part kept the old one, with no exception and no warning. The title is now patched in like everything else. Only the text is replaced, so a title that came out of Excel keeps its font, color, layout and overlay. Setting it to `null` removes the title and sets `c:autoTitleDeleted`, so Excel does not put an automatic one back in its place. This covers extended (ChartEx) charts — waterfall, funnel, treemap, sunburst, box & whisker — as well as the standard ones. The title is the only edit a ChartEx chart takes, since those carry none of the formatting XLibur models. ([#283](https://github.com/XLibur/XLibur/pull/283) by [@jafin](https://github.com/jafin))

- **Positioning a legend that is not there no longer creates one.** `IXLChartLegend.Position` and `Overlay` are documented as ignored while `Visible` is `false`, and a new chart gets no legend from them — but assigning one of them on a *loaded* chart that had no legend added one. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

- **Chart-wide data labels reach every group of a loaded combo chart.** `IXLChart.DataLabels` applies to the whole chart, and a new combo chart gets them on both of its plot groups, but a loaded one was patched on the primary group only — so turning labels on left the secondary series unlabelled. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

- **`Series.Add(...)` on a chart loaded from a file throws instead of being discarded on save.** A loaded chart is patched, not regenerated, so a new series had nowhere to be written and vanished without a word. It now throws `NotSupportedException`, as `UseSecondaryAxis` already did. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

#### Comments

- **`IXLComment.Delete()` removes the note from where it is now, not where it was created** ([#259](https://github.com/XLibur/XLibur/issues/259)): a note remembered the cell it was constructed with, and inserting or deleting rows or columns moves the note without telling it. Deleting a note on `A5` after two rows were inserted above cleared `A5` — by then an empty cell — and left the note sitting on `A7`. The note is now confirmed to still be the one at the remembered address before that address is cleared, and located by identity when it is not. ([#268](https://github.com/XLibur/XLibur/pull/268) by [@jafin](https://github.com/jafin))

#### Data validation

- **A data validation is never written covering nothing.** The schema requires a rule's `sqref` to be non-empty, and Excel reads `sqref=""` as corruption — it repairs the workbook and drops *every* validation on the sheet, not just the broken one. Two paths produced one: adding a validation over a range that wholly contained an existing rule left that rule with no coverage, and the public `ClearRanges`/`RemoveRange` let a caller empty one directly. A rule the split leaves with nothing is now deleted, and the writer skips any rule with no coverage — a rule applying to no cell has nothing to say. Where the emptied rule was the only one on the sheet, the `dataValidations` element is omitted rather than written holding a single broken child. ([#290](https://github.com/XLibur/XLibur/pull/290) by [@jafin](https://github.com/jafin))

#### Autofilter

- **A worksheet autofilter keeps the parts of a filter XLibur does not model.** The runtime state is a lossy view of the element — it has no room for an `iconFilter`, the button attributes, `extLst`, or the dynamic filter types beyond the two averages — and those were dropped on save. A column that has not been changed is now written back from the criteria it was loaded with; every mutation drops them, so an edit is never discarded. ([#310](https://github.com/XLibur/XLibur/pull/310) by [@jafin](https://github.com/jafin))

- **A relative-date filter loads and survives the round trip.** A `dynamicFilter` using any of the ~38 relative date types (`thisMonth`, `yearToDate`, `lastQuarter`, …) threw `KeyNotFoundException`, because the token was looked up in a two-entry map. The token is carried as written, so the file loads and the filter survives. ([#310](https://github.com/XLibur/XLibur/pull/310) by [@jafin](https://github.com/jafin))

#### Pivot tables

- **Pivot field filters survive a round trip.** The `filters` element holds the label, value, date and top-N filters applied to pivot fields — what Excel offers from a field's dropdown — and the reader skipped it, so loading and saving un-filtered every pivot table in the workbook. That changes what the workbook *shows*, not just what it remembers. (Not to be confused with the report-filter axis, which is the `pageFields` element and was already supported.) ([#300](https://github.com/XLibur/XLibur/pull/300), [#310](https://github.com/XLibur/XLibur/pull/310) by [@jafin](https://github.com/jafin))

- **A PivotChart keeps its manual series and point formatting.** The chart part holds the formatting, but the pivot table's `chartFormats` collection is what ties each formatting record to the pivot area it applies to, and the reader skipped the element — so a round trip dropped every link and the formatting disappeared. ([#294](https://github.com/XLibur/XLibur/pull/294) by [@jafin](https://github.com/jafin))

### 🗑️ Deprecations

#### Collections

- **`IXLBaseCollection<TSingle, TMultiple>` is deprecated ahead of removal.** Nothing in XLibur implements, extends or consumes it — the collections it looks like it should describe (`IXLColumns`, `IXLRows`, `IXLCells`, `IXLRangeColumns`, `IXLRangeRows`) all derive from `IEnumerable<T>` alone. Marking it produced no build warnings anywhere in the solution, which is a second confirmation that it is an orphan. It is public, so an external consumer could in principle implement it, and it gets a normal deprecation period rather than being removed outright. ([#296](https://github.com/XLibur/XLibur/pull/296) by [@jafin](https://github.com/jafin))

  Note also that the three groups of already-shipped `[Obsolete]` members will be removed in the next minor version — `XLColor.NoColor`, deprecated below, is the exception and stays. ([#296](https://github.com/XLibur/XLibur/pull/296) by [@jafin](https://github.com/jafin))

#### Colors and styles

- **`XLColor.NoColor` is deprecated in favor of `XLColor.Automatic`.** Some Excel pickers (sheet tab, fill background) label the same value "No Color" — that is a GUI convention, not a different value. `NoColor` still compiles but now warns, which is an error for consumers building with `TreatWarningsAsErrors`. ([#232](https://github.com/XLibur/XLibur/pull/232) by [@jafin](https://github.com/jafin))

  ```csharp
  // Before
  cell.Style.Font.FontColor = XLColor.NoColor;

  // After
  cell.Style.Font.FontColor = XLColor.Automatic;
  ```

  A straight rename — `NoColor` returns `Automatic`, so the two are the same value.

## XLibur.Report — Unreleased

`XLibur.Report` is versioned and released independently of the core library, on its own
`report-v*` tag stream, so its changes are recorded here rather than under the core version
above. Nothing in this section has shipped yet.

### ✨ New Features

#### Report templating

- **`XLibur.Report` — build a report from a template that is itself an ordinary `.xlsx` file.** You author the report in Excel — the fonts, the number formats, the borders, the chart, the pivot table — and mark the parts that come from data with `{{ }}` expressions and `<<Tag>>` markers. At run time you bind .NET objects to it and generate the finished workbook, so none of the report's *appearance* is described in code. ([#273](https://github.com/XLibur/XLibur/pull/273) by [@jafin](https://github.com/jafin))

  ```csharp
  using XLibur.Report;

  using var template = new XLTemplate("SalesReport.xlsx");
  template.AddVariable("Company", "Contoso Ltd");
  template.AddVariable("Sales", sales);

  template.Generate();
  template.SaveAs("SalesReport-2026.xlsx");
  ```

  A template does three things. A `{{ … }}` expression anywhere in a cell's text is evaluated and replaced, resolving to a *typed* value — a decimal total reaches Excel as a number, not as text — and `&=` at the start of a cell builds a formula. A **defined name** matching a bound collection makes the rows it covers repeat, once per item, with `item` in scope. The last row of a bound range is its **options row**: it is not repeated, it carries the tags, and it is deleted if nothing was written into it.

  The tags cover totals (`<<Sum>>`, `<<Avg>>`, `<<Count>>`, `<<Max>>`, `<<Min>>`, `<<StdDev>>`, `<<Var>>` and friends, left as live `SUBTOTAL` formulas rather than computed numbers), `<<Sort>>`, `<<Group>>` with an Excel outline and subtotals, `<<If>>`, `<<Horizontal>>` to repeat across columns instead of down rows, `<<AutoFilter>>`, `<<ColsFit>>`/`<<RowsFit>>`, `<<Hidden>>`, `<<Delete>>` and `<<Pivot>>`. You can register a tag of your own.

  A bad expression records a `TemplateError` and marks the cell red rather than aborting the run, so one broken cell does not cost you the report.

- **Charts, pivot tables, pictures and conditional formats survive range expansion.** ([#273](https://github.com/XLibur/XLibur/pull/273) by [@jafin](https://github.com/jafin))

  - A chart series drawn over the template's repeated row plots every generated row.
  - A template pivot table is re-pointed at the generated rows and refreshed. If the rows grew over it, it moves. `<<Pivot dest="…">>` builds a new one from the template's own shape.
  - A picture below a bound range ends up below the generated rows.
  - A conditional-format rule over the repeated rows is *stretched* over the generated block, not copied per generated cell.

- **Excel's worksheet functions are available inside `{{ }}`, under their Excel names** — `{{ SUM(array.map items "Total") }}` — so a template author reaches for the function they already know rather than a Scriban equivalent. ([#273](https://github.com/XLibur/XLibur/pull/273) by [@jafin](https://github.com/jafin))

- **`XLibur.Report.DynamicLinq` runs ClosedXML.Report's C# expression syntax as written.** The default engine is [Scriban](https://github.com/scriban/scriban), which is a template language rather than C#, so expressions like `{{item.Name.ToUpper()}}` need translating. Installing the DynamicLinq package and passing its engine to `XLTemplate` avoids touching them at all. ([#273](https://github.com/XLibur/XLibur/pull/273) by [@jafin](https://github.com/jafin))

  ```csharp
  using var template = new XLTemplate("SalesReport.xlsx", new DynamicLinqExpressionEngine());
  ```

  The rest of a ClosedXML.Report template carries over unchanged: defined names bind ranges, the last row is the options row, `<<Tag>>` markers keep their names and meanings, property paths in a defined name work the same way (`Order_Lines` binds `Order.Lines`), and the API shape — `new XLTemplate(path)`, `AddVariable`, `Generate`, `SaveAs` — is the same.

- **Report generation is culture-controlled, not machine-controlled.** `IExpressionEngine.Culture` is what orders rows and formats group labels, so the same template over the same data produces the same report wherever it runs. Both engines default to the invariant culture and expose the culture they were constructed with. ([#305](https://github.com/XLibur/XLibur/pull/305) by [@jafin](https://github.com/jafin))

- **Hidden sheets are generated.** A hidden sheet is a normal place to keep the lookup tables and working data a report is built from, so generation follows the data rather than what the reader can see. Hiding a sheet controls what the reader sees; deleting it after `Generate()` is what keeps it out of the report. ([#304](https://github.com/XLibur/XLibur/pull/304) by [@jafin](https://github.com/jafin))

- **Defined names bind under Excel's scoping and matching rules.** Two sheet-scoped names sharing a name — a `Q1!Items` section and a `Q2!Items` section — both bind, where previously the second generated blank. A workbook-scoped name binds except on a sheet declaring its own name of that name, which is what Excel does. And a name is matched to its variable case-insensitively, as Excel holds names in one case-insensitive namespace, so a template author who types `ITEMS` in the name box has named the variable added as `Items`. An exact match still wins, and two variables differing only by case are reported rather than bound arbitrarily. ([#307](https://github.com/XLibur/XLibur/pull/307), [#314](https://github.com/XLibur/XLibur/pull/314) by [@jafin](https://github.com/jafin))

### 🔧 Dependencies

- **`XLibur.Report` no longer pins an exact core version.** It built against core internals, and a package compiled against internals can only honestly declare an exact dependency — so it pinned `[0.200.0]`, which made every core release a Report release and defeated the point of the separate `report-v*` tag stream. Report now builds against public API only (see the core section above) and declares an open floor of `0.201.0`, so a Report release travels with any core at or above that. ([#354](https://github.com/XLibur/XLibur/pull/354) by [@jafin](https://github.com/jafin))

- **`XLibur.Report.DynamicLinq` moves to System.Linq.Dynamic.Core 1.7.3** (from 1.7.1). ([#345](https://github.com/XLibur/XLibur/pull/345) by [@dependabot](https://github.com/apps/dependabot))

## v0.106.0 - 2026-07-25

First XLibur release since forking [ClosedXML v0.105.0](https://github.com/ClosedXML/ClosedXML/)
(May 2025). Everything below is relative to that baseline.

### ✨ New Features

- **Charts — all 78 `XLChartType` values**: End-to-end chart creation, saving, loading and round-tripping. Covers bar/column (clustered, stacked, percent, 2D and 3D), the 21 Bar3D cone/cylinder/pyramid shapes, line, area, pie/doughnut (including pie-to-pie and pie-to-bar), radar, scatter/XY, bubble, and surface types, plus data series, chart titles, combo charts and positioning. The previous `IXLChart`/`XLChart` stubs are now backed by a real implementation.

- **Dynamic arrays**: Modern array functions — `SEQUENCE`, `UNIQUE`, `SORT`, `SORTBY`, `FILTER`, `XLOOKUP` and `XMATCH` — together with a **spill engine**. A dynamic-array formula written into a single cell now auto-fills its computed footprint into the neighboring cells, grows and shrinks as the result changes, and round-trips through save/load. Only the anchor cell holds the formula; spilled cells stay formula-less, matching Excel. A footprint blocked by existing content, or one that would run past the sheet edge, collapses to the new `#SPILL!` error (`XLError.SpillRange`) on the anchor.

- **New worksheet functions**:
  - Conditional aggregates: `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`
  - Logical: `IFS`, `SWITCH`
  - Statistical: `SMALL`, `RANK`, `PERCENTILE`, `QUARTILE`, `MODE`
  - Financial: `PV`, `NPV`, `IRR`, `RATE`, `NPER`, `PPMT`
  - Reference: `INDIRECT`

- **Wildcard support in `HLOOKUP` and `VLOOKUP`**: `*` and `?` patterns now match in lookup values, as they do in Excel.

- **Swappable font engine, and a font-library-free core**: Text measurement (column auto-fit, row heights, glyph metrics) moved behind `IXLFontEngine`, and the font library ships as a separate package rather than being compiled into the core assembly. The MIT-licensed SkiaSharp engine is the default and auto-registers the first time a workbook is created, so no startup call is needed. This lets you choose a font library whose license suits you, and stops library authors inheriting a font dependency they don't need. See the Upgrade Guide below.

- **`XLibur.Bundle` meta-package**: Installs the core library together with the default font engine, so a single package reference behaves like ClosedXML out of the box.

- **Editable pictures inside group shapes**: Pictures nested in `xdr:grpSp` groups — at any nesting depth — can be read, resized, moved, added, removed and grouped through a first-class public API. Geometry is computed through the composed group transform, and moves operate in sheet space.

- **DataBar conditional formats can be modified after creation**, including axis settings, rather than being write-once at creation time.

- **Pivot table improvements**: Named ranges resolve as a pivot cache source, and `autoSortScope` on pivot fields round-trips through load/save.

### 🐛 Bug Fixes

- **Array and dynamic-array formulas no longer break on row/column shifts**: Inserting or deleting rows/columns anywhere in a workbook used to rebuild every formula cell through the `FormulaA1` setter, which turned a single array formula (shared across its whole range) into one *normal* formula per cell. For dynamic arrays this split a single spilled formula such as `=UNIQUE(...)` into multiple implicit-intersection `=@UNIQUE(...)` cells, even when the edit happened on an unrelated sheet. Shifts now update the shared formula instance in place — preserving its array/dynamic-array nature — and relocate the spill range for same-sheet inserts/deletes.

- **Deleting through an array no longer corrupts its stored range**: When a delete overlapped an array formula, relocating the array's top edge could drive the coordinate below 1. `Point` does not bounds-check, so the value silently overflowed and corrupted the stored range.

- **Data-validation formulas are shifted with the sheet**: Inserting or deleting rows/columns relocated each rule's ranges (`sqref`) but left cell references *inside* the criteria formulas (`formula1`/`formula2`) pointing at the pre-shift location. Any `List`, `Custom` or comparison rule referencing other cells broke — most visibly dependent dropdown pairs driven by `OFFSET`/`MATCH`. The in-memory value was wrong immediately after the shift, before any save.

- **Data validations no longer vanish when inserting at row 1 or column 1**: The data-validation index was keyed by address at insert time and never re-keyed, so an insert at the first row/column left it stale. At save time the split logic then treated a rule's own out-of-date entry as a competing rule and stripped its ranges, emitting `<dataValidation sqref="">`. Excel rejected the file on open with *"Removed Records: Data validation"*. The index is now reconciled before consolidation.

- **Conditional-format ranges shift once, not twice** ([ClosedXML #2850](https://github.com/ClosedXML/ClosedXML/issues/2850)): Inserting rows or columns below the first line doubled the shift for any rule whose shifted target address collided with another rule's existing range. A rule at `K13` that should move to `K23` landed at `K33`, while rules whose targets happened to be empty shifted correctly.

- **Page breaks no longer inflate the used range** ([ClosedXML #2842](https://github.com/ClosedXML/ClosedXML/issues/2842)): `AddHorizontalPageBreak()`/`AddVerticalPageBreak()` wrote `brk@max` as the sheet's full row/column count. Excel read that as a huge used range, so a file with ~2000 rows of data rendered with a scrollbar spanning all 1,048,576.

- **Named ranges shrink correctly when their first row or column is deleted**: Deleting the first row of a named range shifted both endpoints up instead of removing the deleted row and shifting the survivors, so `A3:A4` became `A2:A3` — expanding the range to include a row that was never part of it. Excel produces `A3:A3`.

- **Totals-row formulas escape column names containing spaces**: Structured references for headers such as `Feb 2023` used the single-bracket form, producing a formula Excel could not parse.

- **Grouped pictures and shapes survive a load/save round-trip** instead of being dropped.

- **Cached formula values are preserved on save**: Cached values are now written whenever they exist and the formula has not been dirtied, regardless of `EvaluateFormulasBeforeSaving`, and the data-type attribute is preserved. This fixes round-trip loss of dynamic-array results (`SORT`, `UNIQUE`, `FILTER`) and spill cell values.

- **Pivot table alignment formatting round-trips**: Alignment in pivot table differential formats (DXF) was lost on load/save.

### ⚡ Performance

- **61% fewer allocations and 16.5% less wall time on load** (250K × 15 benchmark), from removing per-cell and per-entry garbage in the shared-string reader, cell value/attribute reads, and a new style cache.

- **`<sheetData>` is read with a raw `XmlReader`**: Worksheet loading — the dominant cost when opening a workbook — no longer goes through the OpenXML SDK's `OpenXmlPartReader`, which rebuilt a `ReadOnlyCollection<OpenXmlAttribute>` and materialized text through its object model for every `<c>`, `<row>` and `<f>` element. Measured in isolation on a 250K × 15 sheet (3.75M cells), that reader accounted for ~67% of load time and ~80% of load allocations — roughly 4× slower and 5× more garbage than an equivalent raw `XmlReader` traversal.

- **Faster string cell reads**: `GetValue<string>()`/`GetString()` — the most common cell read — no longer runs a compiled regex over the whole string (allocating a `MatchCollection`) to find the rare `_xHHHH_` escape sequence.

- **Reduced allocations in 10 per-cell, per-formula and per-address hot-path methods**, with no public API or behavior change.

- **Load and save hot paths**: The shared-string reader is pre-allocated from the SST count, merged cells stream instead of building a full DOM, worksheet attributes are parsed in a single pass, calc-engine overhead is skipped for formula cells during load, and `uint` boxing was removed from the XML writer.

- **`XmlEncoder.EncodeString` fast-path**: Added a character scan that short-circuits before the `Regex` and `StringBuilder` when a string contains no characters that need encoding (the common case for plain text). For workbooks with ~50K unique shared strings this eliminates ~50K `StringBuilder` allocations, ~50K regex evaluations, and ~50K string copies on save.

- **`IXLWorksheet.SetCellValue(int row, int column, XLCellValue value)`** (new API): Sets a cell value directly on the worksheet's internal storage without allocating an intermediate `XLCell` object. For bulk data population (e.g. 50K × 3) this eliminates ~150K object allocations that the `Cell(row, col).SetValue(...)` pattern would create.

### 📖 Upgrade Guide

#### Migrating from ClosedXML

The public API surface is largely unchanged from ClosedXML 0.105. To migrate:

1. Install `XLibur.Bundle` from NuGet.
2. Replace `using ClosedXML` namespace references with `using XLibur`.

Namespaces are prefixed with `XLibur` so both libraries can be referenced in the same project.

#### Font engine packaging

This is the one area where XLibur's packaging differs from ClosedXML. ClosedXML compiles
[SixLabors.Fonts](https://github.com/SixLabors/Fonts) into its core assembly; XLibur keeps the core
assembly free of any font library and ships the engine as a separate, swappable package.

- **Installing `XLibur.Bundle` (or `XLibur` + `XLibur.Fonts.SkiaSharp`) requires no code changes.**
  The default SkiaSharp engine auto-registers on first workbook creation. It resolves system fonts
  and falls back to an embedded, metric-only Calibri-compatible font, so text measurement works in
  headless and serverless environments with no system fonts installed.
- **Installing the bare `XLibur` package with no font engine** throws an `InvalidOperationException`
  when a workbook is created, telling you to add a font engine package. This is intentional — it is
  how the core stays font-library-agnostic.
- **To keep ClosedXML 0.105's exact engine**, install `XLibur.Fonts.SixLabors.V1` and call
  `SixLaborsV1FontBootstrap.Register()` at startup.

See [docs/font-architecture.md](docs/font-architecture.md) for the full design and the list of
available engines.

#### Using `SetCellValue` for bulk writes

The existing `Cell(row, col).SetValue(value)` API continues to work and remains the correct choice when you need full cell semantics (formula clearing, merged-range checks, table header refresh). No code changes are required.

For **performance-critical bulk data population** where you are writing values into empty or freshly-created cells, you can switch to the new direct API:

```csharp
// Before (allocates an XLCell per call):
for (int row = 1; row <= 50_000; row++)
{
    ws.Cell(row, 1).SetValue(row);
    ws.Cell(row, 2).SetValue($"Item {row}");
    ws.Cell(row, 3).SetValue(row * 1.5);
}

// After (zero intermediate allocations):
for (int row = 1; row <= 50_000; row++)
{
    ws.SetCellValue(row, 1, row);
    ws.SetCellValue(row, 2, $"Item {row}");
    ws.SetCellValue(row, 3, row * 1.5);
}
```

`SetCellValue` handles date/time number format application and quote-prefix stripping, so the resulting cell content and formatting is identical for data values. The following behaviors are **not** performed by `SetCellValue` — use `Cell().SetValue()` if you need them:

| Behavior | `Cell().SetValue()` | `SetCellValue()` |
|---|---|---|
| Set value and number format | Yes | Yes |
| Clear existing formula | Yes | No |
| Check merged range (inferior cell skip) | Yes | No |
| Refresh table header fields | Yes | No |
