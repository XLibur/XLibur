# Changelog

## Unreleased

### Added

- **Chart series formatting**: `IXLChartSeries` gained `FillColor`, `LineColor`, `LineWidthPt`, `MarkerStyle` (new `XLMarkerStyle` enum), `MarkerSize`, `MarkerFillColor` and `Smooth`, so a generated chart can be styled instead of relying on Excel's automatic theme colours. Leaving a property `null` omits its element, which keeps the automatic colour — nothing is ever written as an explicit black.

- **Secondary value axis per series**: `IXLChartSeries.UseSecondaryAxis` plots a series against a value axis on the right, so a percentage can share a chart with values in the thousands. It applies to series of the primary chart type as well as to a combo chart's `SecondarySeries`.

- **Chart data labels**: `IXLDataLabels` on both `IXLChart.DataLabels` (chart-wide) and `IXLChartSeries.DataLabels` (per series, overriding the chart's), with `ShowValue`, `ShowCategoryName`, `ShowSeriesName`, `ShowPercentage`, `NumberFormat` and `Position`. `Position` is validated against the chart type — Excel refuses to open a file that uses a position it does not offer for that type, so the setter throws with the allowed values listed rather than producing a workbook Excel has to repair.

- **Chart legend**: `IXLChart.Legend` with `Visible`, `Position` (right, bottom, left, top, top-right) and `Overlay`. Charts XLibur creates still have no legend unless one is asked for; setting `Visible = false` on a chart read from a file removes the legend it came with.

- **Chart axes**: `IXLChart.CategoryAxis`, `ValueAxis` and `SecondaryValueAxis`, each with `Title`, `NumberFormat`, `Min`, `Max`, `MajorUnit`, `MinorUnit`, `Visible`, `MajorGridlines`, `Orientation` (reversed axes) and `LogScale`/`LogBase`. The unit and log-scale properties belong to a value axis in the file format and are skipped on a category axis — except on scatter and bubble charts, whose horizontal axis holds numbers.

- **Charts loaded from a file can be restyled**: setting the series formatting, data labels, legend or axes on a loaded chart now writes back on save. Only the properties actually assigned are patched into the existing chart part, so trendlines, error bars, gradient fills, per-point colours and label overrides, label and axis fonts, tick marks and the chart's style/colour parts are all preserved — and a chart nobody edited is left byte for byte as it was.

- **Chart anchoring**: `IXLChart.Anchor` (`MoveAndSizeWithCells`, `MoveWithCells`, `Absolute`) with `Width`, `Height`, `Left` and `Top` in pixels, so a chart can keep its size as rows are inserted or be pinned to a spot on the sheet. Two-cell anchoring via `Position`/`SecondPosition` remains the default.

### Fixed

- **Charts anchored with a one-cell or absolute anchor are no longer dropped on load.** The reader only looked at `xdr:twoCellAnchor`, so a chart Excel had anchored either of the other two ways was missing from `IXLWorksheet.Charts` entirely (its XML survived a round trip, but the chart was invisible to the API).

- **3D and of-pie chart groups are read.** `c:pie3DChart`, `c:line3DChart`, `c:area3DChart`, `c:surface3DChart` and `c:ofPieChart` were not recognised, so an Excel-authored chart using one loaded with no series and the wrong chart type. Their series and series formatting now read the same as the 2D groups', and pie-of-pie and bar-of-pie are told apart.

- **Chart XML now passes OpenXML schema validation.** Three long-standing violations in the chart writer are fixed: series names were written as a `c:strRef` with no required `c:f` (a literal name now uses `<c:tx><c:v>`, and both forms are read back), `c:doughnutChart` omitted the required `c:holeSize`, and `c:marker` was written after `c:cat`/`c:val` instead of before. Excel tolerated all three, but stricter readers and `SaveOptions.ValidatePackage` did not.

- **A `Line` chart whose markers are switched off no longer reads back as `LineWithMarkers`.** The reader treated the presence of a `c:marker` element as "has markers", even when it held `<c:symbol val="none"/>`.

- **Charts with more than one plot group of the same type now read all of their series.** The reader took only the first `c:barChart` (or `c:lineChart`, …) of a plot area, so the series of a second group — which is how Excel stores a secondary axis — were dropped.

- **Which of those groups is on the secondary axis no longer depends on the order they appear in the file.** The primary axis pair was taken from whichever group came first, so a file that wrote its secondary group ahead of the primary one read back with `UseSecondaryAxis` inverted on every series and the two axis models swapped. The group whose value axis crosses at the maximum — how a secondary axis comes to be drawn on the right — is now passed over instead.

- **`Smooth` is honoured on a new stock chart.** A stock chart's series are `CT_LineSer` and take `c:smooth`, but the writer never emitted it, so the property worked on a stock chart read from a file and was silently dropped on one XLibur created.

- **Positioning a legend that is not there no longer creates one.** `IXLChartLegend.Position` and `Overlay` are documented as ignored while `Visible` is `false`, and a new chart gets no legend from them — but assigning one of them on a *loaded* chart that had no legend added one.

- **Chart-wide data labels reach every group of a loaded combo chart.** `IXLChart.DataLabels` applies to the whole chart, and a new combo chart gets them on both of its plot groups, but a loaded one was patched on the primary group only — so turning labels on left the secondary series unlabelled.

- **`Series.Add(...)` on a chart loaded from a file throws instead of being discarded on save.** A loaded chart is patched, not regenerated, so a new series had nowhere to be written and vanished without a word. It now throws `NotSupportedException`, as `UseSecondaryAxis` already did.

## v0.106.0 - 2026-07-25

First XLibur release since forking [ClosedXML v0.105.0](https://github.com/ClosedXML/ClosedXML/)
(May 2025). Everything below is relative to that baseline.

### Added

- **Charts — all 78 `XLChartType` values**: End-to-end chart creation, saving, loading and round-tripping. Covers bar/column (clustered, stacked, percent, 2D and 3D), the 21 Bar3D cone/cylinder/pyramid shapes, line, area, pie/doughnut (including pie-to-pie and pie-to-bar), radar, scatter/XY, bubble, and surface types, plus data series, chart titles, combo charts and positioning. The previous `IXLChart`/`XLChart` stubs are now backed by a real implementation.

- **Dynamic arrays**: Modern array functions — `SEQUENCE`, `UNIQUE`, `SORT`, `SORTBY`, `FILTER`, `XLOOKUP` and `XMATCH` — together with a **spill engine**. A dynamic-array formula written into a single cell now auto-fills its computed footprint into the neighbouring cells, grows and shrinks as the result changes, and round-trips through save/load. Only the anchor cell holds the formula; spilled cells stay formula-less, matching Excel. A footprint blocked by existing content, or one that would run past the sheet edge, collapses to the new `#SPILL!` error (`XLError.SpillRange`) on the anchor.

- **New worksheet functions**:
  - Conditional aggregates: `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`
  - Logical: `IFS`, `SWITCH`
  - Statistical: `SMALL`, `RANK`, `PERCENTILE`, `QUARTILE`, `MODE`
  - Financial: `PV`, `NPV`, `IRR`, `RATE`, `NPER`, `PPMT`
  - Reference: `INDIRECT`

- **Wildcard support in `HLOOKUP` and `VLOOKUP`**: `*` and `?` patterns now match in lookup values, as they do in Excel.

- **Swappable font engine, and a font-library-free core**: Text measurement (column auto-fit, row heights, glyph metrics) moved behind `IXLFontEngine`, and the font library ships as a separate package rather than being compiled into the core assembly. The MIT-licensed SkiaSharp engine is the default and auto-registers the first time a workbook is created, so no startup call is needed. This lets you choose a font library whose licence suits you, and stops library authors inheriting a font dependency they don't need. See the Upgrade Guide below.

- **`XLibur.Bundle` meta-package**: Installs the core library together with the default font engine, so a single package reference behaves like ClosedXML out of the box.

- **Editable pictures inside group shapes**: Pictures nested in `xdr:grpSp` groups — at any nesting depth — can be read, resized, moved, added, removed and grouped through a first-class public API. Geometry is computed through the composed group transform, and moves operate in sheet space.

- **DataBar conditional formats can be modified after creation**, including axis settings, rather than being write-once at creation time.

- **Pivot table improvements**: Named ranges resolve as a pivot cache source, and `autoSortScope` on pivot fields round-trips through load/save.

### Fixed

- **Array and dynamic-array formulas no longer break on row/column shifts**: Inserting or deleting rows/columns anywhere in a workbook used to rebuild every formula cell through the `FormulaA1` setter, which turned a single array formula (shared across its whole range) into one *normal* formula per cell. For dynamic arrays this split a single spilled formula such as `=UNIQUE(...)` into multiple implicit-intersection `=@UNIQUE(...)` cells, even when the edit happened on an unrelated sheet. Shifts now update the shared formula instance in place — preserving its array/dynamic-array nature — and relocate the spill range for same-sheet inserts/deletes.

- **Deleting through an array no longer corrupts its stored range**: When a delete overlapped an array formula, relocating the array's top edge could drive the coordinate below 1. `Point` does not bounds-check, so the value silently overflowed and corrupted the stored range.

- **Data-validation formulas are shifted with the sheet**: Inserting or deleting rows/columns relocated each rule's ranges (`sqref`) but left cell references *inside* the criteria formulas (`formula1`/`formula2`) pointing at the pre-shift location. Any `List`, `Custom` or comparison rule referencing other cells silently broke — most visibly dependent dropdown pairs driven by `OFFSET`/`MATCH`. The in-memory value was wrong immediately after the shift, before any save.

- **Data validations no longer vanish when inserting at row 1 or column 1**: The data-validation index was keyed by address at insert time and never re-keyed, so an insert at the first row/column left it stale. At save time the split logic then treated a rule's own out-of-date entry as a competing rule and stripped its ranges, emitting `<dataValidation sqref="">`. Excel rejected the file on open with *"Removed Records: Data validation"*. The index is now reconciled before consolidation.

- **Conditional-format ranges shift once, not twice** ([ClosedXML #2850](https://github.com/ClosedXML/ClosedXML/issues/2850)): Inserting rows or columns below the first line doubled the shift for any rule whose shifted target address collided with another rule's existing range. A rule at `K13` that should move to `K23` landed at `K33`, while rules whose targets happened to be empty shifted correctly.

- **Page breaks no longer inflate the used range** ([ClosedXML #2842](https://github.com/ClosedXML/ClosedXML/issues/2842)): `AddHorizontalPageBreak()`/`AddVerticalPageBreak()` wrote `brk@max` as the sheet's full row/column count. Excel read that as a huge used range, so a file with ~2000 rows of data rendered with a scrollbar spanning all 1,048,576.

- **Named ranges shrink correctly when their first row or column is deleted**: Deleting the first row of a named range shifted both endpoints up instead of removing the deleted row and shifting the survivors, so `A3:A4` became `A2:A3` — expanding the range to include a row that was never part of it. Excel produces `A3:A3`.

- **Totals-row formulas escape column names containing spaces**: Structured references for headers such as `Feb 2023` used the single-bracket form, producing a formula Excel could not parse.

- **Grouped pictures and shapes survive a load/save round-trip** instead of being dropped.

- **Cached formula values are preserved on save**: Cached values are now written whenever they exist and the formula has not been dirtied, regardless of `EvaluateFormulasBeforeSaving`, and the data-type attribute is preserved. This fixes round-trip loss of dynamic-array results (`SORT`, `UNIQUE`, `FILTER`) and spill cell values.

- **Pivot table alignment formatting round-trips**: Alignment in pivot table differential formats (DXF) was silently lost on load/save.

### Performance

- **61% fewer allocations and 16.5% less wall time on load** (250K rows x 15 columns benchmark), from removing per-cell and per-entry garbage in the shared-string reader, cell value/attribute reads, and a new style cache.

- **`<sheetData>` is read with a raw `XmlReader`**: Worksheet loading — the dominant cost when opening a workbook — no longer goes through the OpenXML SDK's `OpenXmlPartReader`, which rebuilt a `ReadOnlyCollection<OpenXmlAttribute>` and materialized text through its object model for every `<c>`, `<row>` and `<f>` element. Measured in isolation on a 250K x 15 sheet (3.75M cells), that reader accounted for ~67% of load time and ~80% of load allocations — roughly 4x slower and 5x more garbage than an equivalent raw `XmlReader` traversal.

- **Faster string cell reads**: `GetValue<string>()`/`GetString()` — the most common cell read — no longer runs a compiled regex over the whole string (allocating a `MatchCollection`) to find the rare `_xHHHH_` escape sequence.

- **Reduced allocations in 10 per-cell, per-formula and per-address hot-path methods**, with no public API or behaviour change.

- **Load and save hot paths**: The shared-string reader is pre-allocated from the SST count, merged cells stream instead of building a full DOM, worksheet attributes are parsed in a single pass, calc-engine overhead is skipped for formula cells during load, and `uint` boxing was removed from the XML writer.

- **`XmlEncoder.EncodeString` fast-path**: Added a character scan that short-circuits before the `Regex` and `StringBuilder` when a string contains no characters that need encoding (the common case for plain text). For workbooks with ~50K unique shared strings this eliminates ~50K `StringBuilder` allocations, ~50K regex evaluations, and ~50K string copies on save.

- **`IXLWorksheet.SetCellValue(int row, int column, XLCellValue value)`** (new API): Sets a cell value directly on the worksheet's internal storage without allocating an intermediate `XLCell` object. For bulk data population (e.g. 50K rows x 3 columns) this eliminates ~150K object allocations that the `Cell(row, col).SetValue(...)` pattern would create.

### Upgrade Guide

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
