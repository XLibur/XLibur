# Spec 10 — Chart Formatting Depth (series styling, data labels, legend, axes)

**Area:** Feature (flagship differentiator — upstream ClosedXML has no charts at all)
**Effort:** L total, but splits into 4 independent PRs
**Dependencies:** None.
**Status:** ✅ All four PRs implemented — see [Results](#results-pr-1) per PR.

## Summary

XLibur's chart support (a fork addition) covers ~75 classic chart types plus 5 ChartEx types, combo charts, anchoring, and title/series/axis emission — but `IXLChartSeries` exposes only name/category/value refs + index/order. Users cannot set a series color, add data labels, control the legend, or title/scale/format an axis. This spec adds the formatting layer that makes generated charts presentation-ready, plus two reader gaps.

## Current state

- Model: `XLibur/Excel/Charts/` (`IXLChart`, `IXLChartSeries`, `XLCharts`, `XLChartType` enum with ~75 values, `SecondaryChartType`/`SecondarySeries` for combos).
- IO: `XLibur/Excel/IO/ChartWriter.cs` (1151 lines), `ChartReader.cs` (549 lines). ChartEx defaults bundled as `ChartExDefaultColors.xml`/`ChartExDefaultStyle.xml`.
- Gaps: no per-series fill/line/marker; no data labels; no legend API; no axis title/scale/number-format; no trendlines/error bars; charts anchored via `OneCellAnchor`/`AbsoluteAnchor` are **skipped on read** (only `TwoCellAnchor` handled); no chart sheets (they fall into `UnsupportedSheets`); no per-series secondary-axis binding.

## Design

Keep the API deliberately smaller than DrawingML: expose the ~15 properties that cover 95% of real chart styling; everything else remains writer defaults. Unknown/unsupported chart XML read from existing files must be **preserved on round-trip** (verify the reader/writer round-trips untouched chart parts rather than regenerating them — if the writer regenerates, preserving loaded raw XML for properties the model doesn't understand is part of PR 1's groundwork; state the finding in the PR).

### PR 1 — Series formatting

```csharp
public interface IXLChartSeries // additions
{
    XLColor? FillColor { get; set; }          // solid fill; null = automatic
    XLColor? LineColor { get; set; }
    double? LineWidthPt { get; set; }
    XLMarkerStyle MarkerStyle { get; set; }    // None/Circle/Square/Diamond/Triangle/X/Auto
    double? MarkerSize { get; set; }
    XLColor? MarkerFillColor { get; set; }
    bool Smooth { get; set; }                  // line charts
    bool UseSecondaryAxis { get; set; }        // binds series to secondary value axis
    IXLDataLabels DataLabels { get; }          // see PR 2
}
```
Writer: emit `c:spPr` (a:solidFill/a:ln) per series, `c:marker`, `c:smooth`; secondary-axis binding moves the series into the secondary plot group (the combo plumbing for `SecondarySeries` exists — generalize it). Reader: parse the same back.

### PR 2 — Data labels

```csharp
public interface IXLDataLabels
{
    bool ShowValue { get; set; }
    bool ShowCategoryName { get; set; }
    bool ShowSeriesName { get; set; }
    bool ShowPercentage { get; set; }          // pie/doughnut
    string? NumberFormat { get; set; }
    XLDataLabelPosition Position { get; set; } // Center/InsideEnd/OutsideEnd/BestFit...
}
```
Per-series (`c:dLbls` under `c:ser`) and chart-level defaults. Position enum validity varies by chart type — validate and throw a clear message for invalid combos (mirror Excel's rules for bar/line/pie only; others accept Center).

### PR 3 — Legend + axis API

```csharp
public interface IXLChartLegend { bool Visible { get; set; } XLLegendPosition Position { get; set; } bool Overlay { get; set; } }
public interface IXLChartAxis
{
    string? Title { get; set; }
    string? NumberFormat { get; set; }
    double? Min { get; set; } double? Max { get; set; }
    double? MajorUnit { get; set; } double? MinorUnit { get; set; }
    bool Visible { get; set; }
    bool MajorGridlines { get; set; }
    XLAxisOrientation Orientation { get; set; }   // MinMax / MaxMin (reversed)
    bool LogScale { get; set; } double LogBase { get; set; }
}
// IXLChart additions: Legend, CategoryAxis, ValueAxis, SecondaryValueAxis (created on demand)
```
The writer already emits axes — this PR parameterizes what it emits (`c:title`, `c:numFmt`, `c:scaling` min/max/orientation/logBase, `c:majorUnit`, `c:majorGridlines`, `c:delete` for hidden) and the reader parses it back.

### PR 4 — Reader gaps

1. Read charts anchored via `OneCellAnchor` and `AbsoluteAnchor` (currently skipped) — map to the existing anchor model; if the model only supports two-cell anchors, add the other anchor kinds to the drawing model (pictures may already have `FreeFloating` placement — reuse that pattern from `XLibur/Excel/Drawings/`).
2. Trendlines/error bars: **round-trip preservation only** (no API) — do not drop them when rewriting a chart whose other properties were edited.

## Work plan

| PR | Content | Size |
|----|---------|------|
| 1 | Series formatting + secondary-axis binding + round-trip-preservation groundwork | L |
| 2 | Data labels | M |
| 3 | Legend + axes | M |
| 4 | Anchor reader gaps + trendline/error-bar preservation | M |

PRs 2 and 3 are independent of each other once PR 1's groundwork lands.

## Acceptance criteria (each PR)

1. Every new property: set via API → save → **open in Excel renders as expected** (manual matrix recorded once per PR) and → reload via XLibur reads the same value back (automated).
2. Excel-authored charts using these features load with correct property values (test resources authored in Excel, checked into `XLibur.Tests/Resource/Charts/`).
3. Round-trip of a chart using *unsupported* features (trendlines, rich gradients) does not lose them (PR 1 groundwork + PR 4).
4. Existing chart tests green; ChartEx types unaffected.
5. Examples: extend `XLibur.Examples` with a "formatted chart" sample — serves as living documentation.

## Risks

- DrawingML defaulting is subtle (automatic colors come from the theme); `null` must mean "omit element" (Excel default), never "emit black". Test against Excel-authored files, not assumptions.
- If the writer regenerates chart XML from the model on save (rather than patching), preservation of unmodeled properties is the hard part — resolve this in PR 1 before building on top.

## Results (PR 1)

Series formatting, secondary-axis binding and the round-trip-preservation groundwork landed.
`IXLChartSeries` gained `FillColor`, `LineColor`, `LineWidthPt`, `MarkerStyle` (new `XLMarkerStyle`
enum), `MarkerSize`, `MarkerFillColor`, `Smooth` and `UseSecondaryAxis`. `DataLabels` is left to
PR 2. 6366 tests pass on net8.0 and net10.0.

### The preservation question is answered: the writer does not regenerate

`ChartWriter.WriteCharts` only ever emitted charts with `IsNew == true`, i.e. charts created through
`Charts.Add`. A chart read from a file was skipped entirely, and because `SaveAs` copies the original
package bytes to the target and then patches it, its chart part passed through **byte for byte** —
trendlines, error bars, gradient fills, per-point formatting and the sibling chart style/colour parts
included. Acceptance criterion 3 was therefore already met before this PR, and there is no need to
stash raw XML for unmodeled properties.

The flip side is that edits to a loaded chart were silently dropped. PR 1 keeps the
never-regenerate rule and adds `ChartPatcher`, which writes back **only** the properties the caller
actually assigned:

- `XLChartSeries` tracks assignments in an `XLChartSeriesFormat` flag set. The reader seeds values
  through `SeedLoadedFormat`, which does not set the flags, so a chart nobody edited is not touched
  at all (`LoadAndSaveWithoutEditsLeavesTheChartPartUntouched` asserts byte equality).
- A patched `c:spPr` / `c:marker` / `c:smooth` replaces just its own child; `cap="rnd"`, `a:round`,
  `a:effectLst`, `c:gapWidth`, `c:trendline` and the neighbouring untouched series all survive.
- `ChartPlotAreaScanner` is shared by the reader and the patcher, so the n-th model series still maps
  to the n-th `c:ser` element on save.

### Secondary axis

`UseSecondaryAxis` splits a chart's series into plot groups: series bound to the secondary axis get
their own chart group referencing a second axis pair (hidden `c:catAx`, right-hand `c:valAx` with
`c:crosses val="max"`). It works for the primary chart type as well as the combo `SecondarySeries`,
which is more than the spec asked for — the spec assumed generalising the existing combo plumbing,
but the combo path plotted both types against the *same* axis pair, so the grouping had to be built
from scratch either way.

Two deliberate limits, both documented on the interface:

- Chart types without a single value axis ignore it — pie and doughnut have none, scatter and bubble
  have two, surface has a series axis.
- Setting it on a chart **loaded from a file** throws `NotSupportedException`. Honouring it would mean
  moving a `c:ser` into a newly created chart group, i.e. regenerating the structure the patch
  approach exists to avoid. Colour and marker properties have no such limit.

### Three pre-existing writer bugs found by switching validation on

The new tests save with `SaveAs(stream, validate: true)`, which runs `OpenXmlValidator`. No chart
test had done that before, and it failed immediately on output the writer had always produced:

1. **Series names were schema-invalid.** `c:tx` held a `c:strRef` with a `c:strCache` but no `c:f`,
   which `CT_StrRef` requires. A literal name belongs in `<c:tx><c:v>` instead, which is what the
   writer now emits; the reader accepts both forms, so files written by earlier versions and by Excel
   (where the name does come from a cell) still read correctly.
2. **`c:doughnutChart` was missing the required `c:holeSize`.** Now written as 75%, Excel's own
   default for a new doughnut chart.
3. **Markers were emitted after `c:cat`/`c:val`.** `CT_LineSer` puts `c:marker` before them. The
   `LineWithMarkers*` types had been writing an out-of-order child all along.

Excel tolerated all three, which is why they went unnoticed. Every chart family is now round-tripped
through the validator by `FormattingSurvivesEveryStandardChartFamily`.

### Reader restructure

`ReadPlotArea` no longer takes the first element of each chart-group type. It scans every group in
the plot area, picks the primary kind by the same precedence the old code implied (bar, bar3D, pie,
doughnut, area, line, radar, bubble, scatter, stock, surface), and merges every group of that kind
into `Series` — which is what makes a two-group secondary-axis chart read back correctly. Groups of
another kind become `SecondaryChartType` / `SecondarySeries` as before. Also fixed:
`DetermineLineChartType` treated `<c:marker><c:symbol val="none"/></c:marker>` as "has markers" and
reported a plain `Line` chart as `LineWithMarkers`.

### Acceptance criteria

| # | Criterion | Status |
|---|---|---|
| 1 | Set via API → save → reload reads the same value | ✅ automated per property |
| 1 | Set via API → save → **renders in Excel** | ⚠️ not executed — no Excel in this environment. The test suite leaves `FormattedChartExamples.xlsx` in the test output directory for a manual pass |
| 2 | Excel-authored charts load with correct values | ⚠️ approximated — `ChartRoundTripPreservationTests` reads a hand-written fixture shaped like Excel's own output (scheme colour, `c:strRef` name cache, `cap`/`a:round`/`a:effectLst`, marker `c:spPr`, secondary axis pair, trendline). No file could be authored in Excel in CI; `XLibur.Tests/Resource/Charts/` is still empty |
| 3 | Round-trip of unsupported features loses nothing | ✅ guaranteed by construction, asserted both ways |
| 4 | Existing chart tests green; ChartEx unaffected | ✅ extended charts ignore series formatting and are not patched |
| 5 | `XLibur.Examples` gains a formatted-chart sample | ✅ `FormattedChartExamples` — four sheets, validated by a test |

### Left for later

- Title, chart type and series references are still write-only for new charts; editing them on a
  loaded chart does nothing. Making `chart.Title` work on a loaded chart is one more patch step on
  `ChartFormatting`.
- Chart sheets still land in `UnsupportedSheets` (spec 09's territory).

## Results (PR 2)

`IXLDataLabels` landed on both `IXLChartSeries.DataLabels` (per series, `c:dLbls` under `c:ser`) and
`IXLChart.DataLabels` (chart-wide, `c:dLbls` on the primary chart group), with `ShowValue`,
`ShowCategoryName`, `ShowSeriesName`, `ShowPercentage`, `NumberFormat` and `Position`. 6383 tests pass
on net8.0 and net10.0.

### It reuses PR 1's machinery, as intended

`XLChartDataLabels` carries its own `XLDataLabelsFormat` assignment flags, the reader seeds values
through `SeedLoaded` without setting them, and `ChartPatcher` patches element by element. A chart
whose labels nobody touched is still not modified; editing one keeps the label font (`c:txPr`), the
label fill, the per-point overrides (`c:dLbl`) and the separator. Two bugs the tests caught, both
from bolting a second concern onto the same code path:

- `ChartPatcher.HasPendingChanges` and `ChartFormatting.PatchSeriesFormat` both bailed out on
  `XLChartSeriesFormat.None`, so a chart where *only* labels had been set was never patched. The
  label step is now outside that gate — worth remembering when PR 3 adds legend and axis flags.
- Data label positions are validated against the chart type, and a series in `SecondarySeries` follows
  `SecondaryChartType`, not `ChartType`. `XLChartSeriesCollection` now knows which of the two it is
  and passes that down, so a line series over columns takes the line positions.

### Position validation

The spec asked for Excel's rules for bar, line and pie, and "others accept Center". The rules landed
as specified for those three families, but **"others" accept only `Auto`, not `Center`** — Excel does
not merely ignore `c:dLblPos` on an area, doughnut, bubble, stock or 3D chart, it refuses to open the
file. Accepting `Center` there would have produced workbooks Excel repairs, which is a worse outcome
than a clear exception. The setter throws `ArgumentException` naming the chart type, the rejected
position and the allowed set; on save an already-set position that a later chart type change
invalidated is dropped rather than written.

`c:dLblPos` is also not the only thing the file format withholds: neither `CT_SurfaceChart` nor
`CT_SurfaceSer` has a `c:dLbls` child at all, so surface charts ignore data labels entirely.

### Emission detail worth knowing

When any label property is set, all six `show*` flags are written, not just the ones that were
assigned. Excel treats a missing flag as "inherit from the chart style", which makes the rendered
result depend on the style part; writing them out makes what the caller asked for unambiguous.
`c:delete` — how Excel records "labels switched off" — is removed when a flag is turned on, otherwise
it would override everything next to it.

## Results (PR 3)

`IXLChartLegend` and `IXLChartAxis` landed as specified, reachable through `IXLChart.Legend`,
`CategoryAxis`, `ValueAxis` and `SecondaryValueAxis`. 6402 tests pass on net8.0 and net10.0.

### The writer now parameterises the axes it already emitted

`BuildCategoryAxis`/`BuildValueAxis` take the axis model and fill in `c:scaling`
(`logBase`/`orientation`/`max`/`min`, in that order), `c:delete`, `c:majorGridlines`, `c:title`,
`c:numFmt` and — after `c:crossAx` — `c:majorUnit`/`c:minorUnit`. The defaults reproduce exactly what
the writer emitted before this PR, so an unformatted chart's XML is unchanged: `delete` false,
`orientation` minMax, nothing else. `c:legend` is emitted between `c:plotArea` and `c:plotVisOnly`
only when `Legend.Visible` is set.

### Three places the file format is narrower than one shared axis interface

- **`CT_CatAx` has no unit elements and no room for `c:logBase`.** `MajorUnit`, `MinorUnit` and
  `LogScale` are therefore skipped on a category axis rather than written — Excel refuses a file that
  has them. The spec put all of them on one `IXLChartAxis`, which is the right shape for callers; the
  narrowing is documented on the interface and in the docs.
- **Except on scatter and bubble charts.** There the horizontal axis is a `c:valAx` holding numbers,
  so the same properties do apply. `XLChartAxis` asks the chart for its type to decide, which is why
  the axis models are constructed with a back-reference to the chart rather than standing alone.
- **The hidden helper category axis of a secondary group has no public counterpart.** It exists only
  to give the secondary value axis something to cross, so `BuildCategoryAxis` accepts a null model for
  it. `SecondaryValueAxis` maps to the visible right-hand axis, which is the one callers mean.

### Patching

The legend and both axes follow PR 2's pattern: their own assignment flag sets, seeded (not assigned)
by the reader, patched element by element. `ChartPlotAreaScanner` gained `CategoryAxisId` and a
`FindAxis` helper so the reader and the patcher locate the same axis element by the identifier the
chart group points at — the axis elements themselves carry no marker saying which is which.

Editing an axis keeps its tick marks, tick label position, `c:crossBetween`, line and text formatting;
editing the legend keeps its layout and text properties. `Legend.Visible = false` removes the
`c:legend` element, which is how Excel records a chart without a legend.

### Gridlines are still off by default

`MajorGridlines` defaults to false, so a chart XLibur creates has no gridlines unless asked — which is
what the writer did before this PR. Excel's own new charts do have value axis gridlines, so this is a
plausible thing to change, but doing it silently would alter the output of every existing caller. Left
as a deliberate decision to revisit, not an oversight.

## Results (PR 4)

Both reader gaps closed. 6416 tests pass on net8.0 and net10.0.

### Anchors

`ChartReader.LoadCharts` iterated `Elements<Xdr.TwoCellAnchor>()`, so a chart Excel had anchored with
`xdr:oneCellAnchor` or `xdr:absoluteAnchor` never reached `ws.Charts` — invisible to the API, though
its XML did survive a save because the writer never touches a loaded chart. It now walks every anchor
child.

Following the picture pattern the spec pointed at, the anchor kind is exposed as
`IXLChart.Anchor` of the already-public `XLDrawingAnchor` enum, whose three values line up exactly
with the three anchor elements, plus `Width`, `Height`, `Left` and `Top` in pixels — the same unit and
the same 9525 EMU conversion `IXLPicture` uses. `XLPicturePlacement` would have been the closer
analogue by name, but it lives in `XLibur.Excel.Drawings` and reads oddly on a chart;
`XLDrawingAnchor` is in `XLibur.Excel` and was already public.

The writer emits whichever anchor `Anchor` asks for. It defaults to `MoveAndSizeWithCells`, so charts
built by existing callers are byte-identical to before.

### Trendlines and error bars

Preservation-only, as specified, and it needed no code: PR 1's patch approach never rewrites a series
element wholesale. The fixture in `ChartRoundTripPreservationTests` now carries both a `c:trendline`
and a `c:errBars`, and the test that edits the series' fill and line width asserts they are still
there afterwards.

### Also closed: the 3D and of-pie group kinds

Not in the spec's PR 4 list, but the same class of bug and one line of the spec's "Current state":
`c:pie3DChart`, `c:line3DChart`, `c:area3DChart`, `c:surface3DChart` and `c:ofPieChart` were not
recognised by the plot-area scan, so an Excel-authored chart using one produced a chart object with no
series and whatever `XLChartType` happens to be zero. They are now scanned like the 2D groups, their
series and series formatting read the same way, `c:ofPieType` tells pie-of-pie from bar-of-pie, and
`c:grouping` on `c:area3DChart` picks the stacked variants.

**The writer is still asymmetric here** and this PR deliberately did not change it: XLibur writes
`Pie3D` as a plain `c:pieChart`, `Line3D` as `c:lineChart`, `Area3D` as `c:areaChart` and every surface
type as `c:surfaceChart`. Its own 3D charts therefore round-trip as their 2D equivalents. Fixing the
writer would change the output of existing callers and is a separate, larger job — the 3D groups carry
`c:view3D`, `c:floor`, `c:sideWall` and `c:backWall` on the chart, none of which is modelled. For the
same reason `c:surfaceChart` is still read back as `Surface` rather than the `SurfaceContour` it
strictly means: the writer emits it for `Surface`, and matching the reader to the spec's semantics
would break XLibur's own round trip.
