# Spec 10 — Chart Formatting Depth (series styling, data labels, legend, axes)

**Area:** Feature (flagship differentiator — upstream ClosedXML has no charts at all)
**Effort:** L total, but splits into 4 independent PRs
**Dependencies:** None.
**Status:** In progress — PR 1 implemented; see [Results](#results-pr-1). PRs 2–4 open.

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

### Notes for PRs 2–4

- Build data labels, legend and axes on `ChartPatcher`/`ChartFormatting`, not on a second mechanism:
  add flags to the assignment set and a patch step per element. The "only write what was assigned"
  rule is what keeps preservation free.
- `ChartPlotAreaScanner` is the place to add group kinds the reader still ignores: `c:pie3DChart`,
  `c:line3DChart`, `c:area3DChart`, `c:surface3DChart`, `c:ofPieChart`. Today a chart built from
  those reads as zero series with a default chart type — worth folding into PR 3 or 4.
- Title, chart type and series references are still write-only for new charts; editing them on a
  loaded chart does nothing. If PR 3 wants `chart.Title` to work on loaded charts, that is another
  patch step.
