# Spec 10 — Chart Formatting Depth (series styling, data labels, legend, axes)

**Area:** Feature (flagship differentiator — upstream ClosedXML has no charts at all)
**Effort:** L total, but splits into 4 independent PRs
**Dependencies:** None.
**Status:** Proposed

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
