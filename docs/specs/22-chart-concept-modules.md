# Spec 22 — Chart IO: one module per chart concept, not three per lifecycle

**Area:** Architecture · Refactor · Correctness (latent)
**Effort:** M (~1–1.5 weeks)
**Dependencies:** **Hard prerequisite: spec 16 must land first** — all three of its tasks. Spec 16
extracts the DrawingML property layer out of `ChartFormatting.cs`, and its task 1 harness is the gate
this spec is measured with. Conflicts with specs 15 and 17 only through 16.
**Status:** ✅ Done — tasks 0–6, branch `task/22`. Run **before** spec 16, which had not been
scheduled; 16 now rebases onto `Charts/ChartSeriesFormatXml.cs`, where the DrawingML colour, fill
and outline helpers ended up. Deviations and final counts are in the spec 22 section of
[TASKLIST-architecture-deepening.md](TASKLIST-architecture-deepening.md).

## Goal

Split chart formatting by **concept** (legend, title, axis, series format, data labels) instead of by
**lifecycle** (build / read / patch), so each concept's OpenXML mapping lives in one module with one
interface.

## Why this spec exists

`ChartFormatting` is the only shallow module in an otherwise deep IO layer. Every other part reader
and writer under `XLibur/Excel/IO/` exposes one or two entry points —
`PivotTableDefinitionPartReader.Load`, `PivotTableDefinitionPartWriter2.WriteContent`,
`ChartReader.LoadCharts`, `ChartPatcher.PatchChart`. `ChartFormatting` exposes **21**, and three
modules call them at **29 sites**:

| Caller | `ChartFormatting.*` call sites |
|---|---:|
| `ChartWriter.cs` | 13 |
| `ChartPatcher.cs` | 9 |
| `ChartReader.cs` | 7 |

The split is by lifecycle, so one concept is three functions:

| Concept | Build (new chart) | Read (load) | Patch (edit loaded) |
|---|---|---|---|
| Legend | `BuildLegend` | `ReadLegend` | `PatchLegend` |
| Title | `BuildTitleText` | — | `PatchTitle` |
| Extended title | `BuildExtendedTitle` | — | `PatchExtendedTitle` |
| Axis | `BuildScaling`, `AppendAxisBody`, `AppendAxisUnits` | `ReadAxis` | `PatchAxis` |
| Series format | `BuildSeriesShapeProperties`, `BuildMarker`, `BuildSmooth` | `ReadSeriesFormat` | `PatchSeriesFormat` |
| Data labels | `BuildDataLabels` | `ReadDataLabels` | `PatchSeriesDataLabels`, `PatchGroupDataLabels`, `SupportsDataLabels` |

### The agreement is currently maintained by comment

`BuildLegend` and the `element == null` branch of `PatchLegend` construct the same `C.Legend` with the
same two children, inline, twice. `PatchLegend` carries this comment
(`ChartFormatting.cs:438-440`):

```csharp
// Position and Overlay are ignored while the legend is hidden, so assigning one of them on
// a chart that has no legend must not conjure one — BuildLegend does not either.
```

That is a cross-reference between two functions that must agree and have no structural reason to.
The same shape repeats for title, axis and data labels. This is the defect class the spec removes:
not a bug that exists today, but the mechanism by which one arrives.

### The unification is exact, not approximate

`Build` is `Apply` against an absent element. Checked case by case for the legend:

| Model state | `BuildLegend` (new chart) | `PatchLegend` (loaded, no `c:legend`) | Same? |
|---|---|---|---|
| `AssignedFormat == None` | returns `null` → no element | returns early → no element | ✅ |
| assigned, `Visible == false` | returns `null` → no element | `element?.Remove()` on nothing → no element | ✅ |
| assigned, `Visible == true` | builds `c:legend` + position + overlay | builds `c:legend` + position + overlay | ✅ |

A single `Apply(parent, model)` that creates the element when absent and patches it when present
reproduces both columns. `InsertOrdered` already handles an empty parent, so ordering is unchanged.
**No `isNew` flag is required** — this is what makes the collapse clean rather than a merge of two
paths behind a boolean.

## Non-goals

- **No new chart capability.** Behaviour-preserving refactor only. The open 3D-group gap from spec 10
  lives in `ChartWriter.cs` (`AppendBar3DChart`) and is untouched.
- **No change to the round-trip guarantee.** A loaded chart is still never regenerated; only assigned
  properties are patched into the part it was read from. See `ChartPatcher`'s class remarks.
- **No DrawingML layer work.** Spec 16 owns that and lands first.
- **No public API change.** Everything added is `internal`.

## Current state

Verified against the tree at `d05b0753` (2026-08-23).

| File | Lines | Role |
|---|---:|---|
| `XLibur/Excel/IO/ChartFormatting.cs` | 1,453 | 21 internal entry points; the subject |
| `XLibur/Excel/IO/ChartWriter.cs` | 1,299 | builds new charts; 13 call sites |
| `XLibur/Excel/IO/ChartReader.cs` | 605 | loads charts; 7 call sites |
| `XLibur/Excel/IO/ChartPatcher.cs` | 212 | patches loaded charts; 9 call sites |
| `XLibur/Excel/IO/ChartPlotAreaScanner.cs` | 266 | plot-area group scan; unchanged by this spec |

Test surface today: `XLibur.Tests/Excel/Charts/` — 3,836 lines across 9 files, of which
`ChartRoundTripPreservationTests.cs` is 1,152. `ChartPartFixture.cs` feeds hand-written chart XML
through a real workbook, which is how reader paths are tested without XLibur having to be able to
write the shape first.

**After spec 16 task 3**, `ChartFormatting` is smaller: `SetFill`/`SetOutline`/`BuildColor`/
`MapSchemeColor`/ordering predicates move to `IO/DrawingML/ShapePropertiesWriter`, and
`ChartFormatting` keeps thin mask-driven adapters. **Re-read `ChartFormatting.cs` before starting** —
the line numbers in this spec are from before that move.

## File structure

```
XLibur/Excel/IO/Charts/            (new folder)
  ChartLegendXml.cs                Read + Apply for c:legend
  ChartTitleXml.cs                 Read + Apply for c:title and cx:title
  ChartAxisXml.cs                  Read + Apply for c:catAx / c:valAx / c:dateAx
  ChartSeriesFormatXml.cs          Read + Apply for a series' spPr / marker / smooth
  ChartDataLabelsXml.cs            Read + Apply for c:dLbls at series and group level
XLibur/Excel/IO/ChartFormatting.cs deleted at the end of task 6
```

Each module is a static class with the same two-method shape. That uniformity is the point: a reader
who has understood one has understood all five.

## Interfaces

Every concept module exposes exactly this pair, plus concept-specific predicates only where a caller
genuinely needs one before it has a parent element.

```csharp
internal static class ChartLegendXml
{
    /// <summary>Seeds <paramref name="legend"/> from the chart's c:legend, or from its absence.</summary>
    internal static void Read(C.Chart chart, XLChartLegend legend);

    /// <summary>
    /// Writes the assigned legend properties into <paramref name="chart"/>, creating, editing or
    /// removing the c:legend child as the model requires. A chart with no assigned legend
    /// properties is not modified.
    /// </summary>
    internal static void Apply(C.Chart chart, XLChartLegend legend);
}
```

```csharp
internal static class ChartTitleXml
{
    internal static void Apply(C.Chart chart, XLChart xlChart);
    internal static void ApplyExtended(Cx.Chart chart, XLChart xlChart);
}
```

```csharp
internal static class ChartAxisXml
{
    internal static void Read(OpenXmlCompositeElement? axis, XLChartAxis model);
    internal static void Apply(OpenXmlCompositeElement axis, XLChartAxis model);
}
```

```csharp
internal static class ChartSeriesFormatXml
{
    internal static void Read(C.BarChartSeries seriesElement, XLChartSeries series, bool useSecondaryAxis);
    internal static void Apply(OpenXmlCompositeElement seriesElement, XLChartSeries series,
        XLChartGroupKind kind, XLChartType chartType);
}
```

```csharp
internal static class ChartDataLabelsXml
{
    internal static bool Supports(XLChartGroupKind kind);
    internal static void Read(C.DataLabels? dataLabels, XLChartDataLabels labels);
    internal static void Apply(OpenXmlCompositeElement parent, XLChartDataLabels labels, XLChartType chartType);
}
```

`ChartSeriesFormatXml.Read` keeps its concrete `C.BarChartSeries` parameter — that is what
`ChartReader.cs:231` passes today, and widening it is a separate change with no consumer.

## Global constraints

- Warnings are errors (`TreatWarningsAsErrors=true`); nullable enabled — new code must be annotated.
- Branch per task; never commit to main. Commit prefix `refactor:` for tasks 1–6, `test:` for task 0.
- No compound shell commands (`&&`, `;`) in agent tool calls.
- Do not upgrade SixLabors.Fonts.
- Build: `dotnet build XLibur/XLibur.csproj -c Release -v q`
- Tests: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
- Chart subset: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/Chart*/*"`
  — use `--treenode-filter`, never `--filter` (exit 5 = bad option, exit 8 = zero matched).
- Never filter at solution level; always name the `.csproj`.

## Work plan

Tasks are ordered. Each ends with a green suite and a commit. Tasks 2–5 are mutually independent
once task 1 lands, but all five touch `ChartFormatting.cs`, so **run them sequentially in one
branch** rather than in parallel worktrees.

| # | Task | Size | Gate |
|---|---|---|---|
| 0 | Golden byte-identity baseline for the chart corpus | S | Golden test green on unmodified code |
| 1 | `ChartLegendXml` — the pattern-setting extraction | S | Chart suite green; golden identical |
| 2 | `ChartTitleXml` | S | as above |
| 3 | `ChartAxisXml` | M | as above |
| 4 | `ChartDataLabelsXml` | M | as above |
| 5 | `ChartSeriesFormatXml` | M | as above |
| 6 | Delete `ChartFormatting.cs`; collapse callers | S | Full suite green; golden identical |

---

### Task 0 — Golden byte-identity baseline

Spec 16 task 1 builds an XML change-set harness and a golden corpus. **If spec 16 landed, reuse it
and skip to step 4 to add the chart-writer fixtures this spec needs.** This task exists so spec 22
is executable even if 16's corpus did not cover the new-chart write path.

**Files:**
- Create: `XLibur.Tests/Excel/Charts/ChartGoldenCorpusTests.cs`
- Create: `XLibur.Tests/Resource/Other/Charts/Golden/` (committed `.xml` fixtures)

**Interfaces:**
- Produces: `ChartGoldenCorpus.CaptureChartPartXml(Action<IXLWorksheet> build) → string`, used by
  tasks 1–6 to prove byte-identity.

- [ ] **Step 1: Write the capture helper**

```csharp
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Charts;

internal static class ChartGoldenCorpus
{
    /// <summary>
    /// Saves a workbook built by <paramref name="build"/> and returns the raw XML of its first
    /// chart part, verbatim. Byte-identity of this string across a refactor is the gate every
    /// task in spec 22 is measured with.
    /// </summary>
    internal static string CaptureChartPartXml(Action<IXLWorksheet> build)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Data");
            ws.Cell("A1").Value = "Q1";
            ws.Cell("A2").Value = "Q2";
            ws.Cell("B1").Value = 100;
            ws.Cell("B2").Value = 200;
            build(ws);
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var chartPart = doc.WorkbookPart!
            .WorksheetParts
            .SelectMany(p => p.DrawingsPart is null ? [] : p.DrawingsPart.ChartParts)
            .First();

        using var stream = chartPart.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
```

- [ ] **Step 2: Write the golden test over five chart shapes**

```csharp
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Charts;

/// <summary>
/// Pins the exact chart-part XML XLibur writes for a representative set of new charts. Spec 22
/// reorganises the code that produces this XML without changing it; any diff here is a finding to
/// investigate, never noise to re-baseline without a written explanation.
/// </summary>
public class ChartGoldenCorpusTests
{
    private const string GoldenDir = "Excel/Charts/Golden";

    [Test]
    [Arguments("bar-plain")]
    [Arguments("line-legend-bottom")]
    [Arguments("bar-titled")]
    [Arguments("line-datalabels")]
    [Arguments("bar-secondary-axis")]
    public async Task Chart_part_xml_matches_the_golden_fixture(string name)
    {
        var actual = ChartGoldenCorpus.CaptureChartPartXml(ws => BuildFixture(name, ws));
        var path = Path.Combine(GoldenDir, name + ".xml");

        if (!File.Exists(path))
        {
            Directory.CreateDirectory(GoldenDir);
            File.WriteAllText(path, actual);
        }

        await Assert.That(actual).IsEqualTo(File.ReadAllText(path));
    }

    private static void BuildFixture(string name, IXLWorksheet ws)
    {
        switch (name)
        {
            case "bar-plain":
                ws.Charts.Add(XLChartType.Bar).SetSourceData(ws.Range("A1:B2"));
                break;
            case "line-legend-bottom":
            {
                var chart = ws.Charts.Add(XLChartType.Line);
                chart.SetSourceData(ws.Range("A1:B2"));
                chart.Legend.Position = XLLegendPosition.Bottom;
                break;
            }
            case "bar-titled":
            {
                var chart = ws.Charts.Add(XLChartType.Bar);
                chart.SetSourceData(ws.Range("A1:B2"));
                chart.Title = "Quarterly";
                break;
            }
            case "line-datalabels":
            {
                var chart = ws.Charts.Add(XLChartType.Line);
                chart.SetSourceData(ws.Range("A1:B2"));
                chart.DataLabels.ShowValue = true;
                break;
            }
            case "bar-secondary-axis":
            {
                var chart = ws.Charts.Add(XLChartType.Bar);
                chart.SetSourceData(ws.Range("A1:B2"));
                chart.SecondaryValueAxis.MajorGridlines = true;
                break;
            }
        }
    }
}
```

- [ ] **Step 3: Run it twice — once to write the fixtures, once to assert them**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/ChartGoldenCorpusTests/*"`
Expected: PASS on the first run (fixtures written), PASS on the second (fixtures asserted).

- [ ] **Step 4: Verify the gate actually bites**

Temporarily change `ChartFormatting.MapLegendPosition` so `XLLegendPosition.Bottom` maps to
`C.LegendPositionValues.Top`. Re-run.
Expected: FAIL on `line-legend-bottom`. Revert the change.

If it does not fail, the corpus does not cover the code you are about to move — widen it before
continuing. **A refactor gated by a test that cannot fail is not gated.**

- [ ] **Step 5: Commit**

```bash
git add XLibur.Tests/Excel/Charts/ChartGoldenCorpusTests.cs XLibur.Tests/Excel/Charts/ChartGoldenCorpus.cs XLibur.Tests/Resource/Other/Charts/Golden
git commit -m 'test(charts): pin chart-part XML with a golden corpus (spec 22 task 0)'
```

---

### Task 1 — `ChartLegendXml`

The pattern-setting extraction. Do this one first and review it carefully; tasks 2–5 copy its shape.

**Files:**
- Create: `XLibur/Excel/IO/Charts/ChartLegendXml.cs`
- Modify: `XLibur/Excel/IO/ChartFormatting.cs` — delete `BuildLegend`, `ReadLegend`, `PatchLegend`,
  `ReadLegendPosition`, `MapLegendPosition`
- Modify: `XLibur/Excel/IO/ChartWriter.cs:393`
- Modify: `XLibur/Excel/IO/ChartReader.cs:100`
- Modify: `XLibur/Excel/IO/ChartPatcher.cs:57`
- Test: `XLibur.Tests/Excel/Charts/ChartLegendAndAxisTests.cs` (unmodified — it is the gate)

**Interfaces:**
- Consumes: `ChartGoldenCorpus.CaptureChartPartXml` from task 0.
- Produces: `ChartLegendXml.Read(C.Chart, XLChartLegend)`, `ChartLegendXml.Apply(C.Chart, XLChartLegend)`.

- [ ] **Step 1: Write the failing test — Build and Apply agree on an absent element**

This is the property the current code maintains by comment. Add to
`XLibur.Tests/Excel/Charts/ChartLegendAndAxisTests.cs`:

```csharp
/// <summary>
/// A new chart and a loaded chart with no c:legend must reach the same XML from the same model.
/// Before spec 22 these were two functions (BuildLegend and PatchLegend's null-element branch)
/// agreeing by hand; afterwards they are one.
/// </summary>
[Test]
[Arguments(XLLegendPosition.Bottom)]
[Arguments(XLLegendPosition.Left)]
[Arguments(XLLegendPosition.TopRight)]
public async Task A_new_chart_and_a_legendless_loaded_chart_write_the_same_legend(
    XLLegendPosition position)
{
    var fromNew = ChartGoldenCorpus.CaptureChartPartXml(ws =>
    {
        var chart = ws.Charts.Add(XLChartType.Bar);
        chart.SetSourceData(ws.Range("A1:B2"));
        chart.Legend.Position = position;
    });

    // Round-trip a chart that was saved with no legend, then assign the same position.
    using var ms = new MemoryStream();
    using (var wb = new XLWorkbook())
    {
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Q1";
        ws.Cell("A2").Value = "Q2";
        ws.Cell("B1").Value = 100;
        ws.Cell("B2").Value = 200;
        var chart = ws.Charts.Add(XLChartType.Bar);
        chart.SetSourceData(ws.Range("A1:B2"));
        chart.Legend.Visible = false;
        wb.SaveAs(ms);
    }

    ms.Position = 0;
    using var reloaded = new XLWorkbook(ms);
    reloaded.Worksheet("Data").Charts.First().Legend.Position = position;

    await Assert.That(ReadLegendPositionFromSave(reloaded))
        .IsEqualTo(ReadLegendPositionFromXml(fromNew));
}
```

Add the two small readers as private helpers in the same file:

```csharp
private static string? ReadLegendPositionFromXml(string chartPartXml)
{
    var doc = System.Xml.Linq.XDocument.Parse(chartPartXml);
    System.Xml.Linq.XNamespace c =
        "http://schemas.openxmlformats.org/drawingml/2006/chart";
    return doc.Descendants(c + "legendPos").FirstOrDefault()?.Attribute("val")?.Value;
}

private static string? ReadLegendPositionFromSave(XLWorkbook wb)
{
    using var ms = new MemoryStream();
    wb.SaveAs(ms);
    ms.Position = 0;
    using var doc = SpreadsheetDocument.Open(ms, false);
    var part = doc.WorkbookPart!.WorksheetParts
        .SelectMany(p => p.DrawingsPart is null ? [] : p.DrawingsPart.ChartParts).First();
    using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
    using var reader = new StreamReader(stream);
    return ReadLegendPositionFromXml(reader.ReadToEnd());
}
```

- [ ] **Step 2: Run it to confirm it passes on current code**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/ChartLegendAndAxisTests/*"`
Expected: PASS. This is a **characterization** test — it pins behaviour that is already correct so the
refactor cannot silently change it. If it fails on unmodified code you have found a real bug; stop and
report it before refactoring.

- [ ] **Step 3: Create `ChartLegendXml` with `Read` and `Apply`**

```csharp
using System;
using System.Linq;
using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XLibur.Excel.IO.Charts;

/// <summary>
/// The c:legend element: how it is read into <see cref="XLChartLegend"/> and how an assigned model
/// is written back.
/// </summary>
/// <remarks>
/// <see cref="Apply"/> covers both a chart being created and a chart loaded from a file. The two
/// used to be separate functions — <c>BuildLegend</c> and <c>PatchLegend</c> — that had to agree by
/// hand; creating the element is simply the branch <see cref="Apply"/> takes when there is none.
/// </remarks>
internal static class ChartLegendXml
{
    /// <summary>
    /// Seeds the model from the chart's c:legend. A chart with no legend element seeds
    /// <see cref="IXLChartLegend.Visible"/> as <c>false</c>.
    /// </summary>
    internal static void Read(C.Chart chart, XLChartLegend legend)
    {
        var element = chart.Elements<C.Legend>().FirstOrDefault();
        if (element == null)
        {
            legend.SeedLoaded(visible: false, XLLegendPosition.Right, overlay: false);
            return;
        }

        legend.SeedLoaded(
            visible: true,
            position: ReadPosition(element),
            overlay: element.Elements<C.Overlay>().FirstOrDefault()?.Val?.Value ?? false);
    }

    /// <summary>
    /// Writes the assigned legend properties into <paramref name="chart"/>. A chart with no
    /// assigned legend properties is not modified at all, which is what lets an untouched legend
    /// round-trip byte for byte.
    /// </summary>
    internal static void Apply(C.Chart chart, XLChartLegend legend)
    {
        var assigned = legend.AssignedFormat;
        if (assigned == XLChartLegendFormat.None)
            return;

        var element = chart.Elements<C.Legend>().FirstOrDefault();

        if ((assigned & XLChartLegendFormat.Visible) != 0 && !legend.Visible)
        {
            element?.Remove();
            return;
        }

        if (element == null)
        {
            // Position and Overlay are ignored while the legend is hidden, so assigning one of
            // them on a chart that has no legend must not conjure one.
            if (!legend.Visible)
                return;

            element = new C.Legend();
            element.Append(new C.LegendPosition { Val = MapPosition(legend.Position) });
            element.Append(new C.Overlay { Val = legend.Overlay });
            ChartFormatting.InsertOrdered(chart, element, ChartFormatting.ChartChildOrder);
            return;
        }

        if ((assigned & XLChartLegendFormat.Position) != 0)
        {
            foreach (var existing in element.Elements<C.LegendPosition>().ToList())
                existing.Remove();
            ChartFormatting.InsertOrdered(element,
                new C.LegendPosition { Val = MapPosition(legend.Position) },
                ChartFormatting.LegendChildOrder);
        }

        if ((assigned & XLChartLegendFormat.Overlay) != 0)
        {
            foreach (var existing in element.Elements<C.Overlay>().ToList())
                existing.Remove();
            ChartFormatting.InsertOrdered(element, new C.Overlay { Val = legend.Overlay },
                ChartFormatting.LegendChildOrder);
        }
    }

    private static XLLegendPosition ReadPosition(C.Legend element)
    {
        var position = element.Elements<C.LegendPosition>().FirstOrDefault()?.Val;
        if (position == null)
            return XLLegendPosition.Right;

        var value = position.Value;
        if (value == C.LegendPositionValues.Bottom) return XLLegendPosition.Bottom;
        if (value == C.LegendPositionValues.Left) return XLLegendPosition.Left;
        if (value == C.LegendPositionValues.Top) return XLLegendPosition.Top;
        if (value == C.LegendPositionValues.TopRight) return XLLegendPosition.TopRight;
        return XLLegendPosition.Right;
    }

    private static C.LegendPositionValues MapPosition(XLLegendPosition position) => position switch
    {
        XLLegendPosition.Right => C.LegendPositionValues.Right,
        XLLegendPosition.Bottom => C.LegendPositionValues.Bottom,
        XLLegendPosition.Left => C.LegendPositionValues.Left,
        XLLegendPosition.Top => C.LegendPositionValues.Top,
        XLLegendPosition.TopRight => C.LegendPositionValues.TopRight,
        _ => throw new ArgumentOutOfRangeException(nameof(position), position,
            "Unknown legend position.")
    };
}
```

`InsertOrdered`, `ChartChildOrder` and `LegendChildOrder` are currently `private` in
`ChartFormatting`. Widen them to `internal` for now; task 6 moves them to a shared
`ChartElementOrder` type once every concept module needs them.

- [ ] **Step 4: Repoint the three callers**

`ChartWriter.cs:393` — the new-chart path stops calling `BuildLegend` and appending its result.
Replace:

```csharp
var legend = ChartFormatting.BuildLegend(xlChart.LegendInternal);
if (legend != null)
    chart.Append(legend);
```

with:

```csharp
ChartLegendXml.Apply(chart, xlChart.LegendInternal);
```

`ChartReader.cs:100` — replace:

```csharp
ChartFormatting.ReadLegend(chart.Elements<C.Legend>().FirstOrDefault(), xlChart.LegendInternal);
```

with:

```csharp
ChartLegendXml.Read(chart, xlChart.LegendInternal);
```

`ChartPatcher.cs:57` — replace `ChartFormatting.PatchLegend(chart, xlChart.LegendInternal);` with
`ChartLegendXml.Apply(chart, xlChart.LegendInternal);`.

Add `using XLibur.Excel.IO.Charts;` to all three files.

- [ ] **Step 5: Delete the five members from `ChartFormatting`**

Delete `BuildLegend`, `ReadLegend`, `PatchLegend`, `ReadLegendPosition`, `MapLegendPosition`.
Interface drops 21 → 18.

- [ ] **Step 6: Run the chart suite and the golden corpus**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/Chart*/*"`
Expected: PASS, with **no test file modified except the one test added in step 1**. The golden
fixtures must be byte-identical — if `line-legend-bottom` diffs, the `Apply`-on-new-chart path is not
reproducing `BuildLegend`'s output. Investigate; do not re-baseline.

- [ ] **Step 7: Commit**

```bash
git add XLibur/Excel/IO/Charts/ChartLegendXml.cs XLibur/Excel/IO/ChartFormatting.cs XLibur/Excel/IO/ChartWriter.cs XLibur/Excel/IO/ChartReader.cs XLibur/Excel/IO/ChartPatcher.cs XLibur.Tests/Excel/Charts/ChartLegendAndAxisTests.cs
git commit -m 'refactor(charts): give the legend one module with read and apply (spec 22 task 1)'
```

---

### Task 2 — `ChartTitleXml`

**Files:**
- Create: `XLibur/Excel/IO/Charts/ChartTitleXml.cs`
- Modify: `ChartFormatting.cs` — delete `PatchTitle`, `BuildTitleText`, `PatchExtendedTitle`,
  `BuildExtendedTitle`, and the `SetRichText` helpers they own
- Modify: `ChartWriter.cs:284`, `ChartWriter.cs:387`, `ChartPatcher.cs:56`, `ChartPatcher.cs:189`

**Interfaces:**
- Consumes: nothing from tasks 1 or 0 beyond the golden corpus.
- Produces: `ChartTitleXml.Apply(C.Chart, XLChart)`, `ChartTitleXml.ApplyExtended(Cx.Chart, XLChart)`.

- [ ] **Step 1: Write the characterization test**

Add to `XLibur.Tests/Excel/Charts/ChartTests.cs`:

```csharp
/// <summary>
/// The title text a new chart writes and the title text a loaded chart is patched to must match.
/// </summary>
[Test]
public async Task A_new_chart_and_a_reloaded_chart_carry_the_same_title()
{
    const string title = "Quarterly revenue";

    var fromNew = ChartGoldenCorpus.CaptureChartPartXml(ws =>
    {
        var chart = ws.Charts.Add(XLChartType.Bar);
        chart.SetSourceData(ws.Range("A1:B2"));
        chart.Title = title;
    });

    using var ms = new MemoryStream();
    using (var wb = new XLWorkbook())
    {
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Q1";
        ws.Cell("B1").Value = 100;
        var chart = ws.Charts.Add(XLChartType.Bar);
        chart.SetSourceData(ws.Range("A1:B1"));
        wb.SaveAs(ms);
    }

    ms.Position = 0;
    using var reloaded = new XLWorkbook(ms);
    reloaded.Worksheet("Data").Charts.First().Title = title;

    using var out2 = new MemoryStream();
    reloaded.SaveAs(out2);

    await Assert.That(FirstTitleText(out2)).IsEqualTo(title);
    await Assert.That(fromNew).Contains(title);
}

private static string? FirstTitleText(MemoryStream saved)
{
    saved.Position = 0;
    using var doc = SpreadsheetDocument.Open(saved, false);
    var part = doc.WorkbookPart!.WorksheetParts
        .SelectMany(p => p.DrawingsPart is null ? [] : p.DrawingsPart.ChartParts).First();
    return part.ChartSpace!.Descendants<DocumentFormat.OpenXml.Drawing.Text>()
        .FirstOrDefault()?.Text;
}
```

- [ ] **Step 2: Run it — expect PASS on current code**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/ChartTests/*"`
Expected: PASS.

- [ ] **Step 3: Move the four title members into `ChartTitleXml`**

Move `PatchTitle` → `Apply`, `PatchExtendedTitle` → `ApplyExtended`, and fold `BuildTitleText` and
`BuildExtendedTitle` into them as the absent-element branch, exactly as task 1 folded `BuildLegend`
into `Apply`. Keep `SetRichText` and its `IsRunLevel` / `endParaRPr` rules with the title — spec 16
deliberately left them in place for spec 15 to relocate later, and this spec does not move them out
of the chart tree, only into the title's own module.

- [ ] **Step 4: Repoint the four callers**

`ChartWriter.cs:387` (standard) and `ChartWriter.cs:284` (extended) call `Apply` / `ApplyExtended`
against the chart element they have just built, instead of appending a built title.
`ChartPatcher.cs:56` and `:189` call the same two methods.

- [ ] **Step 5: Run the chart suite and golden corpus**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/Chart*/*"`
Expected: PASS; `bar-titled` byte-identical.

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel/IO/Charts/ChartTitleXml.cs XLibur/Excel/IO/ChartFormatting.cs XLibur/Excel/IO/ChartWriter.cs XLibur/Excel/IO/ChartPatcher.cs XLibur.Tests/Excel/Charts/ChartTests.cs
git commit -m 'refactor(charts): give the title one module with apply (spec 22 task 2)'
```

---

### Task 3 — `ChartAxisXml`

The widest concept: three build-side entry points (`BuildScaling`, `AppendAxisBody`,
`AppendAxisUnits`) plus `ReadAxis` and `PatchAxis`, called at eight sites.

**Files:**
- Create: `XLibur/Excel/IO/Charts/ChartAxisXml.cs`
- Modify: `ChartFormatting.cs` — delete `BuildScaling`, `AppendAxisBody`, `AppendAxisUnits`,
  `ReadAxis`, `PatchAxis`
- Modify: `ChartWriter.cs:559, :564, :576, :580, :586`
- Modify: `ChartReader.cs:179, :182, :190`
- Modify: `ChartPatcher.cs:97, :100, :108`

**Interfaces:**
- Produces: `ChartAxisXml.Read(OpenXmlCompositeElement?, XLChartAxis)`,
  `ChartAxisXml.Apply(OpenXmlCompositeElement, XLChartAxis)`.

- [ ] **Step 1: Write the characterization test for the new/loaded agreement**

Add to `XLibur.Tests/Excel/Charts/ChartLegendAndAxisTests.cs`:

```csharp
/// <summary>
/// Gridlines assigned on a new chart and on a reloaded one must produce the same c:majorGridlines
/// state. Before spec 22 the two travelled through AppendAxisBody and PatchAxis independently.
/// </summary>
[Test]
[Arguments(true)]
[Arguments(false)]
public async Task Gridlines_agree_between_a_new_chart_and_a_reloaded_one(bool gridlines)
{
    using var ms = new MemoryStream();
    using (var wb = new XLWorkbook())
    {
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Q1";
        ws.Cell("B1").Value = 100;
        var chart = ws.Charts.Add(XLChartType.Bar);
        chart.SetSourceData(ws.Range("A1:B1"));
        wb.SaveAs(ms);
    }

    ms.Position = 0;
    using var reloaded = new XLWorkbook(ms);
    reloaded.Worksheet("Data").Charts.First().ValueAxis.MajorGridlines = gridlines;

    using var patched = new MemoryStream();
    reloaded.SaveAs(patched);

    var fromNew = ChartGoldenCorpus.CaptureChartPartXml(ws =>
    {
        var chart = ws.Charts.Add(XLChartType.Bar);
        chart.SetSourceData(ws.Range("A1:B2"));
        chart.ValueAxis.MajorGridlines = gridlines;
    });

    await Assert.That(HasMajorGridlines(patched)).IsEqualTo(gridlines);
    await Assert.That(fromNew.Contains("majorGridlines")).IsEqualTo(gridlines);
}

private static bool HasMajorGridlines(MemoryStream saved)
{
    saved.Position = 0;
    using var doc = SpreadsheetDocument.Open(saved, false);
    var part = doc.WorkbookPart!.WorksheetParts
        .SelectMany(p => p.DrawingsPart is null ? [] : p.DrawingsPart.ChartParts).First();
    return part.ChartSpace!.Descendants<C.MajorGridlines>().Any();
}
```

- [ ] **Step 2: Run it — expect PASS on current code**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/ChartLegendAndAxisTests/*"`
Expected: PASS.

- [ ] **Step 3: Build `ChartAxisXml.Apply` as the union of the three build-side helpers and `PatchAxis`**

`Apply` takes an axis element that may be freshly built (writer path: `c:axId`, `c:scaling`,
`c:delete`, `c:axPos` already appended) or loaded (patcher path: fully populated). Its contract:

- Write `c:scaling` if absent, patch it if present — absorbing `BuildScaling`.
- Write the optional body children in schema order — absorbing `AppendAxisBody`.
- Write the display units — absorbing `AppendAxisUnits`.
- Skip anything the model has not assigned, so an untouched axis round-trips unchanged.

Keep `ReadAxis` verbatim as `Read`.

- [ ] **Step 4: Repoint all eight call sites**

Writer: the two axis-building paths (`ChartWriter.cs:559-586`) each collapse to
`ChartAxisXml.Apply(axis, model);` after the required children are appended.
Reader: three `ChartAxisXml.Read(...)`. Patcher: three `ChartAxisXml.Apply(...)`.

- [ ] **Step 5: Run the chart suite and golden corpus**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/Chart*/*"`
Expected: PASS; `bar-secondary-axis` byte-identical.

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel/IO/Charts/ChartAxisXml.cs XLibur/Excel/IO/ChartFormatting.cs XLibur/Excel/IO/ChartWriter.cs XLibur/Excel/IO/ChartReader.cs XLibur/Excel/IO/ChartPatcher.cs XLibur.Tests/Excel/Charts/ChartLegendAndAxisTests.cs
git commit -m 'refactor(charts): give the axis one module with read and apply (spec 22 task 3)'
```

---

### Task 4 — `ChartDataLabelsXml`

**Files:**
- Create: `XLibur/Excel/IO/Charts/ChartDataLabelsXml.cs`
- Modify: `ChartFormatting.cs` — delete `BuildDataLabels`, `ReadDataLabels`, `SupportsDataLabels`,
  `PatchSeriesDataLabels`, `PatchGroupDataLabels`
- Modify: `ChartWriter.cs:978, :990`, `ChartReader.cs:154, :232`, `ChartPatcher.cs:127, :136`

**Interfaces:**
- Produces: `ChartDataLabelsXml.Supports(XLChartGroupKind)`,
  `ChartDataLabelsXml.Read(C.DataLabels?, XLChartDataLabels)`,
  `ChartDataLabelsXml.Apply(OpenXmlCompositeElement, XLChartDataLabels, XLChartType)`.

`PatchSeriesDataLabels` and `PatchGroupDataLabels` differ only in the parent element they attach to
and both delegate to the same body. `Apply` takes `OpenXmlCompositeElement parent`, which covers
both — two entry points become one.

- [ ] **Step 1: Write the characterization test**

Add to `XLibur.Tests/Excel/Charts/ChartDataLabelTests.cs`:

```csharp
/// <summary>
/// Series-level and group-level data labels reach the same c:dLbls body from the same model.
/// PatchSeriesDataLabels and PatchGroupDataLabels differed only in the parent they attached to.
/// </summary>
[Test]
public async Task Series_and_group_data_labels_write_the_same_body()
{
    var seriesLevel = ChartGoldenCorpus.CaptureChartPartXml(ws =>
    {
        var chart = ws.Charts.Add(XLChartType.Line);
        chart.SetSourceData(ws.Range("A1:B2"));
        chart.Series.First().DataLabels.ShowValue = true;
    });

    var groupLevel = ChartGoldenCorpus.CaptureChartPartXml(ws =>
    {
        var chart = ws.Charts.Add(XLChartType.Line);
        chart.SetSourceData(ws.Range("A1:B2"));
        chart.DataLabels.ShowValue = true;
    });

    await Assert.That(seriesLevel).Contains("<c:showVal val=\"1\"");
    await Assert.That(groupLevel).Contains("<c:showVal val=\"1\"");
}
```

- [ ] **Step 2: Run it — expect PASS on current code**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/ChartDataLabelTests/*"`
Expected: PASS.

- [ ] **Step 3: Move the five members, merging the two patch entry points into one `Apply`**

`PatchSeriesDataLabels` and `PatchGroupDataLabels` differ only in the parent they attach to. Their
shared body becomes `Apply`, typed on the common base:

```csharp
    /// <summary>
    /// Writes the assigned data-label properties into <paramref name="parent"/>, which is either a
    /// series element (c:ser) or a chart-group element (c:barChart and friends). The two used to be
    /// separate entry points that differed only in this parameter's type.
    /// </summary>
    internal static void Apply(OpenXmlCompositeElement parent, XLChartDataLabels labels,
        XLChartType chartType)
    {
        if (labels.AssignedFormat == XLChartDataLabelsFormat.None)
            return;

        var element = parent.Elements<C.DataLabels>().FirstOrDefault();
        if (element == null)
        {
            element = BuildBody(labels, chartType);
            if (element == null)
                return;
            ChartFormatting.InsertOrdered(parent, element, OrderFor(parent));
            return;
        }

        PatchBody(element, labels, chartType);
    }

    /// <summary>
    /// The schema child order of the element c:dLbls is being inserted into. A series and a chart
    /// group order their children differently, which is the only thing the two former patch entry
    /// points actually disagreed about.
    /// </summary>
    private static IReadOnlyList<Type> OrderFor(OpenXmlCompositeElement parent) =>
        parent is C.BarChartSeries or C.LineChartSeries or C.PieChartSeries or C.AreaChartSeries
            ? ChartFormatting.SeriesChildOrder
            : ChartFormatting.GroupChildOrder;
```

`BuildBody` is the former `BuildDataLabels` made private; `PatchBody` is the shared body the two
patch methods called; `OrderFor` is private to this module and stays private through task 6.

**Two things to verify against the code before writing this**, because they are what the merge turns
on: that `ChartFormatting` really does hold distinct series- and group-level order tables under some
name — if the two patch methods shared one table, `OrderFor` collapses to a constant and should be
deleted rather than written — and the exact set of series element types in the pattern above. Widen
`ChartFormatting.InsertOrdered` and both order tables to `internal` for now; task 6 moves them to
`ChartElementOrder` and sweeps the references.

- [ ] **Step 4: Repoint the six call sites**

`ChartWriter.cs:978` (series level) — replace:

```csharp
var dataLabels = ChartFormatting.BuildDataLabels(s.DataLabelsInternal, chartType);
if (dataLabels != null)
    seriesElement.Append(dataLabels);
```

with:

```csharp
ChartDataLabelsXml.Apply(seriesElement, s.DataLabelsInternal, chartType);
```

`ChartWriter.cs:990` (group level) — the same substitution against the group element and
`xlChart.DataLabelsInternal`.

`ChartReader.cs:154` and `:232` — `ChartFormatting.ReadDataLabels(...)` becomes
`ChartDataLabelsXml.Read(...)`, arguments unchanged.

`ChartPatcher.cs:127` — `ChartFormatting.SupportsDataLabels(group.Kind)` becomes
`ChartDataLabelsXml.Supports(group.Kind)`.

`ChartPatcher.cs:136` — replace:

```csharp
ChartFormatting.PatchGroupDataLabels(group.Element, xlChart.DataLabelsInternal, chartType);
```

with:

```csharp
ChartDataLabelsXml.Apply(group.Element, xlChart.DataLabelsInternal, chartType);
```

Add `using XLibur.Excel.IO.Charts;` to all three callers.

- [ ] **Step 5: Run the chart suite and golden corpus**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/Chart*/*"`
Expected: PASS; `line-datalabels` byte-identical.

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel/IO/Charts/ChartDataLabelsXml.cs XLibur/Excel/IO/ChartFormatting.cs XLibur/Excel/IO/ChartWriter.cs XLibur/Excel/IO/ChartReader.cs XLibur/Excel/IO/ChartPatcher.cs XLibur.Tests/Excel/Charts/ChartDataLabelTests.cs
git commit -m 'refactor(charts): give data labels one module with read and apply (spec 22 task 4)'
```

---

### Task 5 — `ChartSeriesFormatXml`

**Files:**
- Create: `XLibur/Excel/IO/Charts/ChartSeriesFormatXml.cs`
- Modify: `ChartFormatting.cs` — delete `BuildSeriesShapeProperties`, `BuildMarker`, `BuildSmooth`,
  `ReadSeriesFormat`, `PatchSeriesFormat`
- Modify: `ChartWriter.cs:952, :959, :966`, `ChartReader.cs:231`, `ChartPatcher.cs:176`

**Interfaces:**
- Consumes: after spec 16, the DrawingML setters live in `IO/DrawingML/ShapePropertiesWriter` —
  call them, do not reimplement.
- Produces: `ChartSeriesFormatXml.Read(C.BarChartSeries, XLChartSeries, bool)`,
  `ChartSeriesFormatXml.Apply(OpenXmlCompositeElement, XLChartSeries, XLChartGroupKind, XLChartType)`.

- [ ] **Step 1: Write the characterization test**

Add to `XLibur.Tests/Excel/Charts/ChartSeriesFormattingTests.cs`:

```csharp
/// <summary>
/// A fill assigned on a new chart and the same fill assigned on a reloaded one must produce the
/// same a:srgbClr. Before spec 22 these ran through BuildSeriesShapeProperties and
/// PatchSeriesFormat independently.
/// </summary>
[Test]
public async Task Series_fill_agrees_between_a_new_chart_and_a_reloaded_one()
{
    var fromNew = ChartGoldenCorpus.CaptureChartPartXml(ws =>
    {
        var chart = ws.Charts.Add(XLChartType.Bar);
        chart.SetSourceData(ws.Range("A1:B2"));
        chart.Series.First().Fill = XLColor.Red;
    });

    using var ms = new MemoryStream();
    using (var wb = new XLWorkbook())
    {
        var ws = wb.AddWorksheet("Data");
        ws.Cell("A1").Value = "Q1";
        ws.Cell("B1").Value = 100;
        var chart = ws.Charts.Add(XLChartType.Bar);
        chart.SetSourceData(ws.Range("A1:B1"));
        wb.SaveAs(ms);
    }

    ms.Position = 0;
    using var reloaded = new XLWorkbook(ms);
    reloaded.Worksheet("Data").Charts.First().Series.First().Fill = XLColor.Red;

    using var patched = new MemoryStream();
    reloaded.SaveAs(patched);

    await Assert.That(fromNew).Contains("FF0000");
    await Assert.That(SavedXml(patched)).Contains("FF0000");
}

private static string SavedXml(MemoryStream saved)
{
    saved.Position = 0;
    using var doc = SpreadsheetDocument.Open(saved, false);
    var part = doc.WorkbookPart!.WorksheetParts
        .SelectMany(p => p.DrawingsPart is null ? [] : p.DrawingsPart.ChartParts).First();
    using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
    using var reader = new StreamReader(stream);
    return reader.ReadToEnd();
}
```

- [ ] **Step 2: Run it — expect PASS on current code**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/ChartSeriesFormattingTests/*"`
Expected: PASS.

- [ ] **Step 3: Move the five members; fold the three builders into `Apply`'s absent-element branches**

`BuildMarker` and `BuildSmooth` return elements the writer appends; in `Apply` they become
"ensure the child exists with this value". Marker's `autoSymbol` and smooth's `smoothByChartType`
arguments are derived from `XLChartType`, which `Apply` already takes — compute them inside rather
than threading two more parameters through.

- [ ] **Step 4: Repoint the five call sites**

`ChartWriter.cs:952-966` — the three appends collapse into one call. Replace:

```csharp
var shapeProperties = ChartFormatting.BuildSeriesShapeProperties(s);
if (shapeProperties != null)
    seriesElement.Append(shapeProperties);

var marker = ChartFormatting.BuildMarker(s, autoSymbol);
if (marker != null)
    seriesElement.Append(marker);

var smooth = ChartFormatting.BuildSmooth(s, smoothByChartType);
if (smooth != null)
    seriesElement.Append(smooth);
```

with:

```csharp
ChartSeriesFormatXml.Apply(seriesElement, s, kind, chartType);
```

`autoSymbol` and `smoothByChartType` are derived from the chart type inside `Apply`, so the locals
that fed them become unused — delete them and whatever computed them, and let the
warnings-as-errors build confirm nothing else used them.

`ChartReader.cs:231` — replace:

```csharp
ChartFormatting.ReadSeriesFormat(seriesElement, series, useSecondaryAxis);
```

with:

```csharp
ChartSeriesFormatXml.Read(seriesElement, series, useSecondaryAxis);
```

`ChartPatcher.cs:176` — replace:

```csharp
ChartFormatting.PatchSeriesFormat(element, series.Items[i], kind, chartType);
```

with:

```csharp
ChartSeriesFormatXml.Apply(element, series.Items[i], kind, chartType);
```

Note the patcher's argument list is already `Apply`'s signature — that is the shape the writer is
being brought into line with, not a new invention.

- [ ] **Step 5: Run the full suite**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS; all five golden fixtures byte-identical.

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel/IO/Charts/ChartSeriesFormatXml.cs XLibur/Excel/IO/ChartFormatting.cs XLibur/Excel/IO/ChartWriter.cs XLibur/Excel/IO/ChartReader.cs XLibur/Excel/IO/ChartPatcher.cs XLibur.Tests/Excel/Charts/ChartSeriesFormattingTests.cs
git commit -m 'refactor(charts): give series formatting one module with read and apply (spec 22 task 5)'
```

---

### Task 6 — Delete `ChartFormatting`

After tasks 1–5, `ChartFormatting` holds only shared plumbing: `InsertOrdered`, the child-order
tables, and whatever thin adapters spec 16 left behind.

**Files:**
- Create: `XLibur/Excel/IO/Charts/ChartElementOrder.cs`
- Delete: `XLibur/Excel/IO/ChartFormatting.cs`
- Modify: the five concept modules to call `ChartElementOrder` instead of `ChartFormatting`

- [ ] **Step 1: Move `InsertOrdered` and the child-order tables into `ChartElementOrder`**

```csharp
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;

namespace XLibur.Excel.IO.Charts;

/// <summary>
/// Schema child ordering for chart elements. The OpenXML SDK does not order children, so every
/// insertion into a chart element goes through here.
/// </summary>
internal static class ChartElementOrder
{
    internal static readonly Type[] ChartChildOrder = { /* moved verbatim */ };
    internal static readonly Type[] LegendChildOrder = { /* moved verbatim */ };
    // Plus every other order table tasks 1–5 widened to internal — at minimum the axis, series
    // and chart-group tables. Enumerate them with:
    //   grep -n 'ChildOrder' XLibur/Excel/IO/ChartFormatting.cs

    internal static void InsertOrdered(OpenXmlCompositeElement parent, OpenXmlElement child,
        IReadOnlyList<Type> order)
    {
        // moved verbatim from ChartFormatting
    }
}
```

Copy the array contents and method body verbatim from `ChartFormatting` — this step must not change
behaviour. Then update the five concept modules to call `ChartElementOrder.InsertOrdered` and
`ChartElementOrder.<Table>` instead of the `ChartFormatting.*` names they used while it still
existed.

- [ ] **Step 2: Confirm `ChartFormatting` has no members left**

Run: `grep -n 'internal static' XLibur/Excel/IO/ChartFormatting.cs`
Expected: no output. If anything remains, it belongs to one of the five concepts — move it there,
not into `ChartElementOrder`.

- [ ] **Step 3: Delete the file and fix the build**

```bash
git rm XLibur/Excel/IO/ChartFormatting.cs
```

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Expected: no errors. Remove any now-unused `using` lines the compiler flags (warnings are errors).

- [ ] **Step 4: Run the full suite on all three frameworks**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Expected: PASS on net8.0 and net10.0; all five golden fixtures byte-identical.

- [ ] **Step 5: Commit**

```bash
git add -A
git commit -m 'refactor(charts): delete ChartFormatting; ordering moves to ChartElementOrder (spec 22 task 6)'
```

---

## Acceptance criteria

1. `XLibur/Excel/IO/ChartFormatting.cs` no longer exists.
2. Five concept modules under `XLibur/Excel/IO/Charts/`, each exposing at most three internal
   members, totalling **12 or fewer** entry points against today's 21.
3. `ChartWriter`, `ChartReader` and `ChartPatcher` between them make **12 or fewer** calls into the
   concept modules, against today's 29.
4. No `Build*` method survives: creating an element is a branch of `Apply`, never a separate entry
   point. Grep gate: `grep -rn 'internal static.*Build' XLibur/Excel/IO/Charts/` returns nothing.
5. Golden byte-identity: every fixture in task 0's corpus saves byte-identically after every task.
   Any diff is a finding to investigate, never noise to re-baseline without a written explanation.
6. No test modified except the five characterization tests added by tasks 1–5, and no existing
   assertion weakened.
7. No public API change: `PublicAPI.Unshipped.txt` untouched.
8. Full suite green on net8.0 and net10.0 after every task.

## Risks

- **The golden corpus may not cover a moved path.** Task 0 step 4 is the mitigation: prove the gate
  fails before trusting it. If a task moves code the corpus does not reach, widen the corpus in that
  task before moving the code.
- **`Apply` on a fresh element may order children differently from the append-based writer.**
  `InsertOrdered` should be equivalent, but the golden fixtures are what prove it. This is the single
  most likely source of a byte diff, and it is why task 1 is small and reviewed carefully first.
- **Spec 16 may reshape `ChartFormatting` more than expected**, invalidating the line numbers here.
  Re-read the file before starting; the concept grouping is stable even if the line numbers are not.
