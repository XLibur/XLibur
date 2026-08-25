# Spec 35 — Pivot table timelines

**Area:** Feature · Compat
**Effort:** M (~1 week, 4 PR-sized tasks)
**Dependencies:** Spec 16 (`DrawingAnchorFactory`, merged as #402) and PRD 5 tasks 1–3
(slicers, merged as #400/#403). Conflicts with any in-flight work in
`XLibur/Excel/IO/SlicerAnchorXml.cs` or `SlicerWriter.cs` — task 1 refactors both.
**Status:** Proposed. Implements PRD 5 task 4 (F10), the last outstanding piece of
`autoresearch/improve-260824-2056/prd-slicers-and-timelines.md`.

## Summary

A timeline is the date-scrubber Excel draws beside a pivot table: a horizontal band of years,
quarters, months or days with a draggable range. It is a slicer in every structural respect — six
package pieces, a cache that binds to a pivot cache, a graphic frame in the sheet's drawing — and
in exactly one respect it is not: its selection is a date range that Excel expresses as a
`dateBetween` filter on the pivot table, not as a list of item indices.

Timelines survive a round trip today because nothing opens their parts. Nothing else is true of
them: there is no model, no way to read one, no way to make one, and deleting the pivot table a
timeline filters leaves it pointing at nothing. `docs/round-trip-fidelity.md` names that last one
as the remaining instance of the dangling-reference hazard that slicers closed for themselves.

This spec models them, following the slicer precedent it inherits, and closes that row.

## What is already true — findings, verified 2026-08-25

Five things the PRD could not know, each of which changes the work.

**1. The fixture already exists.** `XLibur.Tests/Resource/TryToLoad/Timelines_Missing_21232.xlsx`
carries one pivot timeline on a `Date` field, and is already load-bearing in
`RoundTripFidelityTests.Timelines_and_their_caches_survive_a_round_trip` and
`LoadingTests.Can_load_and_save_preserves_timelines`. PRD 5 task 1 had to commission an
Excel-authored fixture for slicers; task 4 does not. The read model has a real witness on day one.

The timeline in it is **unfiltered** — `filterType="unknown"`, no `<selection>` — so it exercises
the bounds path and not the selection path. That is a gap the spec is honest about rather than one
it can close without Excel.

**2. The PRD's binding claim is wrong, in the cheap direction.** It says a timeline cache binds
through `x15:timelineCachePivotCacheDefinition`. It does not. The binding is
`x15:state/@pivotCacheId`, resolved against the *same* `x14:pivotCacheDefinition/@pivotCacheId`
extension slicers already use — the fixture's pivot cache carries
`<ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}"><x14:pivotCacheDefinition pivotCacheId="1"/></ext>`,
and the timeline cache quotes `pivotCacheId="1"`. `XLPivotCache.PivotCacheId`
(`XLPivotCache.cs:171`) is therefore already exactly the hook needed, already emitted on demand by
`SlicerCacheWriter`. There is no new pivot-cache work.

**3. A created timeline needs the date statistics XLibur already writes.** Excel decides a cache
field is date-shaped from `sharedItems/@containsDate`, `@containsNonDate`, `@minDate` and
`@maxDate`. `PivotTableCacheDefinitionPartWriter.cs:345–381` writes all four, and
`XLPivotCacheValues.Stats` (`XLPivotCacheValues.cs:76`) surfaces `ContainsDate`/`MinDate`/`MaxDate`
to the model. This is the single most important de-risking finding in the spec: the failure that
cost PRD 5 task 3 six rounds of manual Excel checks — Excel silently declining to attach a slicer
to a pivot table whose version stamps said it predated slicers — has no equivalent here that is not
already fixed. The fixture's pivot table is stamped `createdVersion="5" updatedVersion="5"`, which
is what `XLPivotTable` now defaults to since that fix. **Verify this claim in Excel before trusting
it** (see Risks); it is reasoning from a fixture, not from a rendered file.

**4. The bounds are computed, not stored anywhere else.** The fixture's cache field has
`minDate="1998-05-19"` / `maxDate="2004-02-06"`, and the timeline's `<bounds>` reads
`startDate="1998-01-01" endDate="2005-01-01"` — the field's range rounded outward to whole years.
A created timeline has to do that arithmetic; nothing hands it over.

**5. The SDK has typed classes for all of it.** `WorksheetPart.TimeLineParts`,
`WorkbookPart.TimeLineCacheParts`, `TimeLinePart.Timelines`,
`TimeLineCachePart.TimelineCacheDefinition`, and under
`DocumentFormat.OpenXml.Office2013.Excel`: `Timelines`, `Timeline`, `TimelineCacheDefinition`,
`TimelineCachePivotTables`, `TimelineCachePivotTable`, `TimelineState`, `BoundsTimelineRange`,
`SelectionTimelineRange`, `TimelineReferences`, `TimelineReference`, `TimelineCacheReferences`,
`TimelineCacheReference`. The structural parity with slicers is exact, including the trap: touching
`TimeLinePart.Timelines` attaches a DOM the SDK writes back over Excel's bytes.

`Timeline.Level` is a plain `UInt32Value` — the SDK defines no enumeration for it — so an unknown
value read from a file has to be preserved as a number, not narrowed to a modelled enum.

## The six pieces

Same count as a slicer, different URIs. All six are required; Excel offers to repair the file if
any is missing.

| # | Piece | Where |
|---|---|---|
| 1 | `x15:timelines` → `x15:timeline` | `xl/timelines/timelineN.xml`, one part per timeline |
| 2 | `x15:timelineCacheDefinition` | `xl/timelineCaches/timelineCacheN.xml` |
| 3 | Worksheet `extLst` ref | `ext uri="{7E03D99C-DC04-49d9-9315-930204A7B6E9}"` → `x15:timelineRefs`/`x15:timelineRef r:id` |
| 4 | Workbook `extLst` registration | `ext uri="{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}"` → `x15:timelineCacheRefs`/`x15:timelineCacheRef r:id` |
| 5 | `#N/A` defined name | named after the cache |
| 6 | Drawing frame | `a:graphicData uri="http://schemas.microsoft.com/office/drawing/2012/timeslicer"` holding `tsle:timeslicer name="…"` |

Relationship types: `http://schemas.microsoft.com/office/2011/relationships/timeline` (sheet →
timeline part) and `.../timelineCache` (workbook → cache part). Content types
`application/vnd.ms-excel.timeline+xml` and `application/vnd.ms-excel.timelineCache+xml`. The SDK's
part classes supply all of these; none needs hand-writing.

Verbatim from the fixture, as the shape to reproduce:

```xml
<!-- xl/timelines/timeline1.xml -->
<timelines xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
  <timeline name="Date" cache="ВстроеннаяВременнаяШкала_Date" caption="Date"
            level="2" selectionLevel="2" scrollPosition="2004-06-07T00:00:00"/>
</timelines>

<!-- xl/timelineCaches/timelineCache1.xml -->
<timelineCacheDefinition xmlns="…/2010/11/main" name="ВстроеннаяВременнаяШкала_Date" sourceName="Date">
  <pivotTables><pivotTable tabId="2" name="СводнаяТаблица2"/></pivotTables>
  <state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="1" filterType="unknown">
    <bounds startDate="1998-01-01T00:00:00" endDate="2005-01-01T00:00:00"/>
  </state>
</timelineCacheDefinition>
```

Note `minimalRefreshVersion`/`lastRefreshVersion` = 6, which is the timeline feature's own version
and unrelated to the pivot table's version stamps. A created timeline writes 6.

## Scope decisions

Taken with the project owner, 2026-08-25.

### The selection is read-only

F10 says "creation over a date field, **with range selection**". It ships without the second half,
deliberately, and this is the decision most likely to be revisited — so the reasoning is recorded
in full.

Writing a range is not writing one element. Excel expresses a timeline's selection three times over:

1. `x15:state/@filterType="dateBetween"` plus `@filterId`, `@filterTabId`, `@filterPivotName`, and
   an `x15:selection startDate endDate` child;
2. a `<filters>` entry on the *pivot table* definition — a `dateBetween` `pivotFilter` whose
   `autoFilter` carries two `customFilter` operands — which is what actually narrows the table;
3. `h="1"` on the pivot field items outside the range, which is what the rendered cells show.

Get any one of the three wrong and the workbook is internally inconsistent in a way no validator
sees. `XLPivotFilter` exists (`XLPivotFilter.cs`) but says of itself, accurately, "XLibur has no
API for creating or interpreting these yet, so the type exists to carry them through a load/save
unchanged". And there is no Excel-authored fixture of a *filtered* timeline anywhere in the repo to
check the output against.

This is the same judgement PRD 5 made for slicer selections — "not modelled rather than modelled
wrongly" — reached for a stronger reason, since a timeline's three-way consistency is harder than a
slicer's two-way. The model **reports** whatever selection a file carries; it does not set one.
Setting a range is its own change, with its own Excel-verified fixture, and it should not ride in
the PR that establishes the subsystem.

### The cascade is in scope

`docs/round-trip-fidelity.md:141` names timelines as "the remaining case of exactly that hazard":
deleting the pivot table a timeline filters leaves it pointing at nothing. Task 4 fixes it, by the
same route slicers took.

### The shared drawing code is extracted, not copied

`SlicerAnchorXml` (296 lines) and the `extLst`-reference plumbing in `SlicerWriter` differ from what
timelines need by a graphic URI, a child element name and a model type. Copying them creates two
implementations of the anchor rules to keep in step; spec 16 exists precisely because that had
already happened once between charts and pictures. They are extracted instead, guarded by the 1,300
lines of slicer tests that already exist.

## Design

### Public surface

```csharp
IXLWorksheet.Timelines          // IXLTimelines — owns: where a timeline is added
IXLPivotTable.Timelines         // IEnumerable<IXLTimeline> — a view, recomputed per access
```

The asymmetry is the slicer one and is deliberate. The worksheet owns, matching the file format, in
which a timeline belongs to the sheet that draws it. The pivot table's is a derived view over the
timelines whose cache names it, so it cannot drift out of step in the case a cached list would get
wrong: something being deleted.

```csharp
public interface IXLTimelines : IEnumerable<IXLTimeline>
{
    int Count { get; }
    IXLTimeline Timeline(string name);
    bool TryGetTimeline(string name, [NotNullWhen(true)] out IXLTimeline? timeline);

    /// <exception cref="ArgumentException">
    /// The pivot cache has no field of that name, or the field holds no dates.
    /// </exception>
    IXLTimeline Add(IXLPivotTable pivotTable, string dateFieldName);
}

public interface IXLTimeline
{
    string Name { get; }
    string Caption { get; set; }
    bool ShowHeader { get; set; }                  // x15:timeline/@showHeader
    bool ShowSelectionLabel { get; set; }
    bool ShowTimeLevel { get; set; }
    bool ShowHorizontalScrollbar { get; set; }
    string? Style { get; set; }                    // e.g. TimeSlicerStyleLight2
    XLTimelineLevel Level { get; set; }            // Years | Quarters | Months | Days
    IXLCell Position { get; set; }

    string SourceFieldName { get; }
    IXLWorksheet Worksheet { get; }
    IReadOnlyList<IXLPivotTable> PivotTables { get; }

    DateTime? BoundsStart { get; }                 // the scrubber's extent
    DateTime? BoundsEnd { get; }

    bool HasSelection { get; }
    DateTime? SelectionStart { get; }              // read-only — see Scope decisions
    DateTime? SelectionEnd { get; }
}
```

`Style` is a string and not an enumeration, for the reason `IXLSlicer.Style` is: a workbook may name
a custom style, and a read model that could only report the styles it knows about would silently
lose the rest.

`Level` is settable because it is pure presentation on the timeline element — changing it redraws
the band and does not touch the pivot table, unlike the selection. The **raw `uint` is stored** and
the enum is a projection over it, so a file carrying a level this build has never heard of
round-trips its number rather than being narrowed to the nearest modelled value. The setter takes
the enum; the patcher writes the stored number.

`BoundsStart`/`BoundsEnd` are read-only: they are the date extent of the underlying field, which
Excel recomputes on refresh, and a settable bound that Excel overwrites on the next refresh would be
honest in only one direction. They are **nullable** because `x15:bounds` is an optional child of
`x15:state` — a file that omits it reports nothing rather than a fabricated date. Every timeline
XLibur creates writes one.

**Naming.** Excel's cache name for a timeline is `NativeTimeline_<Field>` — the fixture shows the
Russian localisation of exactly that string. XLibur writes the English form, uniquified as
`NativeTimeline_Date`, then `NativeTimeline_Date1`, by the same rule and for the same reason
`XLSlicers.NextCacheName` documents: the name is not decoration, since the `#N/A` defined name is
written under it and must therefore be a legal defined name. The timeline's own `name` follows
`XLSlicers.NextSlicerName` — the field name, then `Date 1` — and is unique across the workbook.

### Model — `XLibur/Excel/Timelines/`

Seven files, mirroring `XLibur/Excel/Slicers/`'s eight — there is no timeline analogue of
`XLSlicerSourceKind`, since a timeline has only one kind of source.

| File | Job |
|---|---|
| `IXLTimeline.cs` | the public timeline |
| `IXLTimelines.cs` | the public collection |
| `XLTimelineLevel.cs` | `Years = 0, Quarters = 1, Months = 2, Days = 3` |
| `XLTimelineFormat.cs` | the assigned-flags enum |
| `XLTimeline.cs` | the model, with `SeedLoadedFormat` |
| `XLTimelines.cs` | the collection, `Add`, name allocation, `Remove` + `Removed` |
| `XLTimelineCache.cs` | binding and state |

`XLTimelineFormat` is the assigned-flags enum — `Caption`, `ShowHeader`, `ShowSelectionLabel`,
`ShowTimeLevel`, `ShowHorizontalScrollbar`, `Style`, `Level`, `Position` — and carries the same
contract `XLSlicerFormat` documents: `SeedLoadedFormat` is the **only** way the reader populates a
timeline, because assigning through the properties instead would mark every loaded timeline as
edited and bring parts nobody touched in for patching.

`XLTimelineCache` holds `Name`, `SourceName`, `PivotTableNames`/`PivotTables`, `PivotCache`,
`PivotCacheId`, `BoundsStart`/`BoundsEnd`, `SelectionStart`/`SelectionEnd`, `FilterType`, and the
`minimalRefreshVersion`/`lastRefreshVersion` pair. The last three are seeded from the file and
reproduced verbatim: a loaded cache's `filterType` and refresh versions are not XLibur's to
reinterpret.

### IO — `XLibur/Excel/IO/`

| Class | Job |
|---|---|
| `TimelineReader` | Loads timelines and caches, binds them to pivot tables. **Detached reads only** — see below. |
| `TimelineCacheWriter` | Two-pass: `PrepareTimelineCaches` before the workbook part, `WriteTimelineCaches` after the worksheets. |
| `TimelineWriter` | The worksheet half: one part per new timeline, plus the sheet's `extLst` reference. |
| `TimelinePatcher` | Applies assigned changes to a loaded timeline, in place. |
| `TimelineAnchorXml` | The `tsle:timeslicer` frame and its anchor. |

**The detached-read constraint is the whole fidelity guarantee and it is inherited, not
rediscovered.** Timeline parts survive today for one reason: nothing opens them.
`TimeLinePart.Timelines` and `TimeLineCachePart.TimelineCacheDefinition` are typed DOM properties —
reading through either attaches a tree the SDK tracks and re-serialises on the next save, taking
`mc:Ignorable`, `xr10:uid` and every unmodelled attribute with it. Every read goes through
`OpenXmlPartReader` and returns a detached tree, exactly as `SlicerReader.ReadDetached` does. A
byte-equality test sits behind it.

**The two-pass save ordering is also inherited, for the same two reasons.** The `#N/A` defined name
must be in the model before `WorkbookPartWriter` rebuilds the defined-name block, so cache parts are
*allocated* early; the cache content quotes `pivotCacheId`, which does not exist until the pivot
caches have been written, so content is *emitted* late. `XLWorkbook_Save.CreateParts` already has
both slots — `SlicerCacheWriter.PrepareSlicerCaches` at line 191, `WriteSlicerCaches` at line 220.

**One part per timeline, never appending into an existing one.** This is the defect-4 lesson from
PRD 5 stated as a rule: `SlicerWriter.EnsureSlicersPart` once reused a sheet's existing part and
appended to it, which opened a part Excel had written. Every automated gate passed and Excel
silently stopped drawing the slicer that was already there. Every timelines part Excel writes holds
exactly one `x15:timeline`; reproduce that.

### Shared extraction (task 1)

Two helpers, both in `XLibur/Excel/IO/DrawingML/`, both used by slicers and timelines:

```csharp
// DrawingFrameXml — a named graphic frame under a given graphic-data URI.
static OpenXmlCompositeElement Build(WorksheetDrawing drawing, DrawingFrameSpec spec);
static OpenXmlCompositeElement? Find(WorksheetDrawing drawing, string graphicUri, string name);
static void Move(DrawingsPart part, string graphicUri, string name, XLMarker target);
static void Remove(DrawingsPart? part, string graphicUri, string name);
static (XLMarker? From, XLMarker? To) ReadMarkers(OpenXmlCompositeElement anchor, XLWorksheet sheet);
```

`DrawingFrameSpec` carries the graphic URI, the child element's prefix / local name / namespace, and
the frame name. For a slicer that is `("…/2010/slicer", "sle", "slicer", …)`; for a timeline
`("…/2012/timeslicer", "tsle", "timeslicer", …)`. Everything else — the zero `xdr:xfrm` Excel
writes, `NextFrameId`, the `mc:AlternateContent`-tolerant search, the marker arithmetic — is
identical and moves once.

```csharp
// SheetExtensionRefs — an r:id list under a worksheet extLst URI.
static void Add(Worksheet sheet, XLWorksheetContentManager cm, string extUri,
                Func<OpenXmlCompositeElement> newList, string relId);
static void Remove(Worksheet sheet, XLWorksheetContentManager cm, string extUri, string relId);
```

`SlicerWriter.EnsureSlicerListReference` and `RemoveSlicerListReference` become calls into it. The
`WorksheetExtensionList` creation, the content-manager registration, and the "an empty registry is a
schema violation, not merely untidy" cleanup are the parts worth having once.

**The refactor preserves behaviour and is gated by the existing suite.** `SlicerPositionTests`,
`SlicerReadModelTests` and `SlicerWriteTests` (1,303 lines) must pass unchanged, and the
byte-equality assertions in them are what prove the extraction did not start opening parts it used
to leave alone.

### The A1 trap, restated

`DrawingAnchorFactory` silently anchors a drawing at A1 when handed no marker — no exception, no
missing element. For a picture that is a reasonable default; for a timeline it would drop the band
over the data it filters. Every timeline XLibur creates is given a marker in `XLTimelines.Add`
before the factory sees it, so the fallback stays unreachable, exactly as `XLSlicers` documents.

Placement default: two columns right of the pivot table's area, at its top row — the same rule
`XLSlicers.DefaultPositionBeside` uses, and it inherits the same known defect, that a pivot table
XLibur created from scratch reports a 1×1 placeholder area because XLibur does not lay pivot tables
out. Cosmetic; out of scope here; tracked with the slicer instance of it.

Size: Excel's timeline is wider and shorter than a slicer. The fixture's frame measures roughly
3,333,750 × 1,371,600 EMU ≈ 350 × 144 px. Use those as `DefaultWidthPx`/`DefaultHeightPx`.

### The frame is written bare, and that is a bet

Excel wraps the frame in `mc:AlternateContent` whose `mc:Choice Requires="tsle"` holds the real
frame and whose `mc:Fallback` holds a rectangle reading "Timeline: works in Excel 2013 or later".
XLibur writes the frame directly, for the two reasons `SlicerAnchorXml` records: the fallback serves
only readers that could not have drawn a timeline anyway, and `OpenXmlValidator` rejects
`mc:AlternateContent` as the content of a `xdr:oneCellAnchor`.

That reasoning is sound and, for **table** slicers, confirmed in Excel (PRD 5 criterion 3 passed).
For pivot slicers it is still unverified against merged code. It is a bet here too, and the fallback
if Excel declines to render is named in Risks.

### Hooks

| Site | Change |
|---|---|
| `XLWorkbook_Load.cs:178` | `TimelineReader.LoadTimelines(...)` after `SlicerReader.LoadSlicers` — binding needs pivot tables loaded. |
| `XLWorkbook_Save.cs:191` | `TimelineCacheWriter.PrepareTimelineCaches` beside the slicer call. |
| `XLWorkbook_Save.cs:220` | `TimelineCacheWriter.WriteTimelineCaches` beside the slicer call. |
| `WorksheetPartWriter.cs:220` | `TimelineWriter.WriteTimelines` after `SlicerWriter.WriteSlicers`. |
| `XLWorksheet.cs:209` | `public IXLTimelines Timelines => TimelinesInternal;` plus the lazy internal. |
| `XLPivotTable.cs:812` | `public IEnumerable<IXLTimeline> Timelines` — the view, mirroring `Slicers`. |
| `XLSlicerCascade.cs` | Renamed `XLPivotDependentCascade`; `OnPivotTableDeleted` gains the timeline arm. |

## Tasks

Four PRs. Task 1 is independent of the rest and can land first on its own; 2 → 3 → 4 are ordered.

### Task 1 — extract the shared drawing and extension plumbing

Pure refactor of shipped code, no new behaviour. `DrawingFrameXml` and `SheetExtensionRefs` land;
`SlicerAnchorXml` and `SlicerWriter` move onto them. Gate: the whole suite green with the slicer
tests **unchanged**, and `git diff` showing no change to any generated fixture.

If this task cannot be made behaviour-preserving — if a slicer byte-equality test moves — stop and
reconsider rather than adjusting the test. That test is the guarantee, not the obstacle.

### Task 2 — read model

`IXLTimeline`/`IXLTimelines`, the five model types, `TimelineReader`, and the two collection
properties. Nothing is written; nothing is created.

Tests: `TimelineReadModelTests` against `Timelines_Missing_21232.xlsx` — name, caption, level,
source field, bound pivot table, bounds, and `HasSelection == false`. Plus the byte-equality test
that proves reading did not attach a DOM: load, save, and assert `xl/timelines/timeline1.xml` and
`xl/timelineCaches/timelineCache1.xml` are byte-identical to the originals.

Update `docs/round-trip-fidelity.md` row 39: `❌` → `✅` in the *Modelled* column, and `n/a` → `✅` in
*Sheet references survive* — that cell is wrong today, since `LoadingTests` already asserts the
sheet's `timelineRef` survives.

### Task 3 — create, patch, remove

`XLTimelines.Add(pivotTable, dateFieldName)`, `TimelineWriter`, `TimelineCacheWriter`,
`TimelinePatcher`, `TimelineAnchorXml`, and `IXLTimeline.Position`.

`Add` validates through `cache.GetFieldValues(index).Stats.ContainsDate` and throws
`ArgumentException` for a field that holds no dates — a timeline over a text field is a repair
prompt, not a degraded timeline. Bounds come from `Stats.MinDate`/`MaxDate` rounded outward to whole
years, which is what the fixture shows Excel doing. `Level` starts at `Months`, matching the
fixture. `filterType` starts at `unknown` and the state carries no `<selection>`, which is a
timeline showing everything.

Tests: `TimelineWriteTests` — create → `OpenXmlValidator` (Office2013) clean → reload with the
timeline bound, captioned and positioned → **byte-equality asserted on every part the operation
should not have touched**. That last clause is the explicit lesson PRD 5 hands this task: three
byte-equality guards passed throughout a slicer feature that did not work, because each covered only
sheets where nothing had been added. Add a timeline to a sheet that *already has one* and assert the
existing timeline's part is untouched.

Also `TimelinePositionTests`, mirroring `SlicerPositionTests`.

### Task 4 — the cascade

`XLSlicerCascade` becomes `XLPivotDependentCascade`. Deleting a pivot table drops it from every
timeline cache that named it and removes any timeline left with nothing to filter — with its part,
the workbook registration, the `#N/A` defined name and the drawing anchor, the same five-piece
unpick `XLSlicers.Remove` documents.

Tests extend `PivotTableDeletionTests`. Update the Gaps section of `docs/round-trip-fidelity.md`,
which currently reads "Timelines are the remaining case of exactly that hazard".

## Acceptance criteria

1. A timeline in `Timelines_Missing_21232.xlsx` is readable: name, caption, level, source field,
   bound pivot table, and bounds.
2. Loading and saving that workbook leaves `xl/timelines/timeline1.xml` and
   `xl/timelineCaches/timelineCache1.xml` byte-identical.
3. A created timeline opens in Excel with no repair prompt, is drawn where it was positioned, and
   scrubbing it filters the pivot table.
4. Adding a second timeline to a sheet that already has one leaves the first one's part
   byte-identical and both drawn in Excel.
5. A loaded timeline with attributes XLibur does not model survives an unrelated edit and a save with
   those attributes intact.
6. `OpenXmlValidator` (Office2013) passes on every generated fixture.
7. Deleting a pivot table removes the timelines whose cache served only it, and the saved file opens
   clean.
8. The slicer suite passes unchanged across task 1.
9. Manual "opens clean in Excel" check recorded for criteria 3, 4 and 7, per repo convention — one
   `ac-check` workbook per criterion, **stamped with the commit sha in the filename**.

The Excel half of criteria 3, 4 and 7 cannot be met by the automated suite — criteria 4 and 7 have
an automated half as well, and passing it is not passing the criterion. PRD 5's central finding was
that every automated gate passed on a slicer feature that Excel refused to render; assume the same
exposure here and do not report task 3 complete on a green suite alone.

## Risks

| Risk | Evidence tier | Mitigation |
|---|---|---|
| Excel declines to render a bare frame with no `mc:AlternateContent` wrapper | primary — confirmed working for table slicers, still unverified for pivot slicers | If criterion 3 fails with everything else correct, write the frame inside `mc:AlternateContent` in a `xdr:twoCellAnchor` — the validator accepts that combination, and it is what the fixture uses |
| Task 1's extraction changes slicer behaviour | primary — it edits shipped code carrying a shipped fidelity guarantee | Slicer tests unchanged and green; any byte-equality movement stops the task |
| A created pivot table's date field is not date-shaped enough for Excel | secondary — the writer emits `containsDate`/`minDate`/`maxDate`, but this reasons from a fixture rather than from a rendered file | Criterion 3's manual check settles it; if it fails, compare the created cache field against the fixture's attribute for attribute before hypothesising elsewhere |
| Diagnosis wanders, as it did for PRD 5 criterion 2 | primary — six rounds, three of them lost to a negative result on the right-sounding attribute of the wrong element | Put control and suspect in **one workbook sharing one pivot cache** from the first round, not in separate files. That is what finally converged the slicer diagnosis |
| Reading opens a part and silently costs the fidelity guarantee | primary — this is what `SlicerReader` exists to prevent, and what `SlicerWriter` did anyway | Criterion 2's byte-equality test, plus criterion 4's, which covers the case criterion 2 does not |

## Deliberately out of scope

- **Setting a selection.** See Scope decisions. The read model reports one; nothing writes one.
- **Timeline styles beyond a name string.** Custom style *definitions* belong with the wider
  custom-style work, as F9 put it for slicers. A named style round-trips and is settable.
- **`Width`/`Height`.** A two-cell anchor carries no extent, so a size read back from an
  Excel-authored timeline would have to be derived from grid arithmetic. Deferred rather than exposed
  as a property honest in only one direction — the same call `IXLSlicer` made.
- **Timelines over anything but a pivot table.** The format has no other kind. There is no timeline
  analogue of the table slicer.
- **A bottom-right anchor marker.** As with slicers: the two-cell form Excel writes needs
  pixel-to-grid arithmetic that exists nowhere in XLibur, and a timeline that stretches when a column
  is widened is worse behaviour, not better.
- **Public `Remove` on `IXLTimelines`.** Removal exists internally because the cascade needs it.
  Matching `IXLSlicers`, it is not exposed until there is a reason.
