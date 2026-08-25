# Pivot Table Timelines Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Model Excel pivot-table timelines in XLibur — read them, create them, patch loaded ones in place, and remove them when the pivot table they filter is deleted.

**Architecture:** A timeline is structurally a slicer with different URIs, so this follows `XLibur/Excel/Slicers/` and `XLibur/Excel/IO/Slicer*.cs` one for one: the worksheet owns the collection, the pivot table exposes a recomputed view, every read streams its part **detached** so untouched parts survive byte-for-byte, and an edit is patched into the element the reader saw rather than regenerated. Task 1 first extracts the drawing-frame and worksheet-`extLst` plumbing that slicers already own so timelines reuse it instead of copying it.

**Tech Stack:** C# (net8.0/net9.0/net10.0, nullable enabled, `TreatWarningsAsErrors=true`), DocumentFormat.OpenXml 3.5.1, TUnit on Microsoft.Testing.Platform.

**Spec:** `docs/specs/35-pivot-timelines.md` — read it before Task 1. It carries the file-format reference, the verbatim fixture XML, and the reasoning behind every scope decision below.

## Global Constraints

- **Branch:** all work lands on `feat/pivot-timelines`. Never commit to `main`.
- **Nullable reference types are on and warnings are errors.** A build with any warning fails.
- **Assertions must be awaited.** `await Assert.That(x).IsEqualTo(y)`. A missing `await` means the assertion never runs and the test passes regardless.
- **Test filtering uses `--treenode-filter`, never `--filter`.** Always name the `.csproj`, never the `.slnx` — a solution-level filtered run exits 8 even when everything passes. Always pass `-f net10.0` or the suite runs twice.
- **Never use `sed -i` on a tracked file.** `.gitattributes` checks every source file out as CRLF; `sed -i` rewrites it as LF and turns a one-line change into a whole-file diff. Use the Edit/Write tools. Verify with `git diff --numstat` — a file whose changed-line count approaches its total line count has been rewritten, not edited.
- **No compound shell commands.** Each `git`/`dotnet` invocation is its own call; no `&&`, `||`, `;`.
- **Public API additions go in `XLibur/PublicAPI.Unshipped.txt`**, alphabetically, or the build fails.
- **The selection is read-only.** Nothing in this plan writes an `x15:selection`, a `dateBetween` pivot filter, or hidden item flags. See the spec's "The selection is read-only".
- **Namespace URIs, copied verbatim — a typo here produces a file Excel offers to repair, with no validator error:**
  - x15 main: `http://schemas.microsoft.com/office/spreadsheetml/2010/11/main`
  - worksheet timeline `extLst`: `{7E03D99C-DC04-49d9-9315-930204A7B6E9}`
  - workbook timeline-cache `extLst`: `{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}`
  - timeline graphic data: `http://schemas.microsoft.com/office/drawing/2012/timeslicer`
  - slicer graphic data (existing, task 1 moves it): `http://schemas.microsoft.com/office/drawing/2010/slicer`

## File Structure

**Created**

| File | Responsibility |
|---|---|
| `XLibur/Excel/IO/DrawingML/DrawingFrameSpec.cs` | Identifies one kind of named graphic frame: graphic URI plus the child element's prefix, local name and namespace. |
| `XLibur/Excel/IO/DrawingML/DrawingFrameXml.cs` | Build, find, move, remove and read the markers of a named graphic frame. Shared by slicers and timelines. |
| `XLibur/Excel/IO/DrawingML/SheetExtensionRefs.cs` | Ensure and prune an `r:id` list under a worksheet `extLst` URI. Shared by slicers and timelines. |
| `XLibur/Excel/Timelines/IXLTimeline.cs` | The public timeline. |
| `XLibur/Excel/Timelines/IXLTimelines.cs` | The public collection. |
| `XLibur/Excel/Timelines/XLTimelineLevel.cs` | `Years/Quarters/Months/Days`. |
| `XLibur/Excel/Timelines/XLTimelineFormat.cs` | Which properties the caller assigned. |
| `XLibur/Excel/Timelines/XLTimeline.cs` | The model. |
| `XLibur/Excel/Timelines/XLTimelines.cs` | The collection, `Add`, name allocation, removal bookkeeping. |
| `XLibur/Excel/Timelines/XLTimelineCache.cs` | Binding to pivot tables, bounds, selection, state. |
| `XLibur/Excel/IO/TimelineReader.cs` | Detached reads; binds caches to pivot tables. |
| `XLibur/Excel/IO/TimelineWriter.cs` | The worksheet half: one part per new timeline, plus the sheet `extLst` ref. |
| `XLibur/Excel/IO/TimelineCacheWriter.cs` | The workbook half: cache part, registration, `#N/A` defined name. |
| `XLibur/Excel/IO/TimelinePatcher.cs` | Applies assigned changes to a loaded timeline. |
| `XLibur/Excel/IO/TimelineAnchorXml.cs` | The `tsle:timeslicer` frame and its anchor. |
| `XLibur.Tests/Excel/Timelines/TimelineReadModelTests.cs` | Task 2's gate. |
| `XLibur.Tests/Excel/Timelines/TimelineWriteTests.cs` | Task 3's gate. |
| `XLibur.Tests/Excel/Timelines/TimelinePositionTests.cs` | Task 3's positioning gate. |

**Modified**

| File | Change |
|---|---|
| `XLibur/Excel/IO/SlicerAnchorXml.cs` | Delegates to `DrawingFrameXml` (task 1). |
| `XLibur/Excel/IO/SlicerWriter.cs` | Delegates to `SheetExtensionRefs` (task 1). |
| `XLibur/Excel/IXLWorksheet.cs` | `IXLTimelines Timelines { get; }` (task 2). |
| `XLibur/Excel/XLWorksheet.cs` | `Timelines` / `TimelinesInternal` (task 2). |
| `XLibur/Excel/PivotTables/IXLPivotTable.cs` | `IEnumerable<IXLTimeline> Timelines { get; }` (task 2). |
| `XLibur/Excel/PivotTables/XLPivotTable.cs` | The view (task 2). |
| `XLibur/Excel/XLWorkbook_Load.cs` | `TimelineReader.LoadTimelines` (task 2). |
| `XLibur/Excel/XLWorkbook_Save.cs` | Prepare/Write cache hooks (task 3). |
| `XLibur/Excel/IO/WorksheetPartWriter.cs` | `TimelineWriter.WriteTimelines` (task 3). |
| `XLibur/Excel/Slicers/XLSlicerCascade.cs` | Renamed to `XLPivotDependentCascade`, gains the timeline arm (task 4). |
| `XLibur/Excel/PivotTables/XLPivotTables.cs` | Two call sites follow the rename (task 4). |
| `XLibur.Tests/Excel/PivotTables/PivotTableDeletionTests.cs` | Cascade tests (task 4). |
| `XLibur/PublicAPI.Unshipped.txt` | Tasks 2 and 3. |
| `docs/round-trip-fidelity.md` | Tasks 2 and 4. |

---

## Task 1: Extract the shared drawing and extension plumbing

Pure refactor of shipped code. No new behaviour, no new tests. The existing slicer suite is the gate — it is a **characterization test**, so the cycle is "record green, refactor, prove still green" rather than red-green-refactor.

**Files:**
- Create: `XLibur/Excel/IO/DrawingML/DrawingFrameSpec.cs`
- Create: `XLibur/Excel/IO/DrawingML/DrawingFrameXml.cs`
- Create: `XLibur/Excel/IO/DrawingML/SheetExtensionRefs.cs`
- Modify: `XLibur/Excel/IO/SlicerAnchorXml.cs`
- Modify: `XLibur/Excel/IO/SlicerWriter.cs:161-205` (`EnsureSlicerListReference`) and `:243-268` (`RemoveSlicerListReference`)

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `readonly record struct DrawingFrameSpec(string GraphicUri, string Prefix, string LocalName, string ChildNamespace)`
  - `static Xdr.GraphicFrame DrawingFrameXml.BuildFrame(Xdr.WorksheetDrawing, in DrawingFrameSpec, string name)`
  - `static OpenXmlCompositeElement? DrawingFrameXml.FindAnchor(Xdr.WorksheetDrawing, in DrawingFrameSpec, string name)`
  - `static void DrawingFrameXml.MoveAnchor(DrawingsPart, in DrawingFrameSpec, string name, XLMarker target)`
  - `static void DrawingFrameXml.RemoveAnchor(DrawingsPart?, in DrawingFrameSpec, string name)`
  - `static (XLMarker? From, XLMarker? To) DrawingFrameXml.ReadMarkers(OpenXmlCompositeElement anchor, XLWorksheet)`
  - `static TList SheetExtensionRefs.EnsureList<TList>(Worksheet, XLWorksheetContentManager, string extensionUri, string namespacePrefix, string namespaceUri) where TList : OpenXmlCompositeElement, new()`
  - `static void SheetExtensionRefs.RemoveRefs<TList>(Worksheet, XLWorksheetContentManager, Predicate<OpenXmlElement> matches) where TList : OpenXmlCompositeElement`

- [ ] **Step 1: Record the baseline green**

```
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/SlicerReadModelTests/*"
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/SlicerWriteTests/*"
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/SlicerPositionTests/*"
```

Expected: all three PASS. Write the passing counts down — they must be identical at Step 10. If any is already failing, **stop**: the baseline is what this task is measured against.

- [ ] **Step 2: Create `DrawingFrameSpec`**

```csharp
namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// Identifies one kind of named graphic frame: a slicer's, a timeline's.
/// </summary>
/// <remarks>
/// Excel draws slicers and timelines through the same construct — a <c>xdr:graphicFrame</c> holding
/// a single element that carries nothing but the control's name. What separates them is the
/// <c>a:graphicData/@uri</c> and the name of that element, which is exactly what this carries.
/// <para>
/// <see cref="ChildNamespace"/> currently equals <see cref="GraphicUri"/> for both kinds Excel
/// defines. They are kept apart because nothing in the format requires them to agree, and a future
/// control that separates them would otherwise be silently mis-serialised.
/// </para>
/// </remarks>
/// <param name="GraphicUri">The <c>a:graphicData/@uri</c> Excel resolves the control from.</param>
/// <param name="Prefix">The namespace prefix of the frame's single child, e.g. <c>sle</c>.</param>
/// <param name="LocalName">The local name of that child, e.g. <c>slicer</c>.</param>
/// <param name="ChildNamespace">That child's namespace URI.</param>
internal readonly record struct DrawingFrameSpec(
    string GraphicUri,
    string Prefix,
    string LocalName,
    string ChildNamespace);
```

- [ ] **Step 3: Create `DrawingFrameXml`**

Move the bodies of `SlicerAnchorXml.BuildGraphicFrame`, `NextFrameId`, `FindAnchor`, `NameOfSlicer`, `ReadMarker`, `WriteMarker`, `ReadInt`, `ReadLong` and `EmuToPixels` here, parameterised by the spec. Keep every explanatory comment — they record why the code is shaped this way.

```csharp
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.Drawings;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// The graphic frame Excel draws a named control through, and the anchor that fixes it to the grid.
/// </summary>
/// <remarks>
/// Shared by slicers and timelines. Neither control is drawn by its own part: the part says what the
/// control filters, and a <c>xdr:graphicFrame</c> in the sheet's drawing says where it sits. Without
/// the frame the workbook opens and the control is simply invisible.
/// </remarks>
internal static class DrawingFrameXml
{
    /// <summary>
    /// The frame Excel recognises as the named control: a frame holding nothing but its name, under
    /// the spec's graphic-data URI.
    /// </summary>
    /// <remarks>
    /// Excel wraps this frame in <c>mc:AlternateContent</c>, whose <c>mc:Fallback</c> holds a
    /// rectangle explaining the control to a version of Excel too old to draw one. XLibur writes the
    /// frame directly, for two reasons. The wrapper protects nothing — the only readers the fallback
    /// serves are ones that could not have drawn the control anyway — and
    /// <c>OpenXmlValidator</c> rejects <c>mc:AlternateContent</c> as the content of a
    /// <c>xdr:oneCellAnchor</c>, which is the anchor form both controls use.
    /// </remarks>
    internal static Xdr.GraphicFrame BuildFrame(
        Xdr.WorksheetDrawing worksheetDrawing, in DrawingFrameSpec spec, string name)
    {
        // The SDK has no typed class for sle:slicer or tsle:timeslicer, so it is built as an unknown
        // element — which is also how it comes back when a file carrying one is read.
        var child = new OpenXmlUnknownElement(spec.Prefix, spec.LocalName, spec.ChildNamespace);
        child.SetAttribute(new OpenXmlAttribute(string.Empty, "name", string.Empty, name));
        child.AddNamespaceDeclaration(spec.Prefix, spec.ChildNamespace);

        return new Xdr.GraphicFrame(
            new Xdr.NonVisualGraphicFrameProperties(
                new Xdr.NonVisualDrawingProperties { Id = NextFrameId(worksheetDrawing), Name = name },
                new Xdr.NonVisualGraphicFrameDrawingProperties()),

            // Zero, as Excel writes it: the anchor decides where the frame goes and these are
            // ignored. The element is required all the same.
            new Xdr.Transform(
                new A.Offset { X = 0, Y = 0 },
                new A.Extents { Cx = 0, Cy = 0 }),

            new A.Graphic(new A.GraphicData(child) { Uri = spec.GraphicUri }))
        {
            Macro = string.Empty,
        };
    }

    /// <summary>
    /// The anchor holding the frame for the named control, in whichever of the three forms it uses.
    /// </summary>
    internal static OpenXmlCompositeElement? FindAnchor(
        Xdr.WorksheetDrawing worksheetDrawing, in DrawingFrameSpec spec, string name)
    {
        var graphicUri = spec.GraphicUri;
        var localName = spec.LocalName;

        foreach (var anchor in worksheetDrawing.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            if (anchor is not (Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor))
                continue;

            // The frame may be a direct child or, as Excel writes it, inside mc:AlternateContent.
            foreach (var graphicData in anchor.Descendants<A.GraphicData>())
            {
                if (graphicData.Uri?.Value != graphicUri)
                    continue;

                if (NameOfControl(graphicData, localName) == name)
                    return anchor;
            }
        }

        return null;
    }

    /// <summary>
    /// Moves a frame's anchor, shifting both corners by the same number of rows and columns so the
    /// control keeps the size it had.
    /// </summary>
    /// <remarks>
    /// The anchor is edited rather than replaced. Excel's own frame carries an
    /// <c>mc:AlternateContent</c> wrapper, a fallback shape and an <c>a16:creationId</c>, none of
    /// which XLibur models — replacing the anchor to move a control three columns would throw all of
    /// that away.
    /// </remarks>
    internal static void MoveAnchor(
        DrawingsPart drawingsPart, in DrawingFrameSpec spec, string name, XLMarker target)
    {
        var worksheetDrawing = drawingsPart.WorksheetDrawing;
        if (worksheetDrawing is null)
            return;

        var anchor = FindAnchor(worksheetDrawing, spec, name);
        var from = anchor?.GetFirstChild<Xdr.FromMarker>();
        if (anchor is null || from is null)
            return;

        // The delta is taken from what the file says rather than from the model's own old marker, so
        // a control moved twice before a save still lands where the caller last put it.
        var columnDelta = target.ColumnNumber - 1 - ReadInt(from.ColumnId);
        var rowDelta = target.RowNumber - 1 - ReadInt(from.RowId);

        WriteMarker(from, ReadInt(from.ColumnId) + columnDelta, ReadInt(from.RowId) + rowDelta);

        // A one-cell or absolute anchor has no bottom-right corner to keep in step.
        if (anchor.GetFirstChild<Xdr.ToMarker>() is { } to)
            WriteMarker(to, ReadInt(to.ColumnId) + columnDelta, ReadInt(to.RowId) + rowDelta);
    }

    /// <summary>
    /// Takes the anchored frame of a removed control out of the sheet's drawing.
    /// </summary>
    /// <remarks>
    /// The whole anchor goes, not just the frame inside it. An anchor is a position with one thing
    /// anchored at it, so an emptied one is a position for nothing — and where Excel wrapped the
    /// frame in <c>mc:AlternateContent</c>, the fallback shape would be left behind to be drawn in
    /// its place.
    /// </remarks>
    internal static void RemoveAnchor(DrawingsPart? drawingsPart, in DrawingFrameSpec spec, string name)
    {
        var worksheetDrawing = drawingsPart?.WorksheetDrawing;
        if (worksheetDrawing is null)
            return;

        FindAnchor(worksheetDrawing, spec, name)?.Remove();
    }

    /// <summary>
    /// The two corner markers of an anchor. Either may be absent: a one-cell anchor has no
    /// bottom-right corner, and an absolute anchor has neither.
    /// </summary>
    internal static (XLMarker? From, XLMarker? To) ReadMarkers(
        OpenXmlCompositeElement anchor, XLWorksheet worksheet)
    {
        var from = anchor.GetFirstChild<Xdr.FromMarker>() is { } f ? ReadMarker(worksheet, f) : null;
        var to = anchor.GetFirstChild<Xdr.ToMarker>() is { } t ? ReadMarker(worksheet, t) : null;
        return (from, to);
    }

    /// <summary>
    /// A drawing id no other drawing on the sheet is using. Ids are unique within the drawing part,
    /// not within the workbook.
    /// </summary>
    private static uint NextFrameId(Xdr.WorksheetDrawing worksheetDrawing)
    {
        var used = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().ToList();
        return used.Count > 0 ? used.Max(p => p.Id?.Value ?? 0U) + 1 : 1U;
    }

    /// <summary>
    /// The control name off a graphic frame's single child, which the SDK has no typed class for and
    /// so deserialises as an unknown element.
    /// </summary>
    private static string? NameOfControl(A.GraphicData graphicData, string localName)
    {
        foreach (var child in graphicData.ChildElements)
        {
            if (child.LocalName != localName)
                continue;

            var name = child.GetAttribute("name", string.Empty).Value;
            if (!string.IsNullOrEmpty(name))
                return name;
        }

        return null;
    }

    private static XLMarker ReadMarker(XLWorksheet worksheet, Xdr.MarkerType marker)
    {
        // Markers are written zero-based; the model counts from one.
        var column = ReadInt(marker.ColumnId) + 1;
        var row = ReadInt(marker.RowId) + 1;

        var cell = worksheet.Cell(
            row < 1 ? 1 : row,
            column < 1 ? 1 : column);

        return new XLMarker(cell, new System.Drawing.Point(
            EmuToPixels(ReadLong(marker.ColumnOffset), worksheet.Workbook.DpiX),
            EmuToPixels(ReadLong(marker.RowOffset), worksheet.Workbook.DpiY)));
    }

    private static void WriteMarker(Xdr.MarkerType marker, int column, int row)
    {
        marker.ColumnId = new Xdr.ColumnId((column < 0 ? 0 : column).ToInvariantString());
        marker.RowId = new Xdr.RowId((row < 0 ? 0 : row).ToInvariantString());
    }

    private static int ReadInt(OpenXmlLeafTextElement? element) =>
        int.TryParse(element?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : 0;

    private static long ReadLong(OpenXmlLeafTextElement? element) =>
        long.TryParse(element?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : 0;

    /// <summary>
    /// The inverse of <see cref="DrawingUnits.PixelsToEmu"/>, for reporting an offset a file carries.
    /// </summary>
    private static int EmuToPixels(long emu, double resolution) =>
        emu == 0 ? 0 : (int)System.Math.Round(emu * resolution / 914400d);
}
```

- [ ] **Step 4: Rewrite `SlicerAnchorXml` to delegate**

Replace the whole body below the class declaration. The class keeps its four entry points and its remarks; everything else moves out.

```csharp
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

internal static class SlicerAnchorXml
{
    /// <summary>The graphic data URI Excel uses for a slicer frame, for both kinds of slicer.</summary>
    private const string SlicerGraphicUri = "http://schemas.microsoft.com/office/drawing/2010/slicer";

    private static readonly DrawingFrameSpec Spec =
        new(SlicerGraphicUri, "sle", "slicer", SlicerGraphicUri);

    internal static void Append(Xdr.WorksheetDrawing worksheetDrawing, XLSlicer xlSlicer)
    {
        var frame = DrawingFrameXml.BuildFrame(worksheetDrawing, Spec, xlSlicer.Name);

        var anchor = DrawingAnchorFactory.Create(
            XLPicturePlacement.Move,
            new DrawingAnchorGeometry
            {
                Worksheet = xlSlicer.Worksheet,
                LeftPx = 0,
                TopPx = 0,
                WidthPx = xlSlicer.WidthPx,
                HeightPx = xlSlicer.HeightPx,

                // Never null by the time this runs; see the remarks on the A1 fallback.
                FromMarker = xlSlicer.FromMarker,
                ToMarker = xlSlicer.ToMarker,
            },
            frame);

        worksheetDrawing.Append(anchor);
    }

    internal static void Move(DrawingsPart drawingsPart, XLSlicer xlSlicer)
    {
        if (xlSlicer.FromMarker is not { } target)
            return;

        DrawingFrameXml.MoveAnchor(drawingsPart, Spec, xlSlicer.Name, target);
    }

    internal static void Remove(DrawingsPart? drawingsPart, XLSlicer xlSlicer) =>
        DrawingFrameXml.RemoveAnchor(drawingsPart, Spec, xlSlicer.Name);

    internal static void ReadPositions(DrawingsPart? drawingsPart, XLSlicers slicers)
    {
        var worksheetDrawing = drawingsPart?.WorksheetDrawing;
        if (worksheetDrawing is null)
            return;

        foreach (var slicer in slicers.Items)
        {
            var anchor = DrawingFrameXml.FindAnchor(worksheetDrawing, Spec, slicer.Name);
            if (anchor is null)
                continue;

            var (from, to) = DrawingFrameXml.ReadMarkers(anchor, (XLWorksheet)slicer.Worksheet);
            if (from is not null)
                slicer.FromMarker = from;

            if (to is not null)
                slicer.ToMarker = to;
        }
    }
}
```

Keep the original file's class-level `<remarks>` block verbatim — it documents the A1 trap and the `mc:AlternateContent` decision, and both still apply.

- [ ] **Step 5: Build and run the slicer suite**

```
dotnet build XLibur/XLibur.csproj -c Release --no-restore -v q
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/SlicerPositionTests/*"
```

Expected: build clean, tests PASS with the same count as Step 1. If a byte-equality assertion in `SlicerWriteTests` moves at any point in this task, **stop and revert** — that test is the guarantee, not the obstacle.

- [ ] **Step 6: Create `SheetExtensionRefs`**

```csharp
using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// The list of relationship ids a worksheet keeps under one <c>extLst</c> URI — the sheet's half of
/// a slicer or a timeline.
/// </summary>
/// <remarks>
/// A surviving control part is worthless if the sheet stops referencing it, and the worksheet part
/// is rebuilt from the model on every save while the control's part is not. Both callers need the
/// same three things: create the extension list if the sheet has none, register it with the content
/// manager so it lands in schema order, and prune an emptied registry — which is a schema violation
/// rather than merely untidy.
/// </remarks>
internal static class SheetExtensionRefs
{
    /// <summary>
    /// The reference list under the given extension URI, creating the extension list, the extension
    /// and the list itself if any is missing.
    /// </summary>
    internal static TList EnsureList<TList>(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        string extensionUri,
        string namespacePrefix,
        string namespaceUri)
        where TList : OpenXmlCompositeElement, new()
    {
        var extension = FindExtension(worksheet, extensionUri);

        if (extension is null)
        {
            if (!worksheet.Elements<WorksheetExtensionList>().Any())
            {
                var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.WorksheetExtensionList);
                worksheet.InsertAfter(new WorksheetExtensionList(), previousElement);
            }

            var extensionList = worksheet.Elements<WorksheetExtensionList>().First();
            cm.SetElement(XLWorksheetContents.WorksheetExtensionList, extensionList);

            extension = new WorksheetExtension { Uri = extensionUri };
            extension.AddNamespaceDeclaration(namespacePrefix, namespaceUri);
            extension.AppendChild(new TList());
            extensionList.AppendChild(extension);
        }

        var list = extension.GetFirstChild<TList>();
        if (list is null)
        {
            list = new TList();
            extension.AppendChild(list);
        }

        return list;
    }

    /// <summary>
    /// Drops every reference the predicate matches from every list of that type on the sheet, then
    /// prunes the extension and the extension list if either is left empty.
    /// </summary>
    /// <remarks>
    /// Every extension is scanned rather than one named URI, because a sheet may carry more than one
    /// list of the same type under different URIs — a sheet with both a pivot slicer and a table
    /// slicer has exactly that — and the caller knows only the relationship id.
    /// </remarks>
    internal static void RemoveRefs<TList>(
        Worksheet worksheet, XLWorksheetContentManager cm, Predicate<OpenXmlElement> matches)
        where TList : OpenXmlCompositeElement
    {
        var extensionList = worksheet.Elements<WorksheetExtensionList>().FirstOrDefault();
        if (extensionList is null)
            return;

        foreach (var extension in extensionList.Elements<WorksheetExtension>().ToList())
        {
            var list = extension.GetFirstChild<TList>();
            if (list is null)
                continue;

            foreach (var reference in list.ChildElements.ToList())
            {
                if (matches(reference))
                    reference.Remove();
            }

            if (!list.HasChildren)
                extension.Remove();
        }

        if (!extensionList.HasChildren)
        {
            worksheet.RemoveChild(extensionList);
            cm.SetElement(XLWorksheetContents.WorksheetExtensionList, null);
        }
    }

    private static WorksheetExtension? FindExtension(Worksheet worksheet, string uri) =>
        worksheet.Elements<WorksheetExtensionList>()
            .FirstOrDefault()?
            .Elements<WorksheetExtension>()
            .FirstOrDefault(e => string.Equals(e.Uri?.Value, uri, StringComparison.OrdinalIgnoreCase));
}
```

- [ ] **Step 7: Point `SlicerWriter` at it**

Replace `EnsureSlicerListReference` and `RemoveSlicerListReference`, and delete the now-unused private `FindExtension`.

```csharp
    private static void EnsureSlicerListReference(
        Worksheet worksheet, XLWorksheetContentManager cm, XLSlicerSourceKind kind, string relId)
    {
        var slicerList = SheetExtensionRefs.EnsureList<X14.SlicerList>(
            worksheet, cm, ExtensionUri(kind), "x14", X14Main2009SsNs);

        if (!slicerList.Elements<X14.SlicerRef>().Any(r => r.Id?.Value == relId))
            slicerList.AppendChild(new X14.SlicerRef { Id = relId });
    }

    private static void RemoveSlicerListReference(
        Worksheet worksheet, XLWorksheetContentManager cm, string relId) =>
        SheetExtensionRefs.RemoveRefs<X14.SlicerList>(
            worksheet, cm, r => r is X14.SlicerRef reference && reference.Id?.Value == relId);
```

Add `using XLibur.Excel.IO.DrawingML;` to the file's usings if it is not already there.

- [ ] **Step 8: Build**

```
dotnet build XLibur/XLibur.csproj -c Release --no-restore -v q
```

Expected: clean. Only pre-existing Polyfill warnings are normal.

- [ ] **Step 9: Confirm no file was rewritten**

```
git diff --numstat
```

Expected: `SlicerAnchorXml.cs` shrinks substantially and `SlicerWriter.cs` changes by a few dozen lines. If any file's changed-line count is close to its total, it was rewritten with the wrong line endings — revert and redo with the Edit tool.

- [ ] **Step 10: Run the whole suite**

```
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0
```

Expected: PASS, with the slicer counts from Step 1 unchanged and no test edited. This is the task's gate.

- [ ] **Step 11: Commit**

```
git add XLibur/Excel/IO/DrawingML/DrawingFrameSpec.cs XLibur/Excel/IO/DrawingML/DrawingFrameXml.cs XLibur/Excel/IO/DrawingML/SheetExtensionRefs.cs XLibur/Excel/IO/SlicerAnchorXml.cs XLibur/Excel/IO/SlicerWriter.cs
git commit -m 'refactor(drawingml): share the graphic-frame and sheet extLst plumbing

Slicers own two pieces of machinery a timeline needs unchanged: the named
graphic frame Excel draws a control through, and the r:id list a sheet
keeps under an extLst URI. Both are extracted so timelines reuse them
rather than growing a second copy to keep in step.

Behaviour-preserving. The slicer suite passes unchanged, including its
byte-equality assertions, which are what prove the extraction did not
start opening parts it used to leave alone.'
```

---

## Task 2: Read model

**Files:**
- Create: `XLibur/Excel/Timelines/IXLTimeline.cs`, `IXLTimelines.cs`, `XLTimelineLevel.cs`, `XLTimelineFormat.cs`, `XLTimeline.cs`, `XLTimelines.cs`, `XLTimelineCache.cs`
- Create: `XLibur/Excel/IO/TimelineReader.cs`
- Create: `XLibur/Excel/IO/TimelineAnchorXml.cs` (read side only; `Append` arrives in task 3)
- Modify: `XLibur/Excel/IXLWorksheet.cs:456` (beside `Slicers`), `XLibur/Excel/XLWorksheet.cs:209`, `XLibur/Excel/PivotTables/IXLPivotTable.cs:206`, `XLibur/Excel/PivotTables/XLPivotTable.cs:812`, `XLibur/Excel/XLWorkbook_Load.cs:178`, `XLibur/PublicAPI.Unshipped.txt`, `docs/round-trip-fidelity.md:39`
- Test: `XLibur.Tests/Excel/Timelines/TimelineReadModelTests.cs`

**Interfaces:**
- Consumes: `DrawingFrameXml.FindAnchor`, `DrawingFrameXml.ReadMarkers`, `DrawingFrameSpec` (task 1).
- Produces:
  - `public enum XLTimelineLevel { Years = 0, Quarters = 1, Months = 2, Days = 3 }`
  - `public interface IXLTimeline` with `Name`, `Caption`, `ShowHeader`, `ShowSelectionLabel`, `ShowTimeLevel`, `ShowHorizontalScrollbar`, `Style`, `Level`, `Position`, `SourceFieldName`, `Worksheet`, `PivotTables`, `BoundsStart`, `BoundsEnd`, `HasSelection`, `SelectionStart`, `SelectionEnd`
  - `public interface IXLTimelines : IEnumerable<IXLTimeline>` with `Count`, `Timeline(string)`, `TryGetTimeline(string, out IXLTimeline?)` — **`Add` arrives in task 3**, so no public method exists before there is a writer behind it
  - `internal sealed class XLTimelineCache` — `Name`, `SourceName`, `IsNew`, `WorkbookRelId`, `PivotCacheId`, `PivotTableNames`, `PivotTables`, `PivotCache`, `BoundsStart/End`, `SelectionStart/End`, `FilterType` (string), `MinimalRefreshVersion`, `LastRefreshVersion`
  - `internal sealed class XLTimeline` — plus `PartRelId`, `IsNew`, `AssignedFormat`, `LevelRaw`, `SelectionLevelRaw`, `ScrollPosition`, `FromMarker`, `ToMarker`, `WidthPx`, `HeightPx`, `SeedLoadedFormat(...)`
  - `internal sealed class XLTimelines : IXLTimelines` — plus `Items`, `Add(XLTimeline)`, `Remove(XLTimeline)`, `Removed`
  - `static void TimelineReader.LoadTimelines(WorkbookPart, Sheets, XLWorksheets)`
  - `static void TimelineAnchorXml.ReadPositions(DrawingsPart?, XLTimelines)`
  - `XLWorksheet.TimelinesInternal`, `IXLWorksheet.Timelines`, `IXLPivotTable.Timelines`

- [ ] **Step 1: Write the failing read test**

Create `XLibur.Tests/Excel/Timelines/TimelineReadModelTests.cs`.

```csharp
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Timelines;

/// <summary>
/// The timeline read model, read against the Excel-authored fixture the round-trip suite already
/// uses.
/// </summary>
/// <remarks>
/// <c>Resource/TryToLoad/Timelines_Missing_21232.xlsx</c> carries one timeline on the <c>Date</c>
/// field of the pivot table on sheet <c>Pivot</c>. Its names are Russian, which is a feature rather
/// than an inconvenience: a reader that assumed an English cache-name convention would fail here.
/// The timeline is unfiltered — <c>filterType="unknown"</c> and no <c>x15:selection</c> — so it
/// exercises the bounds path and pins the "no selection" case.
/// </remarks>
public class TimelineReadModelTests
{
    private const string Fixture = @"TryToLoad\Timelines_Missing_21232.xlsx";

    [Test]
    public async Task The_worksheet_owns_the_timeline_drawn_on_it()
    {
        using var wb = Load();

        await Assert.That(wb.Worksheet("Pivot").Timelines.Count).IsEqualTo(1);
        await Assert.That(wb.Worksheet("Data").Timelines.Count).IsEqualTo(0);
    }

    [Test]
    public async Task A_timeline_reports_what_the_file_says()
    {
        using var wb = Load();

        var timeline = wb.Worksheet("Pivot").Timelines.Single();

        await Assert.That(timeline.Name).IsEqualTo("Date");
        await Assert.That(timeline.Caption).IsEqualTo("Date");
        await Assert.That(timeline.SourceFieldName).IsEqualTo("Date");
        await Assert.That(timeline.Level).IsEqualTo(XLTimelineLevel.Months);

        // Absent booleans default to true, which is what Excel means by omitting them.
        await Assert.That(timeline.ShowHeader).IsTrue();
        await Assert.That(timeline.ShowSelectionLabel).IsTrue();

        // The fixture writes no style attribute at all.
        await Assert.That(timeline.Style).IsNull();
    }

    [Test]
    public async Task A_timeline_binds_to_the_pivot_table_its_cache_names()
    {
        using var wb = Load();

        var timeline = wb.Worksheet("Pivot").Timelines.Single();

        await Assert.That(timeline.PivotTables.Select(pt => pt.Name))
            .IsEquivalentTo(new[] { "СводнаяТаблица2" });
        await Assert.That(timeline.Worksheet.Name).IsEqualTo("Pivot");
    }

    [Test]
    public async Task An_unfiltered_timeline_reports_its_bounds_and_no_selection()
    {
        using var wb = Load();

        var timeline = wb.Worksheet("Pivot").Timelines.Single();

        // The bounds are the date field's range rounded outward to whole years.
        await Assert.That(timeline.BoundsStart).IsEqualTo(new DateTime(1998, 1, 1));
        await Assert.That(timeline.BoundsEnd).IsEqualTo(new DateTime(2005, 1, 1));

        await Assert.That(timeline.HasSelection).IsFalse();
        await Assert.That(timeline.SelectionStart).IsNull();
        await Assert.That(timeline.SelectionEnd).IsNull();
    }

    [Test]
    public async Task A_timeline_reports_where_it_is_drawn()
    {
        using var wb = Load();

        // The fixture anchors the frame at xdr:col 2, xdr:row 1 — zero-based, so C2.
        var timeline = wb.Worksheet("Pivot").Timelines.Single();

        await Assert.That(timeline.Position.Address.ToString()).IsEqualTo("C2");
    }

    [Test]
    public async Task A_pivot_table_views_the_timelines_that_filter_it()
    {
        using var wb = Load();

        var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();

        await Assert.That(pivotTable.Timelines.Select(t => t.Name)).IsEquivalentTo(new[] { "Date" });
    }

    [Test]
    public async Task A_timeline_can_be_found_by_name()
    {
        using var wb = Load();

        var timelines = wb.Worksheet("Pivot").Timelines;

        await Assert.That(timelines.Timeline("Date").Caption).IsEqualTo("Date");
        await Assert.That(timelines.TryGetTimeline("Date", out var found)).IsTrue();
        await Assert.That(found!.Name).IsEqualTo("Date");
        await Assert.That(timelines.TryGetTimeline("Nope", out _)).IsFalse();
    }

    [Test]
    public async Task Reading_a_timeline_does_not_rewrite_its_parts()
    {
        // The regression gate for the whole read model. Timeline parts survive a round trip because
        // nothing opens them; reaching one through TimeLinePart.Timelines would attach a DOM the SDK
        // writes back over the original bytes on save, taking mc:Ignorable and every attribute
        // XLibur does not model with it. The reader streams the parts detached instead.
        using var original = Resource();
        var before = PartBytes(original, "xl/timelines/timeline1.xml");
        var beforeCache = PartBytes(original, "xl/timelineCaches/timelineCache1.xml");

        using var saved = LoadAndSave();

        await Assert.That(PartBytes(saved, "xl/timelines/timeline1.xml")).IsEquivalentTo(before);
        await Assert.That(PartBytes(saved, "xl/timelineCaches/timelineCache1.xml")).IsEquivalentTo(beforeCache);
    }

    [Test]
    public async Task Timelines_still_load_after_a_round_trip()
    {
        using var saved = LoadAndSave();
        saved.Position = 0;
        using var wb = new XLWorkbook(saved);

        await Assert.That(wb.Worksheet("Pivot").Timelines.Single().Level).IsEqualTo(XLTimelineLevel.Months);
    }

    #region Helpers

    /// <summary>
    /// The fixture, opened over a copy that outlives this call. The workbook reads its original
    /// stream again on save, so the stream cannot be disposed when this returns.
    /// </summary>
    private static XLWorkbook Load()
    {
        var stream = Resource();
        stream.Position = 0;
        return new XLWorkbook(stream);
    }

    private static MemoryStream Resource()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Fixture));
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms;
    }

    private static MemoryStream LoadAndSave()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Fixture));
        var ms = new MemoryStream();

        using (var wb = new XLWorkbook(stream))
            wb.SaveAs(ms);

        return ms;
    }

    private static byte[] PartBytes(MemoryStream package, string partPath)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.First(e =>
            e.FullName.Equals(partPath, StringComparison.OrdinalIgnoreCase));

        using var entryStream = entry.Open();
        using var buffer = new MemoryStream();
        entryStream.CopyTo(buffer);
        return buffer.ToArray();
    }

    #endregion
}
```

Before writing more code, confirm the two sheet names and the pivot table name against the fixture:

```
unzip -p XLibur.Tests/Resource/TryToLoad/Timelines_Missing_21232.xlsx xl/workbook.xml
```

The `<sheets>` element names them. If they differ from `Pivot`/`Data`, fix the test, not the fixture.

- [ ] **Step 2: Run the test to verify it fails**

```
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/TimelineReadModelTests/*"
```

Expected: compile error — `IXLWorksheet` has no `Timelines`. That is the failure; a compile error is a legitimate red for a type that does not exist yet.

- [ ] **Step 3: Create `XLTimelineLevel` and `XLTimelineFormat`**

```csharp
namespace XLibur.Excel;

/// <summary>
/// How finely a timeline's band is divided.
/// </summary>
/// <remarks>
/// The numbers are the values Excel writes to <c>x15:timeline/@level</c>, not an XLibur invention.
/// A file may carry a value outside this set; <see cref="IXLTimeline.Level"/> is a projection over
/// the raw number, which is preserved through a save either way.
/// </remarks>
public enum XLTimelineLevel
{
    Years = 0,
    Quarters = 1,
    Months = 2,
    Days = 3,
}
```

```csharp
using System;

namespace XLibur.Excel;

/// <summary>
/// Which of a timeline's properties the caller has actually assigned.
/// </summary>
/// <remarks>
/// The same device <see cref="XLSlicerFormat"/> uses, and for the same reason. A timeline read from
/// a file is never regenerated — that is what keeps the parts of its XML XLibur has no model for
/// intact — so an edit has to be patched into the element the reader saw. Seeding a value while
/// reading leaves these flags clear, which is how a timeline nobody touched is left alone entirely.
/// </remarks>
[Flags]
internal enum XLTimelineFormat
{
    None = 0,
    Caption = 1 << 0,
    ShowHeader = 1 << 1,
    ShowSelectionLabel = 1 << 2,
    ShowTimeLevel = 1 << 3,
    ShowHorizontalScrollbar = 1 << 4,
    Style = 1 << 5,
    Level = 1 << 6,

    /// <summary>
    /// The timeline has been moved. Unlike the others this is patched into the drawing part rather
    /// than the timelines part, because that is where a timeline's anchor lives.
    /// </summary>
    Position = 1 << 7,
}
```

- [ ] **Step 4: Create `XLTimelineCache`**

```csharp
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace XLibur.Excel;

/// <summary>
/// A timeline cache: the workbook-level part that binds a timeline to the pivot tables it filters
/// and remembers its date range.
/// </summary>
/// <remarks>
/// <para>
/// Kept internal, as <see cref="XLSlicerCache"/> is. The cache is a level of indirection Excel's
/// file format needs and a caller does not; everything it holds is reached through
/// <see cref="IXLTimeline"/>.
/// </para>
/// <para>
/// The binding is <c>x15:state/@pivotCacheId</c>, which points at the identifier the pivot cache
/// carries in its own <c>x14:pivotCacheDefinition</c> extension — the same identifier slicer caches
/// quote. Caches are registered in <c>xl/workbook.xml</c>'s <c>extLst</c> under
/// <c>x15:timelineCacheRefs</c>, and Excel writes a <c>#N/A</c> defined name per cache.
/// </para>
/// </remarks>
[DebuggerDisplay("{Name} ({SourceName})")]
internal sealed class XLTimelineCache
{
    internal XLTimelineCache(string name, string sourceName)
    {
        Name = name;
        SourceName = sourceName;
    }

    /// <summary>
    /// The cache name, for example <c>NativeTimeline_Date</c>. A timeline refers to its cache by
    /// this name, and so does the <c>#N/A</c> defined name written for it.
    /// </summary>
    internal string Name { get; }

    /// <summary>The pivot cache field the timeline scrubs.</summary>
    internal string SourceName { get; }

    /// <summary>Whether the cache was created through the API rather than read from a package.</summary>
    internal bool IsNew { get; set; }

    /// <summary>
    /// The id of the cache's part relationship on the workbook part, for finding the part again on
    /// save. Null for a cache that has never been in a package.
    /// </summary>
    internal string? WorkbookRelId { get; set; }

    /// <summary>
    /// The <c>x14:pivotCacheDefinition/@pivotCacheId</c> of the pivot cache this cache scrubs.
    /// </summary>
    internal uint? PivotCacheId { get; set; }

    /// <summary>
    /// The names of the pivot tables the cache drives, as written in the part. A name with no
    /// matching pivot table in the workbook is left out of <see cref="PivotTables"/>.
    /// </summary>
    internal List<string> PivotTableNames { get; } = [];

    /// <summary>The pivot tables resolved from <see cref="PivotTableNames"/>.</summary>
    internal List<XLPivotTable> PivotTables { get; } = [];

    /// <summary>The pivot cache behind those pivot tables.</summary>
    internal XLPivotCache? PivotCache { get; set; }

    /// <summary>
    /// The extent of the scrubber. Nullable because <c>x15:bounds</c> is an optional child of
    /// <c>x15:state</c>; a file that omits it reports nothing rather than a fabricated date.
    /// </summary>
    internal DateTime? BoundsStart { get; set; }

    /// <inheritdoc cref="BoundsStart"/>
    internal DateTime? BoundsEnd { get; set; }

    /// <summary>
    /// The selected range, when the file records one. Read-only in the model: changing it has to
    /// move a <c>dateBetween</c> pivot filter and the pivot field's item visibility with it, and a
    /// timeline whose range disagrees with its pivot table is a broken workbook.
    /// </summary>
    internal DateTime? SelectionStart { get; set; }

    /// <inheritdoc cref="SelectionStart"/>
    internal DateTime? SelectionEnd { get; set; }

    /// <summary>
    /// The raw <c>x15:state/@filterType</c> token, held as a string rather than as the SDK's
    /// <c>EnumValue&lt;PivotFilterValues&gt;</c>.
    /// </summary>
    /// <remarks>
    /// The SDK types this attribute as an enumeration, but an <c>EnumValue</c> preserves an
    /// unrecognised token in its <c>InnerText</c>. Carrying the string is what lets a file written
    /// by a newer Excel round-trip a filter type this build has never heard of. A timeline with no
    /// selection says <c>unknown</c>, which is what a created one starts at.
    /// </remarks>
    internal string FilterType { get; set; } = "unknown";

    /// <summary>
    /// The timeline feature's own version stamps. 6 is what Excel 2013 and later write, and it is
    /// unrelated to the pivot table's <c>createdVersion</c>.
    /// </summary>
    internal uint MinimalRefreshVersion { get; set; } = 6;

    /// <inheritdoc cref="MinimalRefreshVersion"/>
    internal uint LastRefreshVersion { get; set; } = 6;
}
```

- [ ] **Step 5: Create `IXLTimeline` and `IXLTimelines`**

```csharp
using System;
using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// A timeline: the date scrubber Excel draws on a worksheet to filter a pivot table by date.
/// </summary>
/// <remarks>
/// <para>
/// A timeline is owned by the worksheet it is drawn on — see <see cref="IXLWorksheet.Timelines"/>.
/// What it filters is a separate relationship, held by its cache: <see cref="IXLPivotTable.Timelines"/>
/// is a view over the timelines whose cache lists that pivot table.
/// </para>
/// <para>
/// A timeline read from a file is reported here in full, including attributes XLibur has no model
/// for. Editing one patches the change into the part it was read from rather than regenerating it,
/// so everything alongside the edited attribute survives; a timeline nobody assigns to is not
/// written to at all. The selection is read-only — changing it has to move the pivot table's
/// <c>dateBetween</c> filter and item visibility with it, and that is not modelled.
/// </para>
/// </remarks>
public interface IXLTimeline
{
    /// <summary>
    /// The timeline's internal name, unique within the workbook. This is what the drawing anchor
    /// refers to, not what the user sees — see <see cref="Caption"/> for that.
    /// </summary>
    string Name { get; }

    /// <summary>The heading shown above the band. Defaults to <see cref="Name"/>.</summary>
    string Caption { get; set; }

    /// <summary>Whether <see cref="Caption"/> is displayed. <c>true</c> unless the file says otherwise.</summary>
    bool ShowHeader { get; set; }

    /// <summary>Whether the selected range is written out under the header.</summary>
    bool ShowSelectionLabel { get; set; }

    /// <summary>Whether the level chooser (Years / Quarters / Months / Days) is shown.</summary>
    bool ShowTimeLevel { get; set; }

    /// <summary>Whether the scrollbar under the band is shown.</summary>
    bool ShowHorizontalScrollbar { get; set; }

    /// <summary>
    /// The name of the timeline style, for example <c>TimeSlicerStyleLight2</c>. <c>null</c> means
    /// the workbook default.
    /// </summary>
    /// <remarks>
    /// Deliberately a string rather than an enumeration of the built-in styles: a workbook may name
    /// a custom style, and a read model that could only report the styles it knows about would
    /// silently lose the rest.
    /// </remarks>
    string? Style { get; set; }

    /// <summary>How finely the band is divided.</summary>
    XLTimelineLevel Level { get; set; }

    /// <summary>The cell the timeline's top-left corner is anchored to. Setting it moves the timeline.</summary>
    /// <remarks>
    /// Moving a timeline read from a file shifts both of its corners together, so it keeps the size
    /// it had. Reading reports the cell the corner sits in; a file may also place the corner some
    /// distance into that cell, and that offset is preserved through a save but is not reported here.
    /// </remarks>
    IXLCell Position { get; set; }

    /// <summary>The pivot cache field the band is drawn from.</summary>
    string SourceFieldName { get; }

    /// <summary>The worksheet the timeline is drawn on.</summary>
    IXLWorksheet Worksheet { get; }

    /// <summary>
    /// The pivot tables this timeline filters. A pivot table listed in the cache but missing from
    /// the workbook is omitted rather than reported as null.
    /// </summary>
    IReadOnlyList<IXLPivotTable> PivotTables { get; }

    /// <summary>
    /// The extent of the scrubber — the date field's range, rounded outward. <c>null</c> when the
    /// file records no bounds.
    /// </summary>
    /// <remarks>
    /// Read-only: Excel recomputes the extent when the pivot cache refreshes, so a settable bound
    /// would be honest in only one direction.
    /// </remarks>
    DateTime? BoundsStart { get; }

    /// <inheritdoc cref="BoundsStart"/>
    DateTime? BoundsEnd { get; }

    /// <summary>
    /// Whether the timeline records an explicit range. <c>false</c> means every date is showing,
    /// which is how Excel represents a timeline nobody has scrubbed.
    /// </summary>
    bool HasSelection { get; }

    /// <summary>
    /// The first date of the selected range, or <c>null</c> when <see cref="HasSelection"/> is
    /// <c>false</c>.
    /// </summary>
    /// <remarks>
    /// Read-only. Excel records a timeline's range in three places at once — the cache's state, a
    /// <c>dateBetween</c> filter on the pivot table, and hidden-item flags on the pivot field — and
    /// a model that wrote one without the others would produce a workbook that disagrees with
    /// itself in a way no validator can see.
    /// </remarks>
    DateTime? SelectionStart { get; }

    /// <inheritdoc cref="SelectionStart"/>
    DateTime? SelectionEnd { get; }
}
```

```csharp
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace XLibur.Excel;

/// <summary>
/// The timelines drawn on a worksheet.
/// </summary>
/// <remarks>
/// The worksheet owns its timelines: this is the collection a timeline is added to and removed from.
/// What a timeline filters is a separate relationship held by its cache, exposed as a view on
/// <see cref="IXLPivotTable.Timelines"/>.
/// </remarks>
public interface IXLTimelines : IEnumerable<IXLTimeline>
{
    /// <summary>The number of timelines on the worksheet.</summary>
    int Count { get; }

    /// <summary>The timeline with the given <see cref="IXLTimeline.Name"/>.</summary>
    /// <param name="name">The timeline's internal name, which is not its caption.</param>
    /// <exception cref="System.Collections.Generic.KeyNotFoundException">
    /// The worksheet has no timeline with that name.
    /// </exception>
    IXLTimeline Timeline(string name);

    /// <summary>Finds the timeline with the given <see cref="IXLTimeline.Name"/>.</summary>
    /// <param name="name">The timeline's internal name, which is not its caption.</param>
    /// <param name="timeline">The timeline, when one was found.</param>
    /// <returns>Whether a timeline with that name is on the worksheet.</returns>
    bool TryGetTimeline(string name, [NotNullWhen(true)] out IXLTimeline? timeline);
}
```

- [ ] **Step 6: Create `XLTimeline`**

```csharp
using System;
using System.Collections.Generic;
using System.Diagnostics;
using XLibur.Excel.Drawings;

namespace XLibur.Excel;

[DebuggerDisplay("{Name} ({SourceFieldName})")]
internal sealed class XLTimeline : IXLTimeline
{
    /// <summary>
    /// The size Excel gives a new timeline, in pixels — measured off the round-trip fixture's frame,
    /// 3,333,750 × 1,371,600 EMU at 96 dpi. A timeline is wider and shorter than a slicer.
    /// </summary>
    internal const int DefaultWidthPx = 350;

    /// <inheritdoc cref="DefaultWidthPx"/>
    internal const int DefaultHeightPx = 144;

    private readonly XLWorksheet _worksheet;
    private string _caption;
    private bool _showHeader = true;
    private bool _showSelectionLabel = true;
    private bool _showTimeLevel = true;
    private bool _showHorizontalScrollbar = true;
    private string? _style;
    private uint _level;

    internal XLTimeline(XLWorksheet worksheet, XLTimelineCache cache, string name)
    {
        _worksheet = worksheet;
        Cache = cache;
        Name = name;
        _caption = name;
    }

    /// <summary>The cache that binds the timeline to what it filters and holds its range.</summary>
    internal XLTimelineCache Cache { get; }

    /// <summary>
    /// The id of the relationship from the worksheet part to the timelines part this timeline was
    /// read from. Together with <see cref="Name"/> this is how the write path finds the element
    /// again to patch it. Null for a timeline not read from a package.
    /// </summary>
    internal string? PartRelId { get; set; }

    /// <summary>
    /// Whether the timeline was created through the API rather than read from a package. A new
    /// timeline is generated on save; a loaded one is only ever patched.
    /// </summary>
    internal bool IsNew { get; set; }

    /// <summary>Which properties the caller has assigned since the timeline was loaded.</summary>
    internal XLTimelineFormat AssignedFormat { get; private set; }

    /// <summary>
    /// The raw <c>@selectionLevel</c>, carried through untouched. XLibur does not model it — it only
    /// means anything alongside a selection, which is read-only.
    /// </summary>
    internal uint? SelectionLevelRaw { get; set; }

    /// <summary>
    /// The raw <c>@scrollPosition</c>, carried through untouched. It records where the user had
    /// scrolled the band, which XLibur has no reason to change.
    /// </summary>
    internal DateTime? ScrollPosition { get; set; }

    public string Name { get; }

    public string Caption
    {
        get => _caption;
        set
        {
            _caption = value ?? throw new ArgumentNullException(nameof(value));
            AssignedFormat |= XLTimelineFormat.Caption;
        }
    }

    public bool ShowHeader
    {
        get => _showHeader;
        set
        {
            _showHeader = value;
            AssignedFormat |= XLTimelineFormat.ShowHeader;
        }
    }

    public bool ShowSelectionLabel
    {
        get => _showSelectionLabel;
        set
        {
            _showSelectionLabel = value;
            AssignedFormat |= XLTimelineFormat.ShowSelectionLabel;
        }
    }

    public bool ShowTimeLevel
    {
        get => _showTimeLevel;
        set
        {
            _showTimeLevel = value;
            AssignedFormat |= XLTimelineFormat.ShowTimeLevel;
        }
    }

    public bool ShowHorizontalScrollbar
    {
        get => _showHorizontalScrollbar;
        set
        {
            _showHorizontalScrollbar = value;
            AssignedFormat |= XLTimelineFormat.ShowHorizontalScrollbar;
        }
    }

    public string? Style
    {
        get => _style;
        set
        {
            _style = value;
            AssignedFormat |= XLTimelineFormat.Style;
        }
    }

    /// <summary>
    /// The level, as the enumeration. The raw number is what is stored and what is written back, so
    /// a file carrying a value outside the enumeration round-trips its number rather than being
    /// narrowed to the nearest modelled one.
    /// </summary>
    public XLTimelineLevel Level
    {
        get => (XLTimelineLevel)_level;
        set
        {
            _level = (uint)value;
            AssignedFormat |= XLTimelineFormat.Level;
        }
    }

    /// <inheritdoc cref="Level"/>
    internal uint LevelRaw => _level;

    public IXLCell Position
    {
        get => FromMarker?.Cell ?? _worksheet.Cell(1, 1);
        set
        {
            if (value is null)
                throw new ArgumentNullException(nameof(value));

            // A fresh marker rather than a mutated one, because a marker registers itself with the
            // workbook's range repository so that inserting rows above the timeline moves it.
            // Setting the position drops any offset within the old cell: the caller named a cell,
            // so the corner goes to that cell's corner.
            FromMarker = new XLMarker(value);
            AssignedFormat |= XLTimelineFormat.Position;
        }
    }

    /// <summary>
    /// The timeline's top-left anchor point, with the offset within the cell that a file may carry.
    /// </summary>
    internal XLMarker? FromMarker { get; set; }

    /// <summary>
    /// The timeline's bottom-right anchor point, when it was read from a two-cell anchor. Kept so
    /// that moving a loaded timeline shifts both corners together and leaves its size alone.
    /// </summary>
    internal XLMarker? ToMarker { get; set; }

    internal int WidthPx { get; set; } = DefaultWidthPx;

    internal int HeightPx { get; set; } = DefaultHeightPx;

    /// <summary>
    /// Sets the properties read from a package without marking them as assigned.
    /// </summary>
    /// <remarks>
    /// This is what keeps <see cref="AssignedFormat"/> honest. It has to stay the only way the
    /// reader populates a timeline: assigning through the properties instead would mark every loaded
    /// timeline as edited, and the patcher would then rewrite parts nobody touched.
    /// </remarks>
    internal void SeedLoadedFormat(
        string caption,
        bool showHeader,
        bool showSelectionLabel,
        bool showTimeLevel,
        bool showHorizontalScrollbar,
        string? style,
        uint level)
    {
        _caption = caption;
        _showHeader = showHeader;
        _showSelectionLabel = showSelectionLabel;
        _showTimeLevel = showTimeLevel;
        _showHorizontalScrollbar = showHorizontalScrollbar;
        _style = style;
        _level = level;
    }

    public string SourceFieldName => Cache.SourceName;

    public IXLWorksheet Worksheet => _worksheet;

    public IReadOnlyList<IXLPivotTable> PivotTables => Cache.PivotTables;

    public DateTime? BoundsStart => Cache.BoundsStart;

    public DateTime? BoundsEnd => Cache.BoundsEnd;

    public bool HasSelection => Cache.SelectionStart is not null || Cache.SelectionEnd is not null;

    public DateTime? SelectionStart => Cache.SelectionStart;

    public DateTime? SelectionEnd => Cache.SelectionEnd;
}
```

- [ ] **Step 7: Create `XLTimelines` (read-only for now)**

`_worksheet` is stored but not read until Task 3's `Add` needs it. That is deliberate and does not
break the build: this repo runs neither `EnforceCodeStyleInBuild` nor the Sonar analyzers locally, so
an unread private readonly field produces no warning and therefore no error.

```csharp
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace XLibur.Excel;

internal sealed class XLTimelines : IXLTimelines
{
    private readonly List<XLTimeline> _timelines = [];
    private readonly XLWorksheet _worksheet;

    internal XLTimelines(XLWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public int Count => _timelines.Count;

    internal IReadOnlyList<XLTimeline> Items => _timelines;

    public IEnumerator<IXLTimeline> GetEnumerator() => _timelines.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IXLTimeline Timeline(string name)
    {
        if (!TryGetTimeline(name, out var timeline))
            throw new KeyNotFoundException($"The worksheet has no timeline named '{name}'.");

        return timeline;
    }

    public bool TryGetTimeline(string name, [NotNullWhen(true)] out IXLTimeline? timeline)
    {
        foreach (var candidate in _timelines)
        {
            if (XLHelper.NameComparer.Equals(candidate.Name, name))
            {
                timeline = candidate;
                return true;
            }
        }

        timeline = null;
        return false;
    }

    internal void Add(XLTimeline timeline) => _timelines.Add(timeline);

    /// <summary>
    /// Drops a timeline from the worksheet and records what the save path has to unpick.
    /// </summary>
    /// <remarks>
    /// Removing a timeline is not a matter of dropping one element. Its cache part, the workbook's
    /// registration of that cache, the <c>#N/A</c> defined name written for it, the worksheet's
    /// <c>extLst</c> reference and the drawing anchor all have to go with it, or the saved file has
    /// an orphan Excel will offer to repair.
    /// </remarks>
    internal void Remove(XLTimeline timeline)
    {
        if (_timelines.Remove(timeline) && !timeline.IsNew)
            Removed.Add(timeline);
    }

    /// <summary>
    /// Timelines removed since the workbook was loaded, still holding the relationship ids and cache
    /// names the save path needs to clean up after them. Cleared once a save has consumed it.
    /// </summary>
    internal List<XLTimeline> Removed { get; } = [];
}
```

- [ ] **Step 8: Create `TimelineAnchorXml` (read side)**

```csharp
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.IO.DrawingML;

namespace XLibur.Excel.IO;

/// <summary>
/// The drawing side of a timeline: the graphic frame Excel draws it through, and the anchor that
/// fixes that frame to the grid.
/// </summary>
/// <remarks>
/// A timeline is not drawn by its own part. <c>xl/timelines/timelineN.xml</c> says what the band
/// scrubs and what it looks like; where the band sits is a <c>xdr:graphicFrame</c> in the sheet's
/// drawing part holding a <c>tsle:timeslicer</c> element that carries nothing but the timeline's
/// name.
/// </remarks>
internal static class TimelineAnchorXml
{
    /// <summary>The graphic data URI Excel uses for a timeline frame.</summary>
    private const string TimelineGraphicUri = "http://schemas.microsoft.com/office/drawing/2012/timeslicer";

    internal static readonly DrawingFrameSpec Spec =
        new(TimelineGraphicUri, "tsle", "timeslicer", TimelineGraphicUri);

    /// <summary>
    /// Reads the anchor of every timeline on the sheet, so that a loaded timeline reports where it
    /// is.
    /// </summary>
    /// <remarks>
    /// The frame names the timeline, and the timeline name is unique within the workbook, which is
    /// what pairs the two up. All three anchor forms are read: Excel writes a two-cell anchor,
    /// XLibur writes a one-cell one, and a file from elsewhere may carry either or an absolute one.
    /// </remarks>
    internal static void ReadPositions(DrawingsPart? drawingsPart, XLTimelines timelines)
    {
        var worksheetDrawing = drawingsPart?.WorksheetDrawing;
        if (worksheetDrawing is null)
            return;

        foreach (var timeline in timelines.Items)
        {
            var anchor = DrawingFrameXml.FindAnchor(worksheetDrawing, Spec, timeline.Name);
            if (anchor is null)
                continue;

            var (from, to) = DrawingFrameXml.ReadMarkers(anchor, (XLWorksheet)timeline.Worksheet);
            if (from is not null)
                timeline.FromMarker = from;

            if (to is not null)
                timeline.ToMarker = to;
        }
    }
}
```

- [ ] **Step 9: Create `TimelineReader`**

```csharp
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads timelines and their caches out of a package and binds them to the pivot tables they filter.
/// </summary>
/// <remarks>
/// <para>
/// <b>Nothing here attaches a DOM to a part.</b> Timeline parts survive a round trip today because
/// <c>xl/timelines/*.xml</c> and <c>xl/timelineCaches/*.xml</c> are never opened, so they are copied
/// through byte for byte with every attribute XLibur has no model for. Reaching a part through
/// <c>TimeLinePart.Timelines</c> or <c>TimeLineCachePart.TimelineCacheDefinition</c> would create an
/// attached DOM that the SDK tracks and writes back on save, replacing those bytes with its own
/// serialisation. Every read below therefore goes through an <see cref="OpenXmlPartReader"/>, which
/// streams the part and hands back a detached tree the part knows nothing about.
/// </para>
/// <para>
/// Reading runs after pivot tables have been loaded, because binding needs them.
/// </para>
/// </remarks>
internal static class TimelineReader
{
    internal static void LoadTimelines(WorkbookPart workbookPart, Sheets sheets, XLWorksheets worksheets)
    {
        var caches = ReadCaches(workbookPart);
        if (caches.Count == 0)
            return;

        BindCaches(caches.Values, worksheets);
        ReadTimelines(workbookPart, sheets, worksheets, caches);
    }

    // ── Caches ──────────────────────────────────────────────────────────

    private static Dictionary<string, XLTimelineCache> ReadCaches(WorkbookPart workbookPart)
    {
        var caches = new Dictionary<string, XLTimelineCache>(XLHelper.NameComparer);

        foreach (var cachePart in workbookPart.TimeLineCacheParts)
        {
            var definition = ReadDetached<X15.TimelineCacheDefinition>(cachePart);
            var name = definition?.Name?.Value;
            if (name is null)
                continue;

            caches[name] = ReadCache(definition!, name, workbookPart.GetIdOfPart(cachePart));
        }

        return caches;
    }

    private static XLTimelineCache ReadCache(
        X15.TimelineCacheDefinition definition, string name, string relId)
    {
        var cache = new XLTimelineCache(name, definition.SourceName?.Value ?? string.Empty)
        {
            WorkbookRelId = relId,
        };

        foreach (var pivotTable in definition.TimelineCachePivotTables?
                     .Elements<X15.TimelineCachePivotTable>() ?? [])
        {
            if (pivotTable.Name?.Value is { } pivotTableName)
                cache.PivotTableNames.Add(pivotTableName);
        }

        if (definition.TimelineState is not { } state)
            return cache;

        cache.PivotCacheId = state.PivotCacheId?.Value;

        // The SDK types filterType as an enumeration, but InnerText is the raw token, which is what
        // carries a value written by a newer Excel through a save unchanged.
        if (state.FilterType?.InnerText is { Length: > 0 } filterType)
            cache.FilterType = filterType;

        if (state.MinimalRefreshVersion?.Value is { } minimal)
            cache.MinimalRefreshVersion = minimal;

        if (state.LastRefreshVersion?.Value is { } last)
            cache.LastRefreshVersion = last;

        if (state.BoundsTimelineRange is { } bounds)
        {
            cache.BoundsStart = bounds.StartDate?.Value;
            cache.BoundsEnd = bounds.EndDate?.Value;
        }

        if (state.SelectionTimelineRange is { } selection)
        {
            cache.SelectionStart = selection.StartDate?.Value;
            cache.SelectionEnd = selection.EndDate?.Value;
        }

        return cache;
    }

    // ── Binding ─────────────────────────────────────────────────────────

    private static void BindCaches(IEnumerable<XLTimelineCache> caches, XLWorksheets worksheets)
    {
        var pivotTables = PivotTablesByName(worksheets);

        foreach (var cache in caches)
        {
            // A cache may name several pivot tables, and may name one no longer in the workbook,
            // which is left out rather than reported as a hole in the list.
            foreach (var pivotTableName in cache.PivotTableNames)
            {
                if (pivotTables.TryGetValue(pivotTableName, out var pivotTable))
                    cache.PivotTables.Add(pivotTable);
            }

            // Pivot tables sharing a timeline cache share a pivot cache, so the first answers for all.
            cache.PivotCache = cache.PivotTables.Count > 0 ? cache.PivotTables[0].PivotCache : null;
        }
    }

    private static Dictionary<string, XLPivotTable> PivotTablesByName(XLWorksheets worksheets)
    {
        var pivotTables = new Dictionary<string, XLPivotTable>(XLHelper.NameComparer);
        foreach (var worksheet in worksheets)
        {
            foreach (var pivotTable in worksheet.PivotTables.Cast<XLPivotTable>())
                pivotTables[pivotTable.Name] = pivotTable;
        }

        return pivotTables;
    }

    // ── Timelines ───────────────────────────────────────────────────────

    private static void ReadTimelines(
        WorkbookPart workbookPart,
        Sheets sheets,
        XLWorksheets worksheets,
        Dictionary<string, XLTimelineCache> caches)
    {
        foreach (var (worksheetPart, worksheet) in WorksheetParts(workbookPart, sheets, worksheets))
        {
            foreach (var timelinePart in worksheetPart.TimeLineParts)
            {
                var timelines = ReadDetached<X15.Timelines>(timelinePart);
                if (timelines is null)
                    continue;

                var relId = worksheetPart.GetIdOfPart(timelinePart);
                foreach (var timeline in timelines.Elements<X15.Timeline>())
                    AddTimeline(timeline, relId, worksheet, caches);
            }

            // Where each timeline sits is in the drawing part, not the timeline part. Read after the
            // timelines exist, because the frames are matched to them by name.
            if (worksheet.TimelinesInternal.Count > 0)
                TimelineAnchorXml.ReadPositions(worksheetPart.DrawingsPart, worksheet.TimelinesInternal);
        }
    }

    private static void AddTimeline(
        X15.Timeline timeline,
        string relId,
        XLWorksheet worksheet,
        Dictionary<string, XLTimelineCache> caches)
    {
        var name = timeline.Name?.Value;
        var cacheName = timeline.Cache?.Value;
        if (name is null || cacheName is null || !caches.TryGetValue(cacheName, out var cache))
            return;

        var xlTimeline = new XLTimeline(worksheet, cache, name)
        {
            PartRelId = relId,
            SelectionLevelRaw = timeline.SelectionLevel?.Value,
            ScrollPosition = timeline.ScrollPosition?.Value,
        };

        // Seeded rather than assigned: going through the properties would mark every loaded timeline
        // as edited and bring parts nobody touched in for patching. The four booleans default to
        // true, which is what Excel means by omitting them.
        xlTimeline.SeedLoadedFormat(
            timeline.Caption?.Value ?? name,
            timeline.ShowHeader?.Value ?? true,
            timeline.ShowSelectionLabel?.Value ?? true,
            timeline.ShowTimeLevel?.Value ?? true,
            timeline.ShowHorizontalScrollbar?.Value ?? true,
            timeline.Style?.Value,
            timeline.Level?.Value ?? 0);

        worksheet.TimelinesInternal.Add(xlTimeline);
    }

    // ── Plumbing ────────────────────────────────────────────────────────

    /// <summary>Pairs each worksheet part with the loaded worksheet it belongs to, in sheet order.</summary>
    private static IEnumerable<(WorksheetPart Part, XLWorksheet Worksheet)> WorksheetParts(
        WorkbookPart workbookPart, Sheets sheets, XLWorksheets worksheets)
    {
        foreach (var sheet in sheets.OfType<Sheet>())
        {
            // A sheet with an empty relationship id comes from a non-Excel producer, and the
            // relationship may point at a chartsheet rather than a worksheet.
            if (string.IsNullOrEmpty(sheet.Id?.Value)
                || sheet.Name?.Value is not { } sheetName
                || workbookPart.GetPartById(sheet.Id!.Value!) is not WorksheetPart worksheetPart
                || !worksheets.TryGetWorksheet(sheetName, out var worksheet))
            {
                continue;
            }

            yield return (worksheetPart, worksheet);
        }
    }

    /// <summary>Reads a part's root element without attaching it to the part.</summary>
    /// <remarks>
    /// This is the whole fidelity guarantee of this reader in three lines: the part is streamed, the
    /// element that comes back is detached, and <c>part.RootElement</c> stays unmaterialised, so the
    /// SDK has nothing to write back over the original bytes when the package is saved.
    /// </remarks>
    private static T? ReadDetached<T>(OpenXmlPart part) where T : OpenXmlElement
    {
        using var reader = new OpenXmlPartReader(part);

        // Create reads the XML declaration only, so the first Read lands on the root element.
        return reader.Read() ? reader.LoadCurrentElement() as T : null;
    }
}
```

- [ ] **Step 10: Wire the collections and the load hook**

In `XLibur/Excel/XLWorksheet.cs`, immediately after the `SlicersInternal` property:

```csharp
    private XLTimelines? _timelines;

    public IXLTimelines Timelines => TimelinesInternal;

    /// <summary>
    /// The timelines drawn on this sheet, as the concrete collection the reader adds to.
    /// </summary>
    internal XLTimelines TimelinesInternal => _timelines ??= new XLTimelines(this);
```

In `XLibur/Excel/IXLWorksheet.cs`, after the `Slicers` declaration:

```csharp
    /// <summary>
    /// The timelines drawn on this worksheet. The worksheet owns them; what each one filters is
    /// reached through <see cref="IXLTimeline.PivotTables"/>.
    /// </summary>
    IXLTimelines Timelines { get; }
```

In `XLibur/Excel/PivotTables/IXLPivotTable.cs`, after `Slicers`:

```csharp
    /// <summary>
    /// The timelines that filter this pivot table. A view over the worksheets that own them
    /// (<see cref="IXLWorksheet.Timelines"/>), which need not be this pivot table's sheet.
    /// </summary>
    IEnumerable<IXLTimeline> Timelines { get; }
```

In `XLibur/Excel/PivotTables/XLPivotTable.cs`, immediately after the `Slicers` property:

```csharp
    /// <summary>
    /// The timelines whose cache lists this pivot table, gathered from every worksheet in the
    /// workbook.
    /// </summary>
    /// <remarks>
    /// Recomputed on each access rather than kept as a back-reference, for the reason
    /// <see cref="Slicers"/> is: a derived view cannot drift out of step with the worksheets that
    /// own them, which matters most in exactly the case a cached list would get wrong.
    /// </remarks>
    public IEnumerable<IXLTimeline> Timelines
    {
        get
        {
            foreach (var worksheet in _worksheet.Workbook.WorksheetsInternal)
            {
                foreach (var timeline in worksheet.TimelinesInternal.Items)
                {
                    if (timeline.Cache.PivotTables.Contains(this))
                        yield return timeline;
                }
            }
        }
    }
```

In `XLibur/Excel/XLWorkbook_Load.cs`, immediately after the `SlicerReader.LoadSlicers` call:

```csharp
        // Same ordering constraint as slicers: a timeline binds to the pivot tables it filters, and
        // they have to exist before it can find them.
        TimelineReader.LoadTimelines(workbookPart, sheets!, WorksheetsInternal);
```

- [ ] **Step 11: Add the public API entries**

Append to `XLibur/PublicAPI.Unshipped.txt`, keeping the file sorted:

```
XLibur.Excel.IXLPivotTable.Timelines.get -> System.Collections.Generic.IEnumerable<XLibur.Excel.IXLTimeline!>!
XLibur.Excel.IXLTimeline
XLibur.Excel.IXLTimeline.BoundsEnd.get -> System.DateTime?
XLibur.Excel.IXLTimeline.BoundsStart.get -> System.DateTime?
XLibur.Excel.IXLTimeline.Caption.get -> string!
XLibur.Excel.IXLTimeline.Caption.set -> void
XLibur.Excel.IXLTimeline.HasSelection.get -> bool
XLibur.Excel.IXLTimeline.Level.get -> XLibur.Excel.XLTimelineLevel
XLibur.Excel.IXLTimeline.Level.set -> void
XLibur.Excel.IXLTimeline.Name.get -> string!
XLibur.Excel.IXLTimeline.PivotTables.get -> System.Collections.Generic.IReadOnlyList<XLibur.Excel.IXLPivotTable!>!
XLibur.Excel.IXLTimeline.Position.get -> XLibur.Excel.IXLCell!
XLibur.Excel.IXLTimeline.Position.set -> void
XLibur.Excel.IXLTimeline.SelectionEnd.get -> System.DateTime?
XLibur.Excel.IXLTimeline.SelectionStart.get -> System.DateTime?
XLibur.Excel.IXLTimeline.ShowHeader.get -> bool
XLibur.Excel.IXLTimeline.ShowHeader.set -> void
XLibur.Excel.IXLTimeline.ShowHorizontalScrollbar.get -> bool
XLibur.Excel.IXLTimeline.ShowHorizontalScrollbar.set -> void
XLibur.Excel.IXLTimeline.ShowSelectionLabel.get -> bool
XLibur.Excel.IXLTimeline.ShowSelectionLabel.set -> void
XLibur.Excel.IXLTimeline.ShowTimeLevel.get -> bool
XLibur.Excel.IXLTimeline.ShowTimeLevel.set -> void
XLibur.Excel.IXLTimeline.SourceFieldName.get -> string!
XLibur.Excel.IXLTimeline.Style.get -> string?
XLibur.Excel.IXLTimeline.Style.set -> void
XLibur.Excel.IXLTimeline.Worksheet.get -> XLibur.Excel.IXLWorksheet!
XLibur.Excel.IXLTimelines
XLibur.Excel.IXLTimelines.Count.get -> int
XLibur.Excel.IXLTimelines.Timeline(string! name) -> XLibur.Excel.IXLTimeline!
XLibur.Excel.IXLTimelines.TryGetTimeline(string! name, out XLibur.Excel.IXLTimeline? timeline) -> bool
XLibur.Excel.IXLWorksheet.Timelines.get -> XLibur.Excel.IXLTimelines!
XLibur.Excel.XLTimelineLevel
XLibur.Excel.XLTimelineLevel.Days = 3 -> XLibur.Excel.XLTimelineLevel
XLibur.Excel.XLTimelineLevel.Months = 2 -> XLibur.Excel.XLTimelineLevel
XLibur.Excel.XLTimelineLevel.Quarters = 1 -> XLibur.Excel.XLTimelineLevel
XLibur.Excel.XLTimelineLevel.Years = 0 -> XLibur.Excel.XLTimelineLevel
```

- [ ] **Step 12: Run the test to verify it passes**

```
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/TimelineReadModelTests/*"
```

Expected: PASS, all ten. If `Reading_a_timeline_does_not_rewrite_its_parts` fails, something reached a typed part property — search the new files for `.Timelines` and `.TimelineCacheDefinition` used as a *getter on a part*, which is the only way this breaks.

- [ ] **Step 13: Run the whole suite**

```
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0
```

Expected: PASS. `RoundTripFidelityTests.Timelines_and_their_caches_survive_a_round_trip` and `LoadingTests.Can_load_and_save_preserves_timelines` both exercise this fixture already and must be unaffected.

- [ ] **Step 14: Update the fidelity doc**

In `docs/round-trip-fidelity.md`, replace the Timelines row:

```
| Timelines | ✅ | n/a | ❌ | `timelines/` and `timelineCaches/` |
```

with:

```
| Timelines | ✅ | ✅ | ✅ | Modelled as of spec 35; a timeline nobody edits is still passed through untouched |
```

The middle column was wrong before this change, not only after it: `LoadingTests` already asserted the sheet's `timelineRef` survives.

- [ ] **Step 15: Commit**

```
git add XLibur/Excel/Timelines XLibur/Excel/IO/TimelineReader.cs XLibur/Excel/IO/TimelineAnchorXml.cs XLibur/Excel/XLWorksheet.cs XLibur/Excel/IXLWorksheet.cs XLibur/Excel/PivotTables/IXLPivotTable.cs XLibur/Excel/PivotTables/XLPivotTable.cs XLibur/Excel/XLWorkbook_Load.cs XLibur/PublicAPI.Unshipped.txt XLibur.Tests/Excel/Timelines/TimelineReadModelTests.cs docs/round-trip-fidelity.md
git commit -m 'feat(timelines): read the timelines a workbook already has

IXLWorksheet.Timelines owns; IXLPivotTable.Timelines is a view over the
timelines whose cache names that pivot table. A timeline reports its
name, caption, level, style, source field, bound pivot tables, bounds
and selection.

Every read streams its part through OpenXmlPartReader and returns a
detached tree. Reaching TimeLinePart.Timelines instead would attach a
DOM the SDK re-serialises on save, replacing Excel bytes that carry
mc:Ignorable and every attribute XLibur does not model — and timeline
parts survive a round trip today for exactly one reason, which is that
nothing opens them. A byte-equality test asserts it.

The selection is read-only. See docs/specs/35-pivot-timelines.md.'
```

---

## Task 3: Create, patch and position

**Files:**
- Create: `XLibur/Excel/IO/TimelineWriter.cs`, `XLibur/Excel/IO/TimelineCacheWriter.cs`, `XLibur/Excel/IO/TimelinePatcher.cs`
- Modify: `XLibur/Excel/IO/TimelineAnchorXml.cs` (add `Append`, `Move`, `Remove`)
- Modify: `XLibur/Excel/Timelines/XLTimelines.cs` (add `Add`), `XLibur/Excel/Timelines/IXLTimelines.cs` (declare `Add`)
- Modify: `XLibur/Excel/XLWorkbook_Save.cs:191` and `:220`, `XLibur/Excel/IO/WorksheetPartWriter.cs:220`, `XLibur/PublicAPI.Unshipped.txt`
- Test: `XLibur.Tests/Excel/Timelines/TimelineWriteTests.cs`, `XLibur.Tests/Excel/Timelines/TimelinePositionTests.cs`

**Interfaces:**
- Consumes: everything from tasks 1 and 2; `DrawingAnchorFactory.Create`, `DrawingPartScaffold.EnsureDrawingsPart/EnsureNamespaces/EnsureDrawingElement`, `XLPivotCache.TryGetFieldIndex`, `XLPivotCache.GetFieldValues(int).Stats`, `XLWorkbook.RelIdGenerator`.
- Produces:
  - `IXLTimeline IXLTimelines.Add(IXLPivotTable pivotTable, string dateFieldName)`
  - `static void TimelineWriter.WriteTimelines(Worksheet, XLWorksheetContentManager, XLWorksheet, WorksheetPart, SaveContext)`
  - `static void TimelineCacheWriter.PrepareTimelineCaches(WorkbookPart, XLWorkbook, SaveContext)`
  - `static void TimelineCacheWriter.WriteTimelineCaches(WorkbookPart, XLWorkbook, SaveContext)`
  - `static void TimelinePatcher.PatchTimeline(WorksheetPart, XLTimeline)`
  - `static void TimelineAnchorXml.Append(Xdr.WorksheetDrawing, XLTimeline)`, `.Move(DrawingsPart, XLTimeline)`, `.Remove(DrawingsPart?, XLTimeline)`

- [ ] **Step 1: Write the failing create test**

Create `XLibur.Tests/Excel/Timelines/TimelineWriteTests.cs`.

```csharp
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Timelines;

/// <summary>
/// Creating timelines, editing loaded ones, and the parts neither operation may touch.
/// </summary>
/// <remarks>
/// A created timeline needs six things or Excel offers to repair the file: the timeline definition,
/// the worksheet's <c>extLst</c> reference to it, the cache part, the workbook's <c>extLst</c>
/// registration of that cache, a <c>#N/A</c> defined name, and a drawing anchor. All six are
/// asserted here.
/// </remarks>
public class TimelineWriteTests
{
    private const string Fixture = @"TryToLoad\Timelines_Missing_21232.xlsx";

    // ── Creating ────────────────────────────────────────────────────────

    [Test]
    public async Task A_created_timeline_writes_all_six_pieces()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
            wb.Worksheet("Data").Timelines.Add(pivotTable, "Date");
            wb.SaveAs(saved);
        }

        var entries = EntryNames(saved);

        // 1 and 2: the timeline definition and the cache part. The fixture already owns
        // timeline1/timelineCache1, so the created pair must be new parts rather than additions to
        // the existing ones.
        await Assert.That(entries.Count(n => n.StartsWith("xl/timelines/", StringComparison.Ordinal))).IsEqualTo(2);
        await Assert.That(entries.Count(n => n.StartsWith("xl/timelineCaches/", StringComparison.Ordinal))).IsEqualTo(2);

        // 3: the worksheet's extLst reference, on the sheet the timeline is drawn on.
        await Assert.That(ReadPart(saved, "xl/worksheets/sheet2.xml")).Contains("timelineRef");

        var workbookXml = ReadPart(saved, "xl/workbook.xml");

        // 4: the workbook registration, and 5: the #N/A defined name.
        await Assert.That(workbookXml).Contains("{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}");
        await Assert.That(workbookXml).Contains("NativeTimeline_Date");
        await Assert.That(workbookXml).Contains("#N/A");

        // 6: the drawing anchor.
        await Assert.That(ReadPart(saved, "xl/drawings/drawing2.xml")).Contains("timeslicer");
    }

    [Test]
    public async Task A_created_timeline_reloads_with_what_it_was_given()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
            var timeline = wb.Worksheet("Data").Timelines.Add(pivotTable, "Date");
            timeline.Caption = "Pick a period";
            timeline.Style = "TimeSlicerStyleLight2";
            timeline.Level = XLTimelineLevel.Quarters;
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var reloadedTimeline = reloaded.Worksheet("Data").Timelines.Single();

        await Assert.That(reloadedTimeline.Caption).IsEqualTo("Pick a period");
        await Assert.That(reloadedTimeline.Style).IsEqualTo("TimeSlicerStyleLight2");
        await Assert.That(reloadedTimeline.Level).IsEqualTo(XLTimelineLevel.Quarters);
        await Assert.That(reloadedTimeline.SourceFieldName).IsEqualTo("Date");
        await Assert.That(reloadedTimeline.PivotTables.Single().Name).IsEqualTo("СводнаяТаблица2");
    }

    [Test]
    public async Task A_created_timeline_takes_its_bounds_from_the_fields_dates()
    {
        using var wb = Load();

        var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
        var timeline = wb.Worksheet("Data").Timelines.Add(pivotTable, "Date");

        // The field's dates run 1998-05-19 to 2004-02-06; Excel rounds outward to whole years.
        await Assert.That(timeline.BoundsStart).IsEqualTo(new DateTime(1998, 1, 1));
        await Assert.That(timeline.BoundsEnd).IsEqualTo(new DateTime(2005, 1, 1));
        await Assert.That(timeline.HasSelection).IsFalse();
    }

    [Test]
    public async Task A_timeline_over_a_field_that_holds_no_dates_is_refused()
    {
        using var wb = Load();

        var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
        var timelines = wb.Worksheet("Data").Timelines;

        // A timeline over a text field is a repair prompt, not a degraded timeline.
        await Assert.That(() => timelines.Add(pivotTable, "Name")).Throws<ArgumentException>();
        await Assert.That(() => timelines.Add(pivotTable, "NoSuchField")).Throws<ArgumentException>();
    }

    // ── The lesson from PRD 5 defect 4 ──────────────────────────────────

    [Test]
    public async Task Adding_a_timeline_beside_an_existing_one_leaves_that_ones_part_untouched()
    {
        // The guard PRD 5's slicer tests were missing. Three byte-equality assertions passed
        // throughout a feature that did not work, because each covered only a sheet where nothing
        // had been added. This adds a timeline to the sheet that already has one.
        using var original = Resource();
        var before = PartBytes(original, "xl/timelines/timeline1.xml");

        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
            wb.Worksheet("Pivot").Timelines.Add(pivotTable, "Date");
            wb.SaveAs(saved);
        }

        await Assert.That(PartBytes(saved, "xl/timelines/timeline1.xml")).IsEquivalentTo(before);
    }

    [Test]
    public async Task A_created_timeline_gets_a_part_of_its_own()
    {
        // Every timelines part Excel writes holds exactly one x15:timeline. Appending into the
        // sheet's existing part instead is what broke slicers: it opens a part Excel authored and
        // hands the SDK the job of serialising it again.
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
            wb.Worksheet("Pivot").Timelines.Add(pivotTable, "Date");
            wb.SaveAs(saved);
        }

        var parts = EntryNames(saved)
            .Where(n => n.StartsWith("xl/timelines/", StringComparison.Ordinal))
            .ToList();

        await Assert.That(parts.Count).IsEqualTo(2);

        foreach (var part in parts)
        {
            var xml = ReadPart(saved, part);
            await Assert.That(CountOccurrences(xml, "<x15:timeline ") + CountOccurrences(xml, "<timeline "))
                .IsEqualTo(1)
                .Because($"{part} must hold exactly one timeline.");
        }
    }

    // ── Editing a loaded timeline ───────────────────────────────────────

    [Test]
    public async Task Editing_a_loaded_timeline_keeps_everything_else_in_its_part()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var timeline = wb.Worksheet("Pivot").Timelines.Single();
            timeline.Caption = "Pick a period";
            timeline.Level = XLTimelineLevel.Quarters;
            wb.SaveAs(saved);
        }

        var xml = ReadPart(saved, "xl/timelines/timeline1.xml");

        await Assert.That(xml).Contains("caption=\"Pick a period\"");
        await Assert.That(xml).Contains("level=\"1\"");

        // Everything XLibur does not model survived, which is the whole point of patching rather
        // than regenerating. selectionLevel and scrollPosition are attributes no XLibur API produces.
        await Assert.That(xml).Contains("selectionLevel=\"2\"");
        await Assert.That(xml).Contains("scrollPosition=\"2004-06-07T00:00:00\"");
        await Assert.That(xml).Contains("mc:Ignorable");
    }

    [Test]
    public async Task A_timeline_nobody_touched_is_not_written_to()
    {
        // Loading a workbook and saving it after an unrelated edit must not open the timeline part.
        using var original = Resource();
        var before = PartBytes(original, "xl/timelines/timeline1.xml");

        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Data").Cell("Z99").Value = "unrelated";
            wb.SaveAs(saved);
        }

        await Assert.That(PartBytes(saved, "xl/timelines/timeline1.xml")).IsEquivalentTo(before);
    }

    // ── Schema ──────────────────────────────────────────────────────────

    [Test]
    public async Task A_package_with_a_created_timeline_is_schema_valid()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
            wb.Worksheet("Data").Timelines.Add(pivotTable, "Date");
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2013)
            .Validate(doc)
            .Select(error => $"{error.Path?.XPath}: {error.Description}")
            .ToList();

        await Assert.That(errors).IsEmpty();
    }

    #region Helpers

    private static XLWorkbook Load()
    {
        var stream = Resource();
        stream.Position = 0;
        return new XLWorkbook(stream);
    }

    private static MemoryStream Resource()
    {
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Fixture));
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms;
    }

    private static byte[] PartBytes(MemoryStream package, string partPath)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.First(e =>
            e.FullName.Equals(partPath, StringComparison.OrdinalIgnoreCase));

        using var entryStream = entry.Open();
        using var buffer = new MemoryStream();
        entryStream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static string ReadPart(MemoryStream package, string partPath) =>
        Encoding.UTF8.GetString(PartBytes(package, partPath));

    private static System.Collections.Generic.List<string> EntryNames(MemoryStream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        return archive.Entries.Select(e => e.FullName).ToList();
    }

    private static int CountOccurrences(string haystack, string needle)
    {
        var count = 0;
        var index = 0;
        while ((index = haystack.IndexOf(needle, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += needle.Length;
        }

        return count;
    }

    #endregion
}
```

- [ ] **Step 2: Run the test to verify it fails**

```
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/TimelineWriteTests/*"
```

Expected: compile error — `IXLTimelines` has no `Add`.

- [ ] **Step 3: Add `Add` to the interface and the collection**

In `IXLTimelines.cs`, add:

```csharp
    /// <summary>
    /// Adds a timeline that filters a pivot table on one of its date fields.
    /// </summary>
    /// <param name="pivotTable">The pivot table to filter. It need not be on this worksheet.</param>
    /// <param name="dateFieldName">The pivot cache field the band is drawn from.</param>
    /// <returns>The new timeline, showing every date and so filtering nothing.</returns>
    /// <exception cref="System.ArgumentException">
    /// The pivot table's cache has no field of that name, or that field holds no dates.
    /// </exception>
    /// <remarks>
    /// The timeline is placed to the right of the pivot table. Use <see cref="IXLTimeline.Position"/>
    /// to move it.
    /// </remarks>
    IXLTimeline Add(IXLPivotTable pivotTable, string dateFieldName);
```

In `XLTimelines.cs`, add these members and the `using System;`, `using System.Globalization;`, `using System.Text;`, `using XLibur.Excel.Drawings;` they need:

```csharp
    public IXLTimeline Add(IXLPivotTable pivotTable, string dateFieldName) =>
        AddTimeline((XLPivotTable)pivotTable, dateFieldName);

    internal XLTimeline AddTimeline(XLPivotTable pivotTable, string dateFieldName)
    {
        var pivotCache = (XLPivotCache)pivotTable.PivotCache;
        if (!pivotCache.TryGetFieldIndex(dateFieldName, out var fieldIndex))
        {
            throw new ArgumentException(
                $"The pivot cache of '{pivotTable.Name}' has no field named '{dateFieldName}'.",
                nameof(dateFieldName));
        }

        // Excel decides a field is timeline-able from the date statistics on its shared items. A
        // timeline over a field that holds no dates is a repair prompt, not a degraded timeline, so
        // it is refused here rather than written and discovered in Excel.
        var stats = pivotCache.GetFieldValues(fieldIndex).Stats;
        if (!stats.ContainsDate || stats.MinDate is not { } minDate || stats.MaxDate is not { } maxDate)
        {
            throw new ArgumentException(
                $"The field '{dateFieldName}' holds no dates, so it cannot carry a timeline.",
                nameof(dateFieldName));
        }

        var cache = new XLTimelineCache(NextCacheName(dateFieldName), dateFieldName)
        {
            IsNew = true,
            PivotCache = pivotCache,

            // Excel rounds the field's range outward to whole years — the round-trip fixture's field
            // runs 1998-05-19 to 2004-02-06 and its bounds read 1998-01-01 to 2005-01-01.
            BoundsStart = new DateTime(minDate.Year, 1, 1),
            BoundsEnd = new DateTime(maxDate.Year + 1, 1, 1),
        };
        cache.PivotTables.Add(pivotTable);
        cache.PivotTableNames.Add(pivotTable.Name);

        var area = pivotTable.Area;
        return AddNew(cache, dateFieldName, DefaultPositionBeside(area.FirstPoint.Row, area.LastPoint.Column));
    }

    private XLTimeline AddNew(XLTimelineCache cache, string sourceName, IXLCell position)
    {
        var name = NextTimelineName(sourceName);
        var timeline = new XLTimeline(_worksheet, cache, name)
        {
            IsNew = true,
            FromMarker = new XLMarker(position),
        };

        // Seeded rather than assigned, so a timeline created and left alone carries no pending
        // edits. Months is the level Excel starts a new timeline at.
        timeline.SeedLoadedFormat(
            name,
            showHeader: true,
            showSelectionLabel: true,
            showTimeLevel: true,
            showHorizontalScrollbar: true,
            style: null,
            level: (uint)XLTimelineLevel.Months);

        _timelines.Add(timeline);
        return timeline;
    }

    /// <summary>
    /// Where a new timeline goes when the caller has not said: two columns to the right of the pivot
    /// table it filters, at that table's top row.
    /// </summary>
    /// <remarks>
    /// A default is not optional here. <c>DrawingAnchorFactory</c> documents that a drawing handed no
    /// marker gets one at A1 — silently, with no exception and no missing element. For a picture
    /// that is reasonable. For a timeline it would drop the band over the top-left of the sheet,
    /// covering the very data it filters, and the caller would have no idea why. So every timeline
    /// XLibur creates is given a marker before the factory sees it, and that fallback stays
    /// unreachable from here.
    /// </remarks>
    private IXLCell DefaultPositionBeside(int topRow, int rightmostColumn) =>
        _worksheet.Cell(
            Math.Max(1, topRow),
            Math.Min(XLHelper.MaxColumnNumber, rightmostColumn + 2));

    /// <summary>
    /// A cache name not already taken, in the shape Excel uses: <c>NativeTimeline_Date</c>, then
    /// <c>NativeTimeline_Date1</c>.
    /// </summary>
    /// <remarks>
    /// The name is not decoration. A timeline refers to its cache by it, and a <c>#N/A</c> defined
    /// name is written under the same name, so it has to be a legal defined name: no spaces, and
    /// nothing that would parse as a cell reference.
    /// </remarks>
    private string NextCacheName(string sourceName)
    {
        var stem = "NativeTimeline_" + Sanitise(sourceName);
        var taken = WorkbookCacheNames();

        if (!taken.Contains(stem))
            return stem;

        for (var suffix = 1; ; suffix++)
        {
            var candidate = stem + suffix.ToString(CultureInfo.InvariantCulture);
            if (!taken.Contains(candidate))
                return candidate;
        }
    }

    /// <summary>
    /// A timeline name not already taken, in the shape Excel uses: <c>Date</c>, then <c>Date 1</c>.
    /// Timeline names are unique across the workbook, not just the sheet.
    /// </summary>
    private string NextTimelineName(string sourceName)
    {
        var taken = new HashSet<string>(XLHelper.NameComparer);
        foreach (var worksheet in _worksheet.Workbook.WorksheetsInternal)
        {
            foreach (var timeline in worksheet.TimelinesInternal.Items)
                taken.Add(timeline.Name);
        }

        if (!taken.Contains(sourceName))
            return sourceName;

        for (var suffix = 1; ; suffix++)
        {
            var candidate = sourceName + " " + suffix.ToString(CultureInfo.InvariantCulture);
            if (!taken.Contains(candidate))
                return candidate;
        }
    }

    private HashSet<string> WorkbookCacheNames()
    {
        var taken = new HashSet<string>(XLHelper.NameComparer);
        foreach (var worksheet in _worksheet.Workbook.WorksheetsInternal)
        {
            foreach (var timeline in worksheet.TimelinesInternal.Items)
                taken.Add(timeline.Cache.Name);
        }

        // A defined name already using the stem would collide with the one written for the cache.
        foreach (var definedName in _worksheet.Workbook.DefinedNamesInternal)
            taken.Add(definedName.Name);

        return taken;
    }

    private static string Sanitise(string sourceName)
    {
        var builder = new StringBuilder(sourceName.Length);
        foreach (var c in sourceName)
            builder.Append(char.IsLetterOrDigit(c) || c == '_' ? c : '_');

        return builder.Length > 0 ? builder.ToString() : "Field";
    }
```

- [ ] **Step 4: Add `Append`, `Move` and `Remove` to `TimelineAnchorXml`**

```csharp
    /// <summary>
    /// Appends the anchored graphic frame for a newly created timeline to the sheet's drawing.
    /// </summary>
    /// <remarks>
    /// Timelines take <see cref="XLPicturePlacement.Move"/>: a one-cell anchor takes a top-left
    /// marker and an explicit size, which is exactly what a timeline has, and it means the band
    /// moves with the rows and columns above it without being stretched by them. The factory's A1
    /// fallback is never taken — a created timeline is always given a marker first.
    /// </remarks>
    internal static void Append(Xdr.WorksheetDrawing worksheetDrawing, XLTimeline xlTimeline)
    {
        var frame = DrawingFrameXml.BuildFrame(worksheetDrawing, Spec, xlTimeline.Name);

        var anchor = DrawingAnchorFactory.Create(
            XLPicturePlacement.Move,
            new DrawingAnchorGeometry
            {
                Worksheet = xlTimeline.Worksheet,
                LeftPx = 0,
                TopPx = 0,
                WidthPx = xlTimeline.WidthPx,
                HeightPx = xlTimeline.HeightPx,
                FromMarker = xlTimeline.FromMarker,
                ToMarker = xlTimeline.ToMarker,
            },
            frame);

        worksheetDrawing.Append(anchor);
    }

    /// <summary>
    /// Moves the frame of a loaded timeline, shifting both corners by the same number of rows and
    /// columns so the band keeps the size it had.
    /// </summary>
    internal static void Move(DrawingsPart drawingsPart, XLTimeline xlTimeline)
    {
        if (xlTimeline.FromMarker is not { } target)
            return;

        DrawingFrameXml.MoveAnchor(drawingsPart, Spec, xlTimeline.Name, target);
    }

    /// <summary>Takes the anchored frame of a removed timeline out of the sheet's drawing.</summary>
    internal static void Remove(DrawingsPart? drawingsPart, XLTimeline xlTimeline) =>
        DrawingFrameXml.RemoveAnchor(drawingsPart, Spec, xlTimeline.Name);
```

Add `using XLibur.Excel.Drawings;` and `using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;` to the file.

- [ ] **Step 5: Create `TimelinePatcher`**

```csharp
using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Applies model changes to a timeline that already exists in the package.
/// </summary>
/// <remarks>
/// <para>
/// XLibur never regenerates the XML of a timeline it read from a file. That is what carries a
/// timeline's <c>selectionLevel</c>, its <c>scrollPosition</c>, the extension list Excel hangs off
/// it and any attribute a future Excel invents through a load and save untouched.
/// </para>
/// <para>
/// The price of that guarantee is that an edit has to be patched into the element the reader saw.
/// This does exactly that, and only for the properties the caller actually assigned (see
/// <see cref="XLTimeline.AssignedFormat"/>): a timeline nobody edited is not written to at all, and
/// the part is not even opened for it.
/// </para>
/// </remarks>
internal static class TimelinePatcher
{
    internal static void PatchTimeline(WorksheetPart worksheetPart, XLTimeline xlTimeline)
    {
        if (xlTimeline.AssignedFormat == XLTimelineFormat.None)
            return;

        // Moving a timeline touches the drawing part rather than the timelines part, so the two are
        // resolved separately: assigning only a position must not open the timelines part, and
        // assigning only a caption must not open the drawing.
        if (xlTimeline.AssignedFormat.HasFlag(XLTimelineFormat.Position)
            && worksheetPart.DrawingsPart is { } drawingsPart)
        {
            TimelineAnchorXml.Move(drawingsPart, xlTimeline);
        }

        if ((xlTimeline.AssignedFormat & ~XLTimelineFormat.Position) == XLTimelineFormat.None)
            return;

        var part = ResolvePart(worksheetPart, xlTimeline);
        if (part?.Timelines is not { } timelines)
            return;

        var timeline = timelines
            .Elements<X15.Timeline>()
            .FirstOrDefault(t => string.Equals(t.Name?.Value, xlTimeline.Name, StringComparison.Ordinal));
        if (timeline is null)
            return;

        Apply(timeline, xlTimeline);
    }

    private static void Apply(X15.Timeline timeline, XLTimeline xlTimeline)
    {
        var assigned = xlTimeline.AssignedFormat;

        // Each optional attribute is cleared with a typed null rather than a bare one. Assigning
        // `null` goes through the implicit conversion from string or bool and produces a value
        // wrapping null — serialised as `caption=""`, not as an absent attribute. The cast is what
        // actually removes it.
        if (assigned.HasFlag(XLTimelineFormat.Caption))
        {
            // Excel omits the caption when it matches the name and shows the name instead, so
            // setting the caption back to the name removes the attribute rather than restating it.
            timeline.Caption = string.Equals(xlTimeline.Caption, xlTimeline.Name, StringComparison.Ordinal)
                ? (StringValue?)null
                : xlTimeline.Caption;
        }

        // The four booleans default to true; writing a value that is already the default is legal
        // but noisy, and it is not what Excel does.
        if (assigned.HasFlag(XLTimelineFormat.ShowHeader))
            timeline.ShowHeader = xlTimeline.ShowHeader ? (BooleanValue?)null : false;

        if (assigned.HasFlag(XLTimelineFormat.ShowSelectionLabel))
            timeline.ShowSelectionLabel = xlTimeline.ShowSelectionLabel ? (BooleanValue?)null : false;

        if (assigned.HasFlag(XLTimelineFormat.ShowTimeLevel))
            timeline.ShowTimeLevel = xlTimeline.ShowTimeLevel ? (BooleanValue?)null : false;

        if (assigned.HasFlag(XLTimelineFormat.ShowHorizontalScrollbar))
        {
            timeline.ShowHorizontalScrollbar =
                xlTimeline.ShowHorizontalScrollbar ? (BooleanValue?)null : false;
        }

        if (assigned.HasFlag(XLTimelineFormat.Style))
            timeline.Style = xlTimeline.Style is { } style ? style : (StringValue?)null;

        // level defaults to 0, so a timeline set back to Years drops the attribute.
        if (assigned.HasFlag(XLTimelineFormat.Level))
            timeline.Level = xlTimeline.LevelRaw == 0 ? (UInt32Value?)null : xlTimeline.LevelRaw;
    }

    /// <summary>
    /// The timelines part a loaded timeline was read from.
    /// </summary>
    /// <remarks>
    /// Opening the part here is what finally attaches its DOM, and it happens only for a timeline
    /// with a pending change. Everything <see cref="TimelineReader"/> does is deliberately detached
    /// so that this is the single point at which a timeline part stops being copied through verbatim.
    /// </remarks>
    private static TimeLinePart? ResolvePart(WorksheetPart worksheetPart, XLTimeline xlTimeline)
    {
        if (xlTimeline.PartRelId is null)
            return null;

        // GetPartById throws for an unknown id, which is reachable when the timeline came from a
        // package this one was not saved from.
        if (!worksheetPart.Parts.Any(p => p.RelationshipId == xlTimeline.PartRelId))
            return null;

        return worksheetPart.GetPartById(xlTimeline.PartRelId) as TimeLinePart;
    }
}
```

- [ ] **Step 6: Create `TimelineWriter`**

```csharp
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using XLibur.Excel.IO.DrawingML;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes the worksheet half of a timeline: the timelines part holding its definition and the
/// <c>extLst</c> reference that makes the worksheet point at it.
/// </summary>
/// <remarks>
/// A surviving timeline part is worthless if the sheet stops referencing it, and the worksheet part
/// is rebuilt from the model on every save while the timeline part is not.
/// </remarks>
internal static class TimelineWriter
{
    /// <summary>The worksheet extension holding the list of timelines on the sheet.</summary>
    private const string TimelineExtensionUri = "{7E03D99C-DC04-49d9-9315-930204A7B6E9}";

    private const string X15Main2010SsNs = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

    internal static void WriteTimelines(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        WorksheetPart worksheetPart,
        SaveContext context)
    {
        var timelines = xlWorksheet.TimelinesInternal;

        RemoveDeletedTimelines(worksheet, cm, worksheetPart, timelines);

        foreach (var timeline in timelines.Items)
        {
            if (!timeline.IsNew)
            {
                // A timeline that already exists in the package is never regenerated — that is what
                // keeps the parts of its XML XLibur does not model intact. Only the properties the
                // caller actually changed are patched into the existing part.
                TimelinePatcher.PatchTimeline(worksheetPart, timeline);
                continue;
            }

            WriteNewTimeline(worksheet, cm, worksheetPart, timeline, context);
            timeline.IsNew = false;
        }
    }

    private static void WriteNewTimeline(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        WorksheetPart worksheetPart,
        XLTimeline xlTimeline,
        SaveContext context)
    {
        var relId = context.RelIdGenerator.GetNext(RelType.Workbook);
        var part = worksheetPart.AddNewPart<TimeLinePart>(relId);
        xlTimeline.PartRelId = relId;

        var root = new X15.Timelines();
        root.AddNamespaceDeclaration("x", Main2006SsNs);

        var timeline = new X15.Timeline
        {
            Name = xlTimeline.Name,
            Cache = xlTimeline.Cache.Name,
            Caption = xlTimeline.Caption,
        };

        // Attributes at their schema default are left off, which is what Excel writes and what keeps
        // a generated part comparable with a hand-made one.
        if (!xlTimeline.ShowHeader)
            timeline.ShowHeader = false;

        if (!xlTimeline.ShowSelectionLabel)
            timeline.ShowSelectionLabel = false;

        if (!xlTimeline.ShowTimeLevel)
            timeline.ShowTimeLevel = false;

        if (!xlTimeline.ShowHorizontalScrollbar)
            timeline.ShowHorizontalScrollbar = false;

        if (xlTimeline.Style is { } style)
            timeline.Style = style;

        if (xlTimeline.LevelRaw != 0)
        {
            timeline.Level = xlTimeline.LevelRaw;

            // Excel writes selectionLevel alongside level and keeps the two in step on a timeline
            // that has never been scrubbed.
            timeline.SelectionLevel = xlTimeline.LevelRaw;
        }

        root.AppendChild(timeline);
        part.Timelines = root;

        EnsureTimelineReference(worksheet, cm, relId);
        WriteAnchor(worksheet, cm, worksheetPart, xlTimeline, context);
    }

    /// <summary>
    /// Draws the timeline: the graphic frame in the sheet's drawing part, and the sheet's reference
    /// to that part.
    /// </summary>
    /// <remarks>
    /// The sixth of the six pieces a created timeline needs. Without it the workbook opens and the
    /// timeline is simply not there, because <c>xl/timelines/timelineN.xml</c> says what a timeline
    /// scrubs but never where it sits.
    /// </remarks>
    private static void WriteAnchor(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        WorksheetPart worksheetPart,
        XLTimeline xlTimeline,
        SaveContext context)
    {
        var drawingsPart = DrawingPartScaffold.EnsureDrawingsPart(worksheetPart, context);
        var worksheetDrawing = drawingsPart.WorksheetDrawing!;
        DrawingPartScaffold.EnsureNamespaces(worksheetDrawing);

        TimelineAnchorXml.Append(worksheetDrawing, xlTimeline);

        DrawingPartScaffold.EnsureDrawingElement(worksheet, cm, worksheetPart, drawingsPart);
    }

    private static void EnsureTimelineReference(
        Worksheet worksheet, XLWorksheetContentManager cm, string relId)
    {
        var list = SheetExtensionRefs.EnsureList<X15.TimelineReferences>(
            worksheet, cm, TimelineExtensionUri, "x15", X15Main2010SsNs);

        if (!list.Elements<X15.TimelineReference>().Any(r => r.Id?.Value == relId))
            list.AppendChild(new X15.TimelineReference { Id = relId });
    }

    /// <summary>
    /// Unpicks the worksheet half of every timeline removed since the workbook was loaded.
    /// </summary>
    /// <remarks>
    /// Because a created timeline always gets a part of its own and a loaded one always has one,
    /// removing a timeline removes the whole part rather than one element inside it. The extension
    /// goes once its list is empty, and the extension list once it is. The anchored frame goes from
    /// the sheet's drawing as well, since that is what Excel actually draws a timeline through.
    /// </remarks>
    private static void RemoveDeletedTimelines(
        Worksheet worksheet, XLWorksheetContentManager cm, WorksheetPart worksheetPart, XLTimelines timelines)
    {
        if (timelines.Removed.Count == 0)
            return;

        foreach (var removed in timelines.Removed)
        {
            // The frame lives in the drawing rather than in the timelines part, so it has to go too —
            // otherwise the sheet still asks Excel to draw something the package no longer defines.
            TimelineAnchorXml.Remove(worksheetPart.DrawingsPart, removed);

            if (removed.PartRelId is not { } relId
                || !worksheetPart.Parts.Any(p => p.RelationshipId == relId)
                || worksheetPart.GetPartById(relId) is not TimeLinePart part)
            {
                continue;
            }

            SheetExtensionRefs.RemoveRefs<X15.TimelineReferences>(
                worksheet, cm, r => r is X15.TimelineReference reference && reference.Id?.Value == relId);

            worksheetPart.DeletePart(part);
        }
    }
}
```

- [ ] **Step 7: Create `TimelineCacheWriter`**

```csharp
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static XLibur.Excel.XLWorkbook;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes the workbook half of a timeline: its cache part, the workbook's registration of that cache
/// and the defined name Excel writes alongside it.
/// </summary>
/// <remarks>
/// <para>
/// A created timeline needs six things and Excel offers to repair the file if any is missing: the
/// timeline definition and its worksheet reference (both <see cref="TimelineWriter"/>), the cache
/// part, the workbook <c>extLst</c> registration, the <c>#N/A</c> defined name, and a drawing
/// anchor. Everything but the anchor is written here or next door.
/// </para>
/// <para>
/// Loaded caches are not rewritten. As with <see cref="TimelinePatcher"/>, the part a timeline was
/// read from is left exactly as it arrived.
/// </para>
/// </remarks>
internal static class TimelineCacheWriter
{
    /// <summary>The workbook extension registering timeline caches.</summary>
    private const string TimelineCachesExtensionUri = "{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}";

    private const string X15Main2010SsNs = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

    /// <summary>
    /// Allocates cache parts and defined names for timelines created since the last save, and takes
    /// them away again for timelines removed since.
    /// </summary>
    /// <remarks>
    /// Runs before the workbook part is generated, because the <c>#N/A</c> defined names have to be
    /// in the model by the time <see cref="WorkbookPartWriter"/> rebuilds the whole defined-name
    /// block from it. The cache parts only get their content later, once the pivot cache identifier
    /// they quote has been assigned.
    /// </remarks>
    internal static void PrepareTimelineCaches(
        WorkbookPart workbookPart, XLWorkbook workbook, SaveContext context)
    {
        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            foreach (var removed in worksheet.TimelinesInternal.Removed)
                RemoveCache(workbookPart, workbook, removed.Cache);

            foreach (var timeline in worksheet.TimelinesInternal.Items)
            {
                if (timeline.IsNew)
                    AddCache(workbookPart, workbook, timeline.Cache, context);
            }
        }
    }

    /// <summary>
    /// Writes the content of every timeline cache part created this save, and registers it in the
    /// workbook's extension list.
    /// </summary>
    /// <remarks>
    /// Runs after the worksheets, because the cache quotes the pivot cache identifier, which does
    /// not exist until the pivot cache part has been generated.
    /// </remarks>
    internal static void WriteTimelineCaches(
        WorkbookPart workbookPart, XLWorkbook workbook, SaveContext context)
    {
        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            foreach (var timeline in worksheet.TimelinesInternal.Items)
            {
                var cache = timeline.Cache;
                if (!cache.IsNew || cache.WorkbookRelId is not { } relId)
                    continue;

                var part = (TimeLineCachePart)workbookPart.GetPartById(relId);
                part.TimelineCacheDefinition = BuildDefinition(cache);

                RegisterCache(workbookPart, relId);
                cache.IsNew = false;
            }

            worksheet.TimelinesInternal.Removed.Clear();
        }
    }

    // ── Cache parts ─────────────────────────────────────────────────────

    private static void AddCache(
        WorkbookPart workbookPart, XLWorkbook workbook, XLTimelineCache cache, SaveContext context)
    {
        if (cache.WorkbookRelId is not null)
            return;

        var relId = context.RelIdGenerator.GetNext(RelType.Workbook);
        cache.WorkbookRelId = relId;
        workbookPart.AddNewPart<TimeLineCachePart>(relId);

        // Excel writes a #N/A defined name per timeline cache, named after the cache. Adding it the
        // way the reader does keeps it out of formula validation, which would reject #N/A.
        if (workbook.DefinedNamesInternal.All<XLDefinedName>(n => !XLHelper.NameComparer.Equals(n.Name, cache.Name)))
        {
            workbook.DefinedNamesInternal.Add(
                cache.Name, "#N/A", comment: null, validateName: false, validateRangeAddress: false);
        }

        // A timeline cache names its pivot cache by an identifier that lives in an extension of the
        // pivot cache definition — the same one a slicer cache quotes. A cache read from a file
        // already has one; a cache XLibur created has none until now, and the pivot cache writer
        // emits it once this is set.
        if (cache.PivotCache is { } pivotCache)
            pivotCache.PivotCacheId ??= NextPivotCacheId(workbook);
    }

    private static void RemoveCache(WorkbookPart workbookPart, XLWorkbook workbook, XLTimelineCache cache)
    {
        if (cache.WorkbookRelId is not { } relId)
            return;

        if (workbookPart.Parts.Any(p => p.RelationshipId == relId)
            && workbookPart.GetPartById(relId) is TimeLineCachePart part)
        {
            workbookPart.DeletePart(part);
        }

        UnregisterCache(workbookPart, relId);

        var definedName = workbook.DefinedNamesInternal
            .FirstOrDefault<XLDefinedName>(n => XLHelper.NameComparer.Equals(n.Name, cache.Name));
        if (definedName is not null)
            workbook.DefinedNamesInternal.Delete(definedName.Name);

        cache.WorkbookRelId = null;
    }

    /// <summary>
    /// An identifier no pivot cache in the workbook is already using.
    /// </summary>
    /// <remarks>
    /// Counting up from the highest in use keeps it deterministic, which matters because a save has
    /// to be reproducible. Shared with the slicer path by convention rather than by code: both read
    /// and write the same <c>XLPivotCache.PivotCacheId</c>, so a workbook holding both never
    /// allocates a collision.
    /// </remarks>
    private static uint NextPivotCacheId(XLWorkbook workbook)
    {
        uint highest = 0;
        foreach (var cache in workbook.PivotCachesInternal)
        {
            if (cache.PivotCacheId is { } id && id > highest)
                highest = id;
        }

        return highest + 1;
    }

    private static X15.TimelineCacheDefinition BuildDefinition(XLTimelineCache cache)
    {
        var definition = new X15.TimelineCacheDefinition
        {
            Name = cache.Name,
            SourceName = cache.SourceName,
        };
        definition.AddNamespaceDeclaration("x15", X15Main2010SsNs);

        var pivotTables = new X15.TimelineCachePivotTables();
        foreach (var pivotTable in cache.PivotTables)
        {
            // tabId is the sheet the pivot table lives on, which need not be the sheet the timeline
            // is drawn on.
            var sheetId = ((XLWorksheet)pivotTable.Worksheet).SheetId;
            pivotTables.AppendChild(new X15.TimelineCachePivotTable
            {
                TabId = (uint)sheetId,
                Name = pivotTable.Name,
            });
        }

        definition.AppendChild(pivotTables);

        var state = new X15.TimelineState
        {
            MinimalRefreshVersion = cache.MinimalRefreshVersion,
            LastRefreshVersion = cache.LastRefreshVersion,
            PivotCacheId = cache.PivotCache?.PivotCacheId ?? 0,

            // The SDK types this as an enumeration, but InnerText is what actually serialises, so a
            // token this build does not know still round-trips.
            FilterType = new EnumValue<PivotFilterValues> { InnerText = cache.FilterType },
        };

        if (cache.BoundsStart is { } boundsStart && cache.BoundsEnd is { } boundsEnd)
        {
            state.AppendChild(new X15.BoundsTimelineRange
            {
                StartDate = boundsStart,
                EndDate = boundsEnd,
            });
        }

        definition.AppendChild(state);
        return definition;
    }

    // ── Workbook registration ───────────────────────────────────────────

    private static void RegisterCache(WorkbookPart workbookPart, string relId)
    {
        var workbook = workbookPart.Workbook!;
        var extensionList = workbook.GetFirstChild<WorkbookExtensionList>();
        if (extensionList is null)
        {
            extensionList = new WorkbookExtensionList();
            workbook.AppendChild(extensionList);
        }

        var extension = FindExtension(extensionList);
        if (extension is null)
        {
            extension = new WorkbookExtension { Uri = TimelineCachesExtensionUri };
            extension.AddNamespaceDeclaration("x15", X15Main2010SsNs);
            extension.AppendChild(new X15.TimelineCacheReferences());
            extensionList.AppendChild(extension);
        }

        var container = extension.GetFirstChild<X15.TimelineCacheReferences>();
        if (container is null)
            return;

        if (!container.Elements<X15.TimelineCacheReference>().Any(c => c.Id?.Value == relId))
            container.AppendChild(new X15.TimelineCacheReference { Id = relId });
    }

    private static void UnregisterCache(WorkbookPart workbookPart, string relId)
    {
        var extensionList = workbookPart.Workbook?.GetFirstChild<WorkbookExtensionList>();
        var extension = extensionList is null ? null : FindExtension(extensionList);
        var container = extension?.GetFirstChild<X15.TimelineCacheReferences>();
        if (container is null)
            return;

        foreach (var registration in container
                     .Elements<X15.TimelineCacheReference>()
                     .Where(c => c.Id?.Value == relId)
                     .ToList())
        {
            registration.Remove();
        }

        // An empty registry is a schema violation rather than merely untidy, so the extension goes
        // once its last cache does.
        if (!container.Elements<X15.TimelineCacheReference>().Any())
            extension!.Remove();

        if (extensionList is { HasChildren: false })
            extensionList.Remove();
    }

    private static WorkbookExtension? FindExtension(WorkbookExtensionList extensionList) =>
        extensionList.Elements<WorkbookExtension>()
            .FirstOrDefault(e => string.Equals(
                e.Uri?.Value, TimelineCachesExtensionUri, System.StringComparison.OrdinalIgnoreCase));
}
```

- [ ] **Step 8: Wire the save hooks**

In `XLibur/Excel/XLWorkbook_Save.cs`, immediately after `SlicerCacheWriter.PrepareSlicerCaches(workbookPart, this, context);`:

```csharp
        TimelineCacheWriter.PrepareTimelineCaches(workbookPart, this, context);
```

Immediately after `SlicerCacheWriter.WriteSlicerCaches(workbookPart, this, context);`:

```csharp
        TimelineCacheWriter.WriteTimelineCaches(workbookPart, this, context);
```

In `XLibur/Excel/IO/WorksheetPartWriter.cs`, immediately after `SlicerWriter.WriteSlicers(...)`:

```csharp
        TimelineWriter.WriteTimelines(worksheet, cm, xlWorksheet, worksheetPart, context);
```

- [ ] **Step 9: Add the public API entry**

Append to `XLibur/PublicAPI.Unshipped.txt`, in sorted position:

```
XLibur.Excel.IXLTimelines.Add(XLibur.Excel.IXLPivotTable! pivotTable, string! dateFieldName) -> XLibur.Excel.IXLTimeline!
```

- [ ] **Step 10: Run the write tests**

```
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/TimelineWriteTests/*"
```

Expected: PASS, all nine.

If `A_package_with_a_created_timeline_is_schema_valid` fails, read the reported XPath before changing anything — the equivalent slicer work found three real schema facts this way, and each was in the writer, not in the validator.

If `A_created_timeline_writes_all_six_pieces` fails on the drawing assertion, check the sheet number in the path: `xl/drawings/drawing2.xml` assumes the `Data` sheet gets a new drawing part. Print `EntryNames(saved)` to see what was actually written and correct the path in the test.

- [ ] **Step 11: Write the positioning test**

Create `XLibur.Tests/Excel/Timelines/TimelinePositionTests.cs`.

```csharp
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Timelines;

/// <summary>
/// Where a timeline sits, which is in the sheet's drawing part rather than in the timeline's own.
/// </summary>
public class TimelinePositionTests
{
    private const string Fixture = @"TryToLoad\Timelines_Missing_21232.xlsx";

    [Test]
    public async Task A_created_timeline_lands_clear_of_the_pivot_table()
    {
        using var wb = Load();

        var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
        var timeline = wb.Worksheet("Data").Timelines.Add(pivotTable, "Date");

        // Two columns right of the pivot table's rightmost column, at its top row. The pivot table
        // occupies A3:B14, so the timeline goes to D3.
        await Assert.That(timeline.Position.Address.ToString()).IsEqualTo("D3");
    }

    [Test]
    public async Task Moving_a_created_timeline_puts_the_anchor_where_it_was_told()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            var pivotTable = wb.Worksheet("Pivot").PivotTables.Single();
            var timeline = wb.Worksheet("Data").Timelines.Add(pivotTable, "Date");
            timeline.Position = wb.Worksheet("Data").Cell("F5");
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);

        await Assert.That(reloaded.Worksheet("Data").Timelines.Single().Position.Address.ToString())
            .IsEqualTo("F5");
    }

    [Test]
    public async Task Moving_a_loaded_timeline_edits_its_anchor_rather_than_replacing_it()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Pivot").Timelines.Single().Position = wb.Worksheet("Pivot").Cell("E4");
            wb.SaveAs(saved);
        }

        var drawing = ReadPart(saved, "xl/drawings/drawing1.xml");

        // The corner moved: C2 (col 2, row 1) to E4 (col 4, row 3).
        await Assert.That(drawing).Contains("<xdr:col>4</xdr:col>");

        // And Excel's own wrapper survived, which is what replacing the anchor would have destroyed.
        await Assert.That(drawing).Contains("mc:AlternateContent");
        await Assert.That(drawing).Contains("mc:Fallback");

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheet("Pivot").Timelines.Single().Position.Address.ToString())
            .IsEqualTo("E4");
    }

    [Test]
    public async Task Moving_a_loaded_timeline_keeps_its_size()
    {
        using var saved = new MemoryStream();
        using (var wb = Load())
        {
            wb.Worksheet("Pivot").Timelines.Single().Position = wb.Worksheet("Pivot").Cell("E4");
            wb.SaveAs(saved);
        }

        var drawing = ReadPart(saved, "xl/drawings/drawing1.xml");

        // Both corners shifted by the same delta — two columns and two rows — so the band is the
        // same size it was. The original spans col 2..8, row 1..9.
        await Assert.That(drawing).Contains("<xdr:col>4</xdr:col>");
        await Assert.That(drawing).Contains("<xdr:col>10</xdr:col>");
        await Assert.That(drawing).Contains("<xdr:row>3</xdr:row>");
        await Assert.That(drawing).Contains("<xdr:row>11</xdr:row>");
    }

    #region Helpers

    private static XLWorkbook Load()
    {
        using var source = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(Fixture));
        var ms = new MemoryStream();
        source.CopyTo(ms);
        ms.Position = 0;
        return new XLWorkbook(ms);
    }

    private static string ReadPart(MemoryStream package, string partPath)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.First(e =>
            e.FullName.Equals(partPath, StringComparison.OrdinalIgnoreCase));

        using var entryStream = entry.Open();
        using var reader = new StreamReader(entryStream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    #endregion
}
```

- [ ] **Step 12: Run the positioning tests**

```
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/TimelinePositionTests/*"
```

Expected: PASS.

`A_created_timeline_lands_clear_of_the_pivot_table` asserts `D3` from the fixture's `<location ref="A3:B14">`. Confirm with `unzip -p XLibur.Tests/Resource/TryToLoad/Timelines_Missing_21232.xlsx xl/pivotTables/pivotTable1.xml | grep -o '<location[^>]*>'` and correct the expected address if the area differs.

- [ ] **Step 13: Run the whole suite**

```
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0
```

Expected: PASS.

- [ ] **Step 14: Generate the acceptance-check workbooks**

Add `XLibur.Tests/Excel/Timelines/AcceptanceCheckWorkbooks.cs`. It is a generator, not a gate, so it is
skipped by default and run by name when workbooks are wanted. Everything it does goes **through the
public API only** — a check workbook built through internals proves nothing about what a caller gets.

```csharp
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.Timelines;

/// <summary>
/// Writes the workbooks a human opens in Excel to settle acceptance criteria 3, 4 and 7.
/// </summary>
/// <remarks>
/// Skipped in normal runs. Nothing here asserts anything: the whole point is that the automated
/// suite cannot see the failure these files exist to catch. PRD 5's central finding was that every
/// automated gate passed on a slicer feature Excel refused to render.
/// </remarks>
public class AcceptanceCheckWorkbooks
{
    private const string OutputDirectory = @"..\..\..\..\scratchpad\ac-check-timelines";

    [Test]
    [Skip("Generator, not a gate. Run by name when check workbooks are wanted.")]
    public async Task Generate()
    {
        Directory.CreateDirectory(OutputDirectory);
        var sha = Environment.GetEnvironmentVariable("AC_SHA") ?? "unstamped";

        WriteCreatedTimeline($"ac3-created-timeline-{sha}.xlsx");
        WriteSecondTimeline($"ac4-timeline-beside-an-existing-one-{sha}.xlsx");
        WriteCascade($"ac7-pivot-deleted-timeline-cascades-{sha}.xlsx");

        await Assert.That(Directory.GetFiles(OutputDirectory, "*.xlsx").Length).IsGreaterThanOrEqualTo(3);
    }

    /// <summary>Criterion 3: a created timeline opens, is drawn where it was put, and filters.</summary>
    private static void WriteCreatedTimeline(string fileName)
    {
        using var wb = new XLWorkbook();
        var data = wb.AddWorksheet("Data");

        data.Cell("A1").Value = "Date";
        data.Cell("B1").Value = "Region";
        data.Cell("C1").Value = "Amount";

        var start = new DateTime(2024, 1, 15);
        for (var i = 0; i < 24; i++)
        {
            data.Cell(i + 2, 1).Value = start.AddDays(i * 11);
            data.Cell(i + 2, 2).Value = i % 2 == 0 ? "North" : "South";
            data.Cell(i + 2, 3).Value = 100 + (i * 7);
        }

        data.Column(1).Style.DateFormat.Format = "yyyy-mm-dd";

        var pivotSheet = wb.AddWorksheet("Pivot");
        var pivotTable = pivotSheet.PivotTables.Add("SalesPivot", pivotSheet.Cell("A3"), data.Range("A1:C25"));
        pivotTable.RowLabels.Add("Region");
        pivotTable.Values.Add("Amount");

        var timeline = pivotSheet.Timelines.Add(pivotTable, "Date");
        timeline.Caption = "Pick a period";
        timeline.Style = "TimeSlicerStyleLight2";
        timeline.Position = pivotSheet.Cell("E3");

        wb.SaveAs(Path.Combine(OutputDirectory, fileName));
    }

    /// <summary>Criterion 4: a second timeline must not stop Excel drawing the first.</summary>
    private static void WriteSecondTimeline(string fileName)
    {
        using var source = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"TryToLoad\Timelines_Missing_21232.xlsx"));

        using var wb = new XLWorkbook(source);
        var pivotSheet = wb.Worksheet("Pivot");
        var added = pivotSheet.Timelines.Add(pivotSheet.PivotTables.Single(), "Date");
        added.Caption = "Added by XLibur";
        added.Position = pivotSheet.Cell("C12");

        wb.SaveAs(Path.Combine(OutputDirectory, fileName));
    }

    /// <summary>Criterion 7: deleting the pivot table leaves no orphan for Excel to repair.</summary>
    private static void WriteCascade(string fileName)
    {
        using var source = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"TryToLoad\Timelines_Missing_21232.xlsx"));

        using var wb = new XLWorkbook(source);
        var pivotSheet = wb.Worksheet("Pivot");
        pivotSheet.PivotTables.Delete(pivotSheet.PivotTables.Single().Name);

        wb.SaveAs(Path.Combine(OutputDirectory, fileName));
    }
}
```

Confirm the `PivotTables.Add` and `RowLabels`/`Values` signatures against an existing create-path test
under `XLibur.Tests/Excel/PivotTables/Create/` before running — that is the API this repo already
exercises, and it is the reference, not this snippet.

Run it with the sha stamped in, **after** the commit in Step 15:

```
git rev-parse --short HEAD
```

```
AC_SHA=<that sha> dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/AcceptanceCheckWorkbooks/*"
```

Generate these last and regenerate them if any later commit touches timeline code. The first PRD 5
check set was generated minutes *before* the fixes it was meant to exercise, and a false failure
nearly cost a correct fix — which is why the sha goes in the filename rather than being inferred from
a timestamp.

Alongside them write `scratchpad/ac-check-timelines/check-sheet.md` giving, for each file: the
click-by-click steps, what correct looks like, and what each failure mode would point at. For `ac3`
that is: the band appears at E3 captioned "Pick a period"; dragging it narrows the pivot table. If
the band is absent while the pivot table renders, the first suspect is the bare graphic frame — see
the spec's Risks table.

- [ ] **Step 15: Commit**

```
git add XLibur/Excel/IO/TimelineWriter.cs XLibur/Excel/IO/TimelineCacheWriter.cs XLibur/Excel/IO/TimelinePatcher.cs XLibur/Excel/IO/TimelineAnchorXml.cs XLibur/Excel/Timelines XLibur/Excel/XLWorkbook_Save.cs XLibur/Excel/IO/WorksheetPartWriter.cs XLibur/PublicAPI.Unshipped.txt XLibur.Tests/Excel/Timelines
git commit -m 'feat(timelines): create, patch and position timelines

Worksheet.Timelines.Add(pivotTable, dateField) writes the six pieces a
timeline needs: the definition, the sheet extLst reference, the cache
part, the workbook registration, the #N/A defined name and the drawing
anchor. A field holding no dates is refused rather than written and
discovered in Excel.

A created timeline always gets a part of its own. Appending into a part
Excel wrote is what made Excel stop drawing the slicer already on the
sheet in PRD 5 defect 4, and the guard those tests were missing is here:
byte equality on the existing timeline part while a second one is added
beside it.

Editing a loaded timeline patches only the assigned attributes into the
element the reader saw, so selectionLevel, scrollPosition and everything
else XLibur does not model survives.'
```

---

## Task 4: Cascade on pivot-table deletion

**Files:**
- Modify (rename): `XLibur/Excel/Slicers/XLSlicerCascade.cs` → `XLibur/Excel/PivotTables/XLPivotDependentCascade.cs`
- Modify: `XLibur/Excel/PivotTables/XLPivotTables.cs:65` and `:74`
- Modify: `docs/round-trip-fidelity.md` (Gaps section)
- Test: `XLibur.Tests/Excel/PivotTables/PivotTableDeletionTests.cs`

**Interfaces:**
- Consumes: `XLTimelines.Remove`, `XLTimelineCache.PivotTables`/`PivotTableNames` (task 2); `TimelineCacheWriter`, `TimelineWriter` removal paths (task 3).
- Produces: `static void XLPivotDependentCascade.OnPivotTableDeleted(XLWorkbook, XLPivotTable)` — replaces `XLSlicerCascade.OnPivotTableDeleted` with identical semantics plus the timeline arm.

- [ ] **Step 1: Write the failing cascade test**

Append to `XLibur.Tests/Excel/PivotTables/PivotTableDeletionTests.cs`:

```csharp
    [Test]
    public async Task Deleting_a_pivot_table_takes_its_timelines_with_it()
    {
        using var saved = new MemoryStream();

        using (var wb = TimelineWorkbook())
        {
            var pivotSheet = wb.Worksheet("Pivot");
            await Assert.That(pivotSheet.Timelines.Count).IsEqualTo(1);

            pivotSheet.PivotTables.Delete(pivotSheet.PivotTables.Single().Name);

            // The cache served only that pivot table, so the timeline has nothing left to filter.
            await Assert.That(pivotSheet.Timelines.Count).IsEqualTo(0);

            wb.SaveAs(saved);
        }

        var entries = EntryNames(saved);

        // The part, the cache part and the #N/A defined name all go, or the saved file has an orphan
        // Excel will offer to repair.
        await Assert.That(entries.Any(n => n.StartsWith("xl/timelines/", StringComparison.Ordinal))).IsFalse();
        await Assert.That(entries.Any(n => n.StartsWith("xl/timelineCaches/", StringComparison.Ordinal))).IsFalse();

        var workbookXml = ReadPart(saved, "xl/workbook.xml");
        await Assert.That(workbookXml).DoesNotContain("timelineCacheRef");
        await Assert.That(workbookXml).DoesNotContain("ВстроеннаяВременнаяШкала_Date");

        // And the drawing no longer asks Excel to draw a band the package does not define.
        await Assert.That(ReadPart(saved, "xl/drawings/drawing1.xml")).DoesNotContain("timeslicer");
    }

    [Test]
    public async Task A_workbook_whose_pivot_table_was_deleted_is_schema_valid()
    {
        using var saved = new MemoryStream();

        using (var wb = TimelineWorkbook())
        {
            var pivotSheet = wb.Worksheet("Pivot");
            pivotSheet.PivotTables.Delete(pivotSheet.PivotTables.Single().Name);
            wb.SaveAs(saved);
        }

        saved.Position = 0;
        using var doc = SpreadsheetDocument.Open(saved, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2013)
            .Validate(doc)
            .Select(error => $"{error.Path?.XPath}: {error.Description}")
            .ToList();

        await Assert.That(errors).IsEmpty();
    }

    private static XLWorkbook TimelineWorkbook()
    {
        using var source = TestHelper.GetStreamFromResource(
            TestHelper.GetResourcePath(@"TryToLoad\Timelines_Missing_21232.xlsx"));
        var ms = new MemoryStream();
        source.CopyTo(ms);
        ms.Position = 0;
        return new XLWorkbook(ms);
    }

    private static System.Collections.Generic.List<string> EntryNames(MemoryStream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        return archive.Entries.Select(e => e.FullName).ToList();
    }
```

If `PivotTableDeletionTests` already declares a `ReadPart` or `EntryNames` helper, reuse it instead of adding a second.

- [ ] **Step 2: Run the test to verify it fails**

```
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/PivotTableDeletionTests/*"
```

Expected: `Deleting_a_pivot_table_takes_its_timelines_with_it` FAILS at `Timelines.Count` — it is still 1, because nothing connects the two.

- [ ] **Step 3: Rename the cascade and add the timeline arm**

Move `XLibur/Excel/Slicers/XLSlicerCascade.cs` to `XLibur/Excel/PivotTables/XLPivotDependentCascade.cs` with `git mv`, then rewrite it:

```csharp
using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// Keeps the controls that filter a pivot table consistent with it.
/// </summary>
/// <remarks>
/// <para>
/// A slicer cache and a timeline cache both name the pivot tables they drive. Deleting one of them
/// and leaving a cache pointing at a pivot table that is no longer there produces a file Excel
/// offers to repair, so the reference has to go with the pivot table. When the deleted pivot table
/// was the last one a cache served, the control has nothing left to filter and goes too — along with
/// its cache part, the workbook's registration of that cache, the <c>#N/A</c> defined name written
/// for it and its drawing anchor.
/// </para>
/// <para>
/// This closes a gap that predates either being modelled: before, deleting a pivot table left the
/// parts untouched in the package, because nothing knew they were connected. Timelines were the last
/// instance of that hazard named in <c>docs/round-trip-fidelity.md</c>.
/// </para>
/// </remarks>
internal static class XLPivotDependentCascade
{
    /// <summary>
    /// Drops the deleted pivot table from every slicer and timeline cache that named it, and removes
    /// any control left with nothing to filter.
    /// </summary>
    internal static void OnPivotTableDeleted(XLWorkbook workbook, XLPivotTable pivotTable)
    {
        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            RemoveOrphanedSlicers(worksheet, pivotTable);
            RemoveOrphanedTimelines(worksheet, pivotTable);
        }
    }

    private static void RemoveOrphanedSlicers(XLWorksheet worksheet, XLPivotTable pivotTable)
    {
        List<XLSlicer>? orphaned = null;

        foreach (var slicer in worksheet.SlicersInternal.Items)
        {
            var cache = slicer.Cache;
            if (!cache.PivotTables.Remove(pivotTable))
                continue;

            cache.PivotTableNames.RemoveAll(name => XLHelper.NameComparer.Equals(name, pivotTable.Name));

            // Other pivot tables still share this cache, so the slicer keeps working and only loses
            // one of its connections.
            if (cache.PivotTables.Count > 0)
                continue;

            (orphaned ??= []).Add(slicer);
        }

        if (orphaned is null)
            return;

        foreach (var slicer in orphaned)
            worksheet.SlicersInternal.Remove(slicer);
    }

    private static void RemoveOrphanedTimelines(XLWorksheet worksheet, XLPivotTable pivotTable)
    {
        List<XLTimeline>? orphaned = null;

        foreach (var timeline in worksheet.TimelinesInternal.Items)
        {
            var cache = timeline.Cache;
            if (!cache.PivotTables.Remove(pivotTable))
                continue;

            cache.PivotTableNames.RemoveAll(name => XLHelper.NameComparer.Equals(name, pivotTable.Name));

            if (cache.PivotTables.Count > 0)
                continue;

            (orphaned ??= []).Add(timeline);
        }

        if (orphaned is null)
            return;

        foreach (var timeline in orphaned)
            worksheet.TimelinesInternal.Remove(timeline);
    }
}
```

- [ ] **Step 4: Point the two call sites at the new name**

In `XLibur/Excel/PivotTables/XLPivotTables.cs`, replace both occurrences of

```csharp
            XLSlicerCascade.OnPivotTableDeleted(Worksheet.Workbook, pivotTable);
```

with

```csharp
            XLPivotDependentCascade.OnPivotTableDeleted(Worksheet.Workbook, pivotTable);
```

- [ ] **Step 5: Run the deletion tests**

```
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/PivotTableDeletionTests/*"
```

Expected: PASS, including the slicer cascade tests already in that file — the rename must not change their behaviour.

- [ ] **Step 6: Run the whole suite**

```
dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0
```

Expected: PASS.

- [ ] **Step 7: Update the fidelity doc's Gaps section**

Replace:

```
- **Timelines are the remaining case of exactly that hazard.** They survive as untouched parts, and
  deleting the pivot table a timeline filters still leaves it pointing at nothing. PRD 5 task 4
  covers them.
```

with:

```
- **No instance of that hazard is left.** Deleting a pivot table now takes both its slicers and its
  timelines with it — the parts, the workbook registrations, the `#N/A` defined names and the
  drawing anchors — rather than leaving them pointing at nothing. See
  `XLPivotDependentCascade`.
```

Also amend the preceding bullet, which currently says "Slicers are the exception", to say slicers and timelines are.

- [ ] **Step 8: Regenerate the acceptance-check workbooks against the final commit**

`AcceptanceCheckWorkbooks` (Task 3, Step 14) already writes all three, including
`ac7-pivot-deleted-timeline-cascades-<sha>.xlsx` — but the copies on disk were stamped with Task 3's
sha and predate the cascade. Regenerate every one of them from this task's commit so all three stamps
match what shipped, and delete the Task 3 copies so no stale sha is left to be checked by mistake:

```
git rev-parse --short HEAD
```

```
AC_SHA=<that sha> dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/AcceptanceCheckWorkbooks/*"
```

Extend `check-sheet.md` with the `ac7` entry: the workbook opens with no repair prompt, the pivot
table is gone, and no orphan band is drawn where the timeline used to be. A repair prompt here means
one of the five pieces was left behind — check the cache part, the workbook registration, the `#N/A`
defined name, the sheet's `extLst` reference and the drawing anchor, in that order.

- [ ] **Step 9: Commit**

```
git add XLibur/Excel/PivotTables/XLPivotDependentCascade.cs XLibur/Excel/PivotTables/XLPivotTables.cs XLibur.Tests/Excel/PivotTables/PivotTableDeletionTests.cs docs/round-trip-fidelity.md
git commit -m 'fix(timelines): deleting a pivot table takes its timelines with it

A timeline cache left naming a pivot table that is no longer in the
workbook is a repair prompt in Excel, and a timeline whose last pivot
table has gone has nothing left to filter. Both now go with it, along
with the cache part, the workbook registration, the #N/A defined name
and the drawing anchor.

XLSlicerCascade becomes XLPivotDependentCascade, since it is no longer
about slicers. This was the last instance of the dangling-reference
hazard docs/round-trip-fidelity.md named.'
```

---

## After the plan

Report to the owner:

1. **What is automated and green** — the read model, creation, patching, positioning, the cascade, and OpenXML validation on every generated fixture.
2. **What is not settled** — criteria 3, 4 and 7 need a human with Excel. Point at `scratchpad/ac-check-timelines/` and its `check-sheet.md`, and say plainly that a green suite does not settle them: PRD 5's whole finding was that every automated gate passed on a slicer feature Excel refused to render.
3. **The one bet worth naming** — the graphic frame is written bare, with no `mc:AlternateContent` wrapper. If the timeline does not appear in Excel but everything else in the package is correct, that is the first thing to change, and the spec's Risks table gives the exact fallback.
