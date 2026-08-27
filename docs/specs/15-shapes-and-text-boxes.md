# Spec 15 — Shapes & Text Boxes (DrawingML `xdr:sp`)

**Area:** Feature + Compatibility
**Effort:** L (2–3 weeks, after spec 16)
**Dependencies:** **Hard dependency on [spec 16](16-drawingml-infrastructure.md)** — the anchor
factory, the shared `ShapePropertiesWriter`, the text-body primitives, and the XML change-set harness
all come from there and must land first. Task 4 here modifies `PictureWriter.WriteDrawings`
orchestration, so it conflicts with any in-flight work in `XLibur/Excel/IO/PictureWriter.cs`.
**Status:** Proposed — **unblocked**, spec 16 landed (PRs #401, #402)

## Summary

XLibur cannot create a drawing shape of any kind. `IXLPictures.Add` takes images only, `IXLCharts.Add`
takes charts, and there is no third option — a caller who wants a floating text box, a callout, or an
arrow on a worksheet has to drop out of XLibur and hand-write `xdr:sp` into the saved package. Shapes
that already exist in a loaded file survive a round trip verbatim, but they are invisible to the model:
their text cannot be read, let alone changed.

This spec adds a first-class shape model — `ws.Shapes` — covering preset geometries that can carry
text, with a text box being the rectangle preset. Shapes read from a file are loaded into the model and
**patched in place** on save, so everything XLibur does not model (effects, 3-D, gradient fills,
hyperlinks, adjust handles) keeps surviving untouched.

## Decisions taken before writing this spec

| Question | Decision | Consequence |
|---|---|---|
| Object-model width | **Preset shapes, all text-capable**; text box is the rectangle preset | One writer, one anchor path, no breaking rename when ellipses/callouts arrive |
| Existing shapes in loaded files | **Read into the model and patch in place** (spec 10's chart approach) | Highest fidelity; needs dirty-tracking and its own fixture corpus |
| Text content model | **Paragraph-aware** (`IXLTextBody` → paragraphs → runs), not `IXLRichText` | True mapping of `a:txBody`; per-paragraph alignment expressible; a second rich-text-shaped model to maintain |
| Formatting depth in v1 | **Core set** — solid/no fill, line colour/width/dash, run fonts, horizontal + vertical alignment, wrap, picture-style anchoring | Rotation, auto-fit, insets, gradients, shadows deferred (see [Out of scope](#out-of-scope-for-v1)) |
| Shared infrastructure (2026-08-01 review) | **Split into [spec 16](16-drawingml-infrastructure.md)**, landing first; property-layer extraction is **full** (setters move, value-only signatures), not primitives-only | This spec consumes; it does not refactor `ChartFormatting`/`AddPictureAnchor` itself |
| Run font API (2026-08-01 review) | **Narrow `IXLShapeTextFont`**, not `IXLFontBase` — the spreadsheet-font contract has members (`Shadow`, `FontFamilyNumbering`, `FontCharSet`, `FontScheme`, accounting underlines) with no clean `a:rPr` mapping | Honest per-property assignment tracking; not the familiar base type cells/comments use |
| Creation defaults (2026-08-01 review) | **Mirror Excel's authored XML** — capture from fixtures, emit as static boilerplate | New shapes look native in Excel; slightly more writer output |
| Spec-16 review follow-ups (2026-08-01) | Text-body primitives are extracted **here, in task 2** (not in 16 — abstraction shaped by the real consumer); task 3 **extends** the shared layer with `a:noFill`/`a:prstDash`, operations charts never perform | 16 stays strictly behaviour-preserving; the new operations are gated by this spec's change-set tests |

## Current state

Verified against the tree at `dfd5f49b` (2026-08-01).

**No creation path.**
- `XLibur/Excel/Drawings/IXLPictures.cs:10-20` — six `Add` overloads, all image sources.
- `XLibur/Excel/Drawings/IXLDrawing.cs` — despite the general name, `IXLDrawing<T>` is the **VML**
  styling model for comments/notes (`XLibur/Excel/IO/VmlDrawingPartWriter.cs:160`,
  `DrawingPartReader.LoadTextBox` at `DrawingPartReader.cs:283`). It has nothing to do with DrawingML
  and must not be extended to cover worksheet shapes.
- `IXLWorksheet` exposes `Charts` (`XLibur/Excel/IXLWorksheet.cs:393`) and `Pictures` (line 536). No
  third drawing collection.

**Existing shapes are preserved but opaque.**
- Read: `DrawingPartReader.LoadDrawings` (`DrawingPartReader.cs:55-88`) walks
  `WorksheetDrawing.ChildElements`; an anchor with no image rel id is skipped outright
  (`DrawingPartReader.cs:74-77`, comment: *"we're probably dealing with a TextBox (or another shape)"*).
- Save: `PictureWriter.RemoveEmptyDrawingPart` (`PictureWriter.cs:127-148`) deliberately keeps the
  DrawingsPart alive when `WorksheetDrawing.HasChildren`, precisely so non-picture shapes are not
  dropped. Save reopens the original package and rewrites only modelled parts, so the shape XML is
  carried through unmodified (the mechanism `docs/round-trip-fidelity.md` documents).
- Regression coverage exists: `XLibur.Tests/Excel/Drawings/PictureTests.cs:41-69` loads
  `TryToLoad/textbox_shapemissing_onload_2377.xlsx` and asserts the `OneCellAnchor` and its `TEXTBOX`
  text survive save. **This test must stay green** — it is the contract this spec must not regress.

**The anchor machinery already exists and is reusable.**
- `PictureWriter.AddPictureAnchor` (`PictureWriter.cs:197-391`) builds all three anchor forms —
  `AbsoluteAnchor` (FreeFloating), `TwoCellAnchor` (MoveAndSize), `OneCellAnchor` (Move) — from
  `XLPicturePlacement` plus pixel geometry, converting through `ConvertToEnglishMetricUnits`. The
  marker/extent construction is identical for a shape; only the `xdr:pic` child differs.
- Non-visual drawing property ids are allocated as `max(existing) + 1` (`PictureWriter.cs:236-239`).
- `RebaseNonVisualDrawingPropertiesIds` (`PictureWriter.cs:755`) renumbers every id in the drawing. It
  is already skipped when the drawing contains a group shape, because renumbering breaks connector
  `a:stCxn/@id` references (`PictureWriter.cs:33-38`). The same hazard applies to shapes.

**Precedent to copy for patch-in-place.**
`XLibur/Excel/IO/ChartPatcher.cs:11-31` states the contract spec 10 settled on: XLibur never regenerates
XML it read, edits are patched into the loaded DOM, and only properties the caller actually assigned are
written (`XLChartSeries.AssignedFormat`, `XLChartSeries.cs:156-198`). Chart formatting properties are
nullable — `XLColor? FillColor`, `double? LineWidthPt` (`IXLChartSeries.cs:51-67`) — with null meaning
"not assigned, don't emit". Shapes follow the same two rules.

## OOXML background

A shape anchor holds an `xdr:sp` instead of an `xdr:pic`:

```xml
<xdr:twoCellAnchor>
  <xdr:from>…</xdr:from><xdr:to>…</xdr:to>
  <xdr:sp macro="" textlink="">
    <xdr:nvSpPr>
      <xdr:cNvPr id="2" name="TextBox 1"/>
      <xdr:cNvSpPr txBox="1"/>            <!-- txBox="1" marks it a text box -->
    </xdr:nvSpPr>
    <xdr:spPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="571500"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill>   <!-- or <a:noFill/> -->
      <a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>
                     <a:prstDash val="dash"/></a:ln>
    </xdr:spPr>
    <xdr:txBody>
      <a:bodyPr vertOverflow="clip" wrap="square" anchor="ctr" rtlCol="0"/>
      <a:lstStyle/>
      <a:p>
        <a:pPr algn="ctr"/>
        <a:r><a:rPr lang="en-US" sz="1100" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
             <a:latin typeface="Calibri"/></a:rPr><a:t>Hello</a:t></a:r>
      </a:p>
    </xdr:txBody>
  </xdr:sp>
  <xdr:clientData/>
</xdr:twoCellAnchor>
```

Points that drive the design:

- **A text box is not a distinct element.** It is `prstGeom prst="rect"` plus `cNvSpPr/@txBox="1"`.
  Excel's Insert ▸ Text Box writes exactly that, and also defaults to `<a:noFill/>` on `spPr` for
  *shapes* but a white fill for text boxes drawn from the ribbon — the writer must emit explicitly what
  the model says rather than relying on defaults, because theme style inheritance (`xdr:style`) is not
  modelled.
- **Geometry lives in two places.** The anchor (`from`/`to` markers, or `pos`/`ext`) positions the shape
  on the sheet; `spPr/a:xfrm` positions it in the shape's own space. For an ungrouped shape Excel writes
  `off x=0 y=0` and `ext` equal to the anchor extent — mirroring what `AddPictureAnchor` already does
  for pictures (`PictureWriter.cs:269-275`).
- **Text formatting is per-run**, in `a:rPr` — a DrawingML font description (`sz` in hundredths of a
  point, `b`/`i`/`u`, `a:solidFill`, `a:latin typeface`), *not* the spreadsheet `<font>` element.
  Comments reuse `IXLFontBase` legitimately because comment rich text uses the spreadsheet font
  schema; `a:rPr` does not, which is why runs get their own narrow font interface (below) instead.
- **`a:ln/@w` is EMU**, not points: `w = pt * 12700`.
- **`a:prstDash` values** (`solid`, `dot`, `dash`, `lgDash`, `dashDot`, `lgDashDot`, `lgDashDotDot`,
  `sysDash`, `sysDot`, `sysDashDot`, `sysDashDotDot`) are DrawingML's own set. The existing
  `XLDashStyle`/`XLLineStyle` enums are **VML** enums used by comment styling
  (`EnumConverter.cs:359-366`, `PublicAPI.Shipped.txt:662-674`); reusing them would force a lossy
  round trip. This spec adds `XLLineDashStyle` mapping 1:1 onto `a:prstDash`.

## Design

### Public surface

New folder `XLibur/Excel/Drawings/Shapes/`. Everything below goes into `PublicAPI.Unshipped.txt`.

```csharp
// Collection — reachable as IXLWorksheet.Shapes
public interface IXLShapes : IEnumerable<IXLShape>
{
    int Count { get; }

    IXLTextBox AddTextBox(string text, IXLCell topLeft);
    IXLTextBox AddTextBox(string text, int left, int top);        // free-floating, pixels
    IXLShape   Add(XLShapeType type, IXLCell topLeft);
    IXLShape   Add(XLShapeType type, int left, int top);
    IXLShape   AddPreset(string presetGeometry, IXLCell topLeft); // escape hatch: raw a:prstGeom/@prst

    IXLShape Shape(string name);
    bool TryGetShape(string name, out IXLShape? shape);
    bool Contains(string name);
    void Delete(string name);
    void Delete(IXLShape shape);
}

public interface IXLShape
{
    int Id { get; }                       // xdr:cNvPr/@id
    string Name { get; set; }
    XLShapeType ShapeType { get; }
    string PresetGeometry { get; }        // the raw prst value, for shapes added via AddPreset
    IXLWorksheet Worksheet { get; }

    // Geometry — pixels, matching IXLPicture
    int Left { get; set; }
    int Top { get; set; }
    int Width { get; set; }
    int Height { get; set; }
    XLPicturePlacement Placement { get; set; }
    IXLCell TopLeftCell { get; }
    IXLCell BottomRightCell { get; }

    // Core formatting — null means "not assigned, leave whatever is in the file"
    XLColor? FillColor { get; set; }
    bool NoFill { get; set; }
    XLColor? LineColor { get; set; }
    double? LineWidthPt { get; set; }
    XLLineDashStyle? LineDash { get; set; }
    bool NoLine { get; set; }

    IXLTextBody TextBody { get; }

    IXLShape MoveTo(int left, int top);
    IXLShape MoveTo(IXLCell cell);
    IXLShape MoveTo(IXLCell cell, int xOffset, int yOffset);
    IXLShape MoveTo(IXLCell fromCell, IXLCell toCell);
    IXLShape WithSize(int width, int height);
    IXLShape WithPlacement(XLPicturePlacement value);
    void Delete();
}

/// <summary>A shape whose geometry is <c>rect</c> and whose <c>cNvSpPr/@txBox</c> is 1.</summary>
public interface IXLTextBox : IXLShape
{
    /// <summary>Whole-body text; paragraphs joined by <c>\n</c>. Setting it replaces the body.</summary>
    string Text { get; set; }
}
```

Text body, paragraph-aware:

```csharp
public interface IXLTextBody : IEnumerable<IXLTextParagraph>
{
    int Count { get; }
    IXLTextParagraph this[int index] { get; }
    IXLTextParagraph AddParagraph(string text = "");
    void Clear();

    string Text { get; set; }                          // convenience over the paragraph list
    XLTextVerticalAlignment VerticalAlignment { get; set; }   // a:bodyPr/@anchor
    bool WrapText { get; set; }                               // a:bodyPr/@wrap (square | none)
}

public interface IXLTextParagraph : IEnumerable<IXLTextRun>
{
    XLTextAlignment Alignment { get; set; }            // a:pPr/@algn — l | ctr | r | just | dist
    IXLTextRun AddText(string text);                   // returns the run, so formatting chains
    string Text { get; }                               // concatenated run text
    int Count { get; }
    IXLTextRun this[int index] { get; }
}

/// <summary>
/// Font properties of a DrawingML text run. Deliberately narrower than <see cref="IXLFontBase"/>:
/// every member here maps cleanly onto <c>a:rPr</c>, and nothing is promised that DrawingML cannot
/// express (no <c>Shadow</c> bool, no accounting underlines, no <c>FontFamilyNumbering</c>/
/// <c>FontCharSet</c>/<c>FontScheme</c>). All nullable: null = not assigned, inherit/preserve.
/// </summary>
public interface IXLShapeTextFont
{
    bool? Bold { get; set; }                           // a:rPr/@b
    bool? Italic { get; set; }                         // a:rPr/@i
    XLShapeTextUnderline? Underline { get; set; }      // a:rPr/@u — None | Single | Double
    bool? Strikethrough { get; set; }                  // a:rPr/@strike
    double? FontSize { get; set; }                     // a:rPr/@sz, hundredths of a point in XML
    XLColor? FontColor { get; set; }                   // a:rPr/a:solidFill
    string? FontName { get; set; }                     // a:rPr/a:latin/@typeface
}

public interface IXLTextRun : IXLShapeTextFont
{
    string Text { get; set; }
    IXLTextRun SetBold(bool value = true);             // fluent setters mirroring the properties
    IXLTextRun SetItalic(bool value = true);
    IXLTextRun SetFontSize(double value);
    IXLTextRun SetFontColor(XLColor value);
    IXLTextRun SetFontName(string value);
}
```

New enums: `XLShapeType` (curated: `TextBox`, `Rectangle`, `RoundedRectangle`, `Ellipse`, `Triangle`,
`RightArrow`, `LeftArrow`, `UpArrow`, `DownArrow`, `Line`, `RoundedCallout`, `WedgeRectCallout`,
`Diamond`, `Pentagon`, `Chevron`, `Star5`, `Can`, `Cloud`, `Custom`), `XLLineDashStyle`,
`XLTextAlignment`, `XLTextVerticalAlignment`, `XLShapeTextUnderline`.

Name semantics follow `XLPictures` exactly: lookups are case-insensitive, `Delete(string)` removes
*all* matches, and names are not required to be unique — loaded files legitimately contain duplicates
(Excel permits them; copy-paste produces them). `Shape(name)` returns the first match in drawing
order; code that needs a specific shape among duplicates enumerates and matches on `Id`.

Usage:

```csharp
ws.Shapes.AddTextBox("Draft — do not circulate", ws.Cell("B2"))
         .WithSize(220, 60);

var callout = ws.Shapes.Add(XLShapeType.WedgeRectCallout, ws.Cell("E4"));
callout.FillColor = XLColor.LightYellow;
callout.LineColor = XLColor.Black;
callout.LineWidthPt = 1.5;
callout.TextBody.VerticalAlignment = XLTextVerticalAlignment.Center;
var p = callout.TextBody.AddParagraph();
p.Alignment = XLTextAlignment.Center;
p.AddText("Revised ").SetBold();
p.AddText("2026-08-01");
```

### Design rules

1. **`XLPicturePlacement` is reused, not duplicated — for creation.** The three placements map to
   exactly the same three anchor forms, and a parallel `XLShapePlacement` with identical members would
   be a second name for one concept. The type name is a wart inherited from the picture API; it is not
   worth a breaking rename. Document it on `IXLShape.Placement`. Two limits, both consequences of
   patch-in-place: (a) on read, the anchor form maps to a placement the way `LoadPicturePlacement`
   does for pictures (`DrawingPartReader.cs:524-543`), but `twoCellAnchor/@editAs`
   (`oneCell`/`absolute`) has no representation in the enum and is simply preserved untouched; (b)
   **setting `Placement` on a loaded shape throws `NotSupportedException` in v1** — changing placement
   means replacing the anchor *element type*, which is anchor regeneration, not patching, and would
   discard `editAs` and anything else on the anchor XLibur does not model. Geometry edits stay within
   the existing anchor form.
2. **Null means unassigned.** Following `IXLChartSeries`, a `null` formatting property is never written.
   A shape loaded from a file whose fill nobody touched keeps whatever fill XML it had — including
   gradients and pictures fills this spec does not model. `NoFill`/`NoLine` are the explicit way to say
   "remove it".
3. **Assignment is tracked, per property.** `XLShape` carries an `AssignedProperties` flags enum, set by
   the setters, mirroring `XLChartSeries.AssignedFormat`. The patcher writes only flagged properties.
   A shape nobody edited is not touched at all.
4. **New shapes are generated; loaded shapes are patched.** Never regenerate a loaded `xdr:sp`.
5. **`AddPreset` is the escape hatch.** DrawingML has ~190 preset geometries; the enum covers the ones
   people ask for and `AddPreset("moon")` covers the rest. `ShapeType` reads back `Custom` for those,
   with the raw string on `PresetGeometry`.
6. **Text boxes are shapes.** `AddTextBox` is `Add(XLShapeType.TextBox, …)` plus `txBox="1"`;
   `IXLTextBox` adds only the `Text` convenience. A shape loaded with `txBox="1"` materialises as
   `IXLTextBox`.
7. **Creation defaults mirror Excel's own authored XML.** "Null is never written" governs *edits to
   loaded shapes*; a *new* shape with nothing assigned must not emit bare `prstGeom` — with no fill,
   no line and no `xdr:style`, rendering is renderer-defined and the shape can be invisible. Instead
   the writer emits, as static boilerplate captured from Excel-authored fixtures: for **text boxes**,
   what Insert ▸ Text Box writes (white/`lt1` fill, thin black line, plus its `xdr:style` block); for
   **preset shapes**, the ribbon's default (`xdr:style` referencing theme accent1 fills/lines, no
   explicit fill in `spPr`). The exact XML is copied from the fixture during task 4, not paraphrased
   from memory. Assigned properties then override the boilerplate through the normal path. Model
   read-back of an untouched new shape reports the defaults as unassigned (`null`), matching how a
   theme-styled loaded shape reads.
8. **Edit granularity matches model granularity.** Every setter mutates the narrowest element that
   holds its value and touches nothing else: `IXLTextRun.Text` rewrites `a:t` and leaves that run's
   `a:rPr` (theme font, `lang`, per-run effects) exactly as it was; `IXLTextParagraph.Alignment` writes
   `a:pPr/@algn` and no sibling; `LineWidthPt` sets `a:ln/@w` on the *existing* `a:ln`, keeping its
   `a:headEnd`/`a:tailEnd`/`a:miter` children. Only `IXLTextBody.Text` (the wholesale setter) and
   `Clear()` rebuild the paragraph list, and both are documented as destructive — that is what the
   caller asked for. This rule is what makes [failure mode 2](#the-three-ways-patching-loses-data)
   impossible by construction rather than by test, and it is the rule that has to hold as v2 adds
   rotation, insets and effects. (Referenced below as "design rule 8".)
9. **DrawingML property writes go through the shared layer.** Shapes do not get their own copy of
   fill/line/ordering logic (spec 16's `ShapePropertiesWriter`, extended by task 3), and text edits go
   through the `TextBodyPrimitives` task 2 extracts (`endParaRPr` stays last, `rPr`-preserving run
   replacement). Shape code never manipulates `spPr` children or `a:p` run-level structure directly.
10. **Save never touches a clean drawing.** The OpenXML SDK re-serializes a part merely because its DOM
   was materialised — stated in the codebase itself at `PictureWriter.cs:138-140`. So on save, if a
   sheet has no added, edited or deleted shape, shape code must not access `WorksheetDrawing` at all;
   the part passes through byte-identical. This also improves the current behaviour:
   `RemoveEmptyDrawingPart` today probes `WorksheetDrawing.HasChildren` even on shape-only sheets, and
   with shapes modelled, `Shapes.Count > 0` short-circuits that probe before it can materialise the
   DOM.

### The three ways patching loses data

Patch-in-place buys fidelity, and the whole of it can be given back by a careless writer. Naming the
failure modes, because tasks 4–7 exist to close them and acceptance criterion 3 exists to prove it:

1. **Sibling loss.** Edit the text, lose the shadow. Cheap to avoid — a patcher that only touches
   assigned properties never goes near `spPr` when the edit was to `txBody` — and correspondingly cheap
   to *test*, which is the trap: a naive implementation passes a sibling-level assertion. This is the
   least of the three and must not be mistaken for the gate.
2. **In-place loss.** The damage happens inside the element being edited: rebuilding an `a:p` to change
   a word discards the run's `a:rPr`; rebuilding `a:ln` to set a width drops the arrowheads that were
   inside it. Design rule 8 is the answer; no sibling-level test detects it.
3. **Structural invalidity.** `CT_ShapeProperties` is a *sequence* (`a:xfrm`, geometry, fill, `a:ln`,
   `a:effectLst`, 3-D, `a:extLst`) and the SDK does not reorder children on `Append`. Fills are a
   *choice* group, so setting a colour over an `a:gradFill` means replacing it, not adding beside it.
   Get either wrong and Excel shows a repair prompt. `validate: true` (criterion 5) catches most of it.

### The shared DrawingML layer (spec 16)

All three failure modes are already solved once in the codebase, in chart-specific form, and
**[spec 16](16-drawingml-infrastructure.md) extracts the property layer before this spec starts**: the
fill/outline/ordering logic from `ChartFormatting.cs:1251-1348` (plus `BuildColor`/`MapSchemeColor`,
which already write theme colours) becomes `DrawingML/ShapePropertiesWriter` — fills as a choice
group, `a:ln` mutated in place, sequence order via `InsertAfterLastOf` and its predicates;
`xdr:sp/spPr` and a chart's `c:spPr` are the same schema type, `a:CT_ShapeProperties`. The layer takes
plain values, never flags masks — this spec's `AssignedProperties` enum stays in its own patcher,
which decides *whether* to write; the layer decides *how*.

Two pieces deliberately land in *this* spec, not 16:

- **Text-body primitives (task 2).** The paragraph-editing rules inside `SetRichText`
  (`ChartFormatting.cs:257-287` — `rPr`-preserving run replacement, `endParaRPr`-stays-last,
  `IsRunLevel`) are extracted into `DrawingML/TextBodyPrimitives` here, where their second consumer
  actually exists to shape the abstraction; the chart title path is moved onto them in the same PR,
  with the chart tests still gating.
- **New operations (task 3).** Charts never emit an explicit `a:noFill` and never write `a:prstDash`,
  so `NoFill`/`NoLine`/`LineDash` are *extensions* of the shared layer, not extractions — added here
  under this spec's change-set tests, keeping 16 strictly behaviour-preserving. (Shared with
  [spec 17](17-picture-styling.md), which needs the same two operations and the `XLLineDashStyle`
  enum for picture borders — whichever spec lands first adds them, the other rebases.)

### Read path

New `XLibur/Excel/IO/ShapeReader.cs`, called from `DrawingPartReader.LoadDrawings` at the point where
the `imgId == null` branch currently `continue`s (`DrawingPartReader.cs:74-77`):

- Anchor contains `xdr:sp` → build an `XLShape`, capture geometry from the anchor (markers/extent →
  pixels), preset from `spPr/a:prstGeom/@prst`, `txBox` flag, fill/line where they are the simple forms
  this spec models (`a:solidFill` holding `a:srgbClr` *or* `a:schemeClr`, `a:noFill`, `a:ln` with
  `w`/`prstDash`), and the text body from `a:txBody`. Anything else — theme colours, gradients, effects — is left unread and unflagged, so
  it is preserved by rule 2.
- **The shape must record its identity, not a DOM handle.** The load package is disposed when
  `LoadSheets` returns (`XLibur/Excel/XLWorkbook_Load.cs:37`) and save *reopens* the package
  (`docs/round-trip-fidelity.md:11`), so an `OpenXmlElement` captured at read time is attached to a
  closed package and is useless to the patcher. `XLShape` therefore stores the `cNvPr/@id` it was read
  with, and `ShapePatcher` re-locates the `xdr:sp` in the reopened drawing by that id — the same way
  `ChartPatcher.ResolvePart` and `PictureWriter.GetAnchorFromImageId` re-find their targets. This makes
  id stability load-bearing, which is a second reason the rebase guard below is not optional.
- **Duplicate ids make a shape read-only.** The schema requires `cNvPr/@id` uniqueness per drawing,
  but real files (copy-paste bugs, other producers) violate it, and an id-keyed patcher would then
  edit the wrong shape. On load, if an id occurs more than once, every shape carrying it is loaded
  for reading but marked unpatchable: any setter on it throws `InvalidOperationException` naming the
  duplicate id. Deleting such a shape is also refused. Rare, explicit, and safe beats silently
  corrupting someone's drawing.
- On read, the anchor form maps to a `Placement` exactly as `LoadPicturePlacement` does for pictures
  (`DrawingPartReader.cs:524-543`); `editAs` and other unmodelled anchor attributes are untouched.
- Nothing is re-serialized on read; `WorksheetDrawing` is already materialised for every sheet with a
  drawing part today.
- Anchors holding `xdr:cxnSp` (connectors), `xdr:graphicFrame`, or `xdr:grpSp` keep the current
  behaviour: not loaded, preserved verbatim. Shapes *inside* a group are out of scope for v1 (as
  grouped pictures were, initially).

### Write path

New `XLibur/Excel/IO/ShapeWriter.cs` + `ShapePatcher.cs`, orchestrated from `PictureWriter.WriteDrawings`
(renaming that entry point to `DrawingWriter.WriteDrawings` is optional; keep the call order —
deletions, then pictures, then shapes, then `RemoveEmptyDrawingPart`).

- **Clean sheets are never touched** (design rule 10): the shape pass runs only when the sheet has an
  added, edited or deleted shape, so an unedited drawing part passes through byte-identical.
- **New shapes** (no loaded id): build the anchor via spec 16's `DrawingAnchorFactory`, append
  `xdr:sp` with the Excel-mirroring boilerplate of rule 7, allocate the id as `max(existing) + 1`.
  Assigned properties are then written by the *same* shared layer the patcher uses, so there is one
  implementation of "write a fill", not two.
- **Loaded shapes**: `ShapePatcher` re-locates the `xdr:sp` by its recorded id and writes only assigned
  properties — geometry into the anchor markers and `a:xfrm`, fill/line through spec 16's
  `ShapePropertiesWriter`, and text at run/paragraph granularity per design rule 8 using the
  text-body primitives. A body is rebuilt only when `IXLTextBody.Text` or `Clear()` was called.
- **Deleted shapes**: remove the whole anchor, matching `ProcessDeletedPicturesAndGroups`.
- An empty text body is written the way Excel writes it — `a:bodyPr`, `a:lstStyle`, and one `a:p`
  holding only an `a:endParaRPr` — never omitted and never an `a:p` with a zero-length run.
- **`RebaseNonVisualDrawingPropertiesIds` must not run when the drawing contains any `xdr:sp`.** The
  existing guard (`PictureWriter.cs:33-38`) only checks for group shapes; extend it. Renumbering ids
  under a connector that references a shape by id produces a file Excel repairs.
- `RemoveEmptyDrawingPart` gains `xlWorksheet.Shapes.Count == 0` to its condition, and the DrawingsPart
  is created on demand for a sheet whose first drawing is a shape.

### Where the model lives

`XLWorksheet` gains `Shapes` alongside `Pictures`/`Charts`, constructed in the same place, backed by
`XLShapes : IXLShapes` holding a `List<XLShape>` plus a `Deleted` list — the `XLPictures` shape exactly
(`XLibur/Excel/Drawings/XLPictures.cs:10-41`). Shapes take no part in row/column shifting in v1, matching
pictures.

## Work plan

Spec 16 (anchor factory, `ShapePropertiesWriter`, text-body primitives, change-set harness) is a
prerequisite for everything below and is planned separately.

| # | Task | Size | Depends on |
|---|------|------|-----------|
| 1 | **Model + collection**: `IXLShape`, `IXLShapes`, `IXLTextBox`, `XLShapeType`, `XLLineDashStyle`, `IXLShapeTextFont`, geometry/placement, `ws.Shapes` wiring, public API files. No IO. | M | — |
| 2 | **Text body model + primitives extraction**: `IXLTextBody`/`IXLTextParagraph`/`IXLTextRun` + `a:txBody` ⇄ model mapping (`a:rPr` ⇄ `IXLShapeTextFont`, `sz` in 1/100 pt, `a:latin`), run/paragraph-granular writes per design rule 8. Extracts `DrawingML/TextBodyPrimitives` from `SetRichText` (`ChartFormatting.cs:257-287`) and moves the chart title path onto them in the same PR — chart tests and spec 16's goldens gate the move. | M | 1 |
| 3 | **Core formatting**: map fill, line, alignment and wrap onto spec 16's `ShapePropertiesWriter`, plus the assigned-property flags; **extend** the layer with the operations charts never perform — explicit `a:noFill` emission and `a:prstDash` — under this spec's change-set tests. | S–M | 1 |
| 4 | **`ShapeWriter`** for new shapes + save orchestration (drawing part creation, Excel-mirroring creation boilerplate from fixtures, id allocation, rebase guard, `RemoveEmptyDrawingPart` short-circuit, clean-sheet no-touch rule). | M | 1, 2, 3 |
| 5 | **`ShapeReader`** — load `xdr:sp` anchors into the model, recording each shape's `cNvPr/@id` for re-location, anchor-form → `Placement` mapping, duplicate-id read-only marking. | M | 1, 2, 3 |
| 6 | **`ShapePatcher`** — assigned-property-only writes into loaded shapes, re-located by id. | M | 4, 5 |
| 7 | **Tests + fixture corpus** (see acceptance criteria), built on spec 16's harness. Excel-authored fixtures into `XLibur.Tests/Resource/`. | M | 4, 5, 6 |
| 8 | **Docs**: a `docs-website/docs/shapes.md` page, and an update to `docs/round-trip-fidelity.md` recording that shapes are now modelled and what still passes through opaquely. | S | 7 |

Parallelisation: tasks 1/2/3 are one workstream, splittable by interface boundary. Tasks 4 and 5 are
independent of each other once 1–3 land; 6 needs both.

## Acceptance criteria

1. `ws.Shapes.AddTextBox("Hello", ws.Cell("B2")).WithSize(200, 60)` on a new workbook produces a file
   that **Excel 365 opens with no repair prompt** (manual check recorded in the PR), showing the text at
   the specified geometry.
2. Every member of `XLShapeType` can be added, saved, reloaded by XLibur, and reads back with its type,
   geometry, fill, line and text intact — as a data-driven test over the enum.
3. **Fidelity under edit** — three tiers on spec 16's change-set harness. Tiers 2–3 are stated in
   *canonical* terms, not bytes, deliberately: materialising a part's DOM re-serializes it even with no
   modifications (`PictureWriter.cs:138-140`), so once one shape on a sheet is patched, the whole part
   is rewritten and only canonical comparison is meaningful. Tier 1 alone demands bytes, which design
   rule 10 makes achievable. Tier 3 is the one with teeth; tiers 1 and 2 are passed by implementations
   that fail tier 3.
   1. *Untouched.* Load an Excel-authored fixture with a gradient-filled, shadowed, rotated text box and
      save without edits → drawing part **byte-identical** (the clean-sheet no-touch rule).
   2. *Cross-area edit.* Change the text only → the canonical change set inside `spPr` is **empty**:
      gradient, shadow and rotation all still present and untouched.
   3. *In-area edit* — change set equals *exactly* the intended nodes; nothing else added, dropped or
      reordered. (a) Change one run's text → that run's `a:rPr` is unchanged and its sibling runs are
      untouched. (b) Set `FillColor` on the gradient-filled shape → the `a:gradFill` is *replaced*, not
      duplicated or left beside the new `a:solidFill`; `a:ln` and `a:effectLst` are unchanged; the
      `spPr` child order is still schema-valid. (c) Set `LineWidthPt` on a shape whose `a:ln` carries
      arrowheads → `a:headEnd`/`a:tailEnd` survive.
4. **No regression**: full suite green on net8.0/net9.0/net10.0, and
   `XLibur.Tests/Excel/Drawings/PictureTests.cs:41-69` still passes unmodified — an untouched text box
   in a loaded file still round-trips with its anchor and text intact (that test asserts structure and
   `InnerText`, not bytes; the byte-level guarantee is tier 1 above).
5. Saving with `validate: true` produces zero schema errors for every shape the API can emit, including
   `AddPreset` with an arbitrary valid `prst` value.
6. A drawing containing a connector that references shapes by id round-trips with those ids unchanged
   (the rebase guard).
7. **No cost when unused**: the write benchmark on a workbook with no shapes shows no time or
   allocation regression beyond noise; on a loaded workbook whose shapes nobody edited, no sheet's
   `WorksheetDrawing` DOM is materialised by shape code (design rule 10) and the drawing parts pass
   through byte-identical. Sheets where one shape *was* edited re-serialize as a whole; the unedited
   sibling shapes on them are asserted canonically identical, not byte-identical.
8. Every new shape the API creates carries the Excel-mirroring boilerplate of rule 7 — verified by
   comparing a created text box and a created preset shape against their Excel-authored fixture
   counterparts on the harness, with only geometry, ids and text expected to differ.

## Risks

- **Excel's tolerance for a hand-built `xdr:sp` is narrower than for `xdr:pic`.** `xdr:txBody` without
  `a:bodyPr` triggers a repair prompt in some builds. (An `a:p` with no run is fine — it is what Excel
  itself writes for an empty body, closed by `a:endParaRPr`; see the write path.) The gate is the
  manual open-in-Excel check; mirror byte-for-byte what Excel writes for a freshly inserted text box
  (author a fixture, inspect the part, copy its element order and attribute set — rule 7 makes this
  the source of the creation boilerplate anyway).
- **Theme-derived formatting.** Excel writes an `xdr:style` block referencing theme fills/lines for
  ribbon-inserted shapes. This spec neither reads nor writes `xdr:style`, so a shape whose colour
  comes only from it reports `FillColor == null`. Setting `FillColor` writes an explicit
  `a:solidFill` into `spPr`, which correctly wins over `xdr:style` — and theme *values* are fine on
  the write side, since the layer's `BuildColor` already emits `a:schemeClr` for
  `XLColorType.Theme`. The reader should likewise read `a:solidFill/a:schemeClr` in `spPr` as a
  theme `XLColor`, not just `srgbClr`; the remaining asymmetry (style-block-only colours read as
  null) needs documenting.
- **Text measurement.** Nothing auto-sizes: a shape is exactly the extent it was given, and long text
  overflows or clips per `wrap`. `SixLabors.Fonts` could measure it, but auto-fit is deferred (and the
  library must not be upgraded — see `CLAUDE.md`).
- **Id collisions in real files.** Patching is keyed on `cNvPr/@id`; duplicates (illegal but seen in
  the wild) are handled by the read-only marking in the read path, but the detection itself must be
  cheap — one pass over loaded ids per sheet, no DOM re-walk at save time.
- **Mixed picture+shape sheets change picture behaviour subtly.** Extending the rebase guard means
  sheets holding both a picture and a shape no longer get their nvPr ids renumbered on save (today
  they do — the guard only checks for group shapes). This is a fix, not a regression, but it is an
  observable output change for such files and needs a test pinning it.
- **Group membership.** A shape inside `xdr:grpSp` is not loaded; if a user deletes a *picture* from a
  group the existing code already refuses to dismantle it (`PictureWriter.cs:63-67`). Keep the same
  posture for shapes and state it in the docs rather than half-supporting it.

## Out of scope for v1

Rotation and flips (`a:xfrm/@rot`, `@flipH`, `@flipV`), auto-fit (`spAutoFit`/`normAutofit`), text
insets (`lIns`/`tIns`/`rIns`/`bIns`), vertical/rotated text (`bodyPr/@vert`), gradient/picture/pattern
fills, shadows and other `a:effectLst` content, 3-D, bullets and numbering, hyperlinks on shapes, z-order
manipulation, connectors and their routing, custom geometry (`a:custGeom`), grouping shapes with each
other or with pictures, `CopyTo`/`Duplicate` across sheets, cell-linked shape text (`xdr:sp/@textlink`),
form controls and ActiveX (VML/legacy — a different mechanism entirely), and shape participation in
row/column insert/delete shifting.

Each of these is additive: the model, the assigned-property flags and the patcher are all designed so a
later spec can extend them without changing what v1 emits.
