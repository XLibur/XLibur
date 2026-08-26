# Spec 17 — Picture Styling & Round-Trip Fidelity (`IXLPicture`)

**Area:** Feature + Compatibility (and a fidelity **defect fix**)
**Effort:** M–L (2–3 weeks, after spec 16)
**Dependencies:** **Hard dependency on [spec 16](16-drawingml-infrastructure.md)**
(`ShapePropertiesWriter`, `DrawingAnchorFactory`, change-set harness). Shares two shared-layer
extensions with [spec 15](15-shapes-and-text-boxes.md) — see [Interplay](#interplay-with-specs-15-and-16).
Conflicts with spec 15 in `XLibur/Excel/IO/PictureWriter.cs` (both rework its save orchestration);
run them sequentially in either order, not concurrently.
**Status:** Proposed — **unblocked**, spec 16 landed (PRs #401, #402)

## Summary

`IXLPicture` models geometry and placement only — no border, fill, rotation, flip, transparency,
shadow, recolor, crop, or brightness/contrast. But the gap is worse than "missing API": **the picture
save path destroys existing styling.** Every save rebuilds every picture's anchor from scratch with a
hardcoded `spPr` and `blipFill`, so a loaded file whose pictures were rotated, bordered, shadowed,
recolored or cropped in Excel silently loses all of it the first time XLibur saves the workbook. The
"preservation by default" story in `docs/round-trip-fidelity.md` does not apply here, because
preservation only covers what XLibur does *not* rewrite — and pictures are rewritten.

This spec (a) moves loaded pictures to **patch-in-place**, closing the destruction for modelled and
unmodelled styling alike, and (b) adds the styling model: border, fill, rotation, flip, transparency,
shadow, recolor, crop, brightness/contrast.

## Decisions taken before writing this spec (2026-08-01)

| Question | Decision | Consequence |
|---|---|---|
| Loaded-picture write path | **Patch-in-place** (charts/spec-10 and shapes/spec-15 approach), replacing wholesale regeneration | Fixes destruction of *unmodelled* content too (effects, adjust data); needs dirty-tracking; placement changes need an explicit fallback (below) |
| Recolor model | **Preset enum** — `None`, `Grayscale`, `Sepia`, `Washout`, `BlackAndWhite`, plus read-only `Custom` | Matches Excel's gallery; arbitrary duotones read as `Custom` and are preserved, not editable |
| Shadow depth | **Single outer shadow** — colour, transparency, blur, distance, angle (`a:outerShdw`) | Excel's whole preset gallery is parameter values of this one element; inner/perspective shadows stay preserved-unmodelled |
| Extra scope | **Crop (`a:srcRect`) and brightness/contrast (`a:lum`) included** | Both are destroyed today; both are cheap once `blipFill` is patched rather than rebuilt |

## Current state

Verified against the tree at `dfd5f49b` (2026-08-01).

**The model has no styling.** `IXLPicture` (`XLibur/Excel/Drawings/IXLPicture.cs`) exposes geometry,
placement, name, format, group membership. `XLPicture` stores nothing beyond that.

**The reader reads no styling.** `DrawingPartReader.LoadPictureFromAnchor` / `LoadPictureTransform` /
`LoadPicturePlacement` (`DrawingPartReader.cs`) read offsets, extents and markers. `a:xfrm/@rot`,
`@flipH`/`@flipV`, `a:ln`, `a:effectLst`, `a:srcRect`, and every `a:blip` child effect are ignored.

**The writer destroys styling on save.** `PictureWriter.AddPictureAnchor` (`PictureWriter.cs:197-391`)
locates each picture's existing anchor by image rel id (`GetAnchorFromImageId`, line 223) and
**replaces it** (`AttachAnchor`, line 395) with a freshly built anchor whose `spPr` is hardcoded to a
plain `Transform2D` + `PresetGeometry rect` (lines 269-275) and whose `blipFill` is hardcoded to
`Blip + Stretch(FillRectangle)` (lines 261-267). Consequences, all silent:

- rotation, flips → gone (`a:xfrm` rebuilt without `@rot`/`@flipH`/`@flipV`);
- border, shadow, any effect → gone (`a:ln`/`a:effectLst` never emitted);
- crop → gone (`a:srcRect` replaced by `Stretch`);
- transparency/recolor/brightness → gone (`a:blip` rebuilt bare);
- the image binary is re-fed on every save (`FeedData`, line 220) even when untouched.

The only pictures spared are grouped ones — `UpdateGroupedPicture` patches the `xdr:pic` in place —
and anchors hosting a group (the guard at lines 229-230).

**What exists to build on.** Spec 16 provides the anchor factory, the `ShapePropertiesWriter` (fills,
outlines, ordering — `xdr:pic/spPr` is the same `a:CT_ShapeProperties` as shapes and charts), the
change-set harness with golden fixtures, and `BuildColor` with theme-colour support. Spec 15's design
rules 2/3/8/10 (null = unassigned; per-property assignment flags; edit granularity matches model
granularity; never touch a clean drawing) apply here verbatim and are not restated.

## OOXML background

Everything this spec models lives in two places inside `xdr:pic`:

```xml
<xdr:pic>
  <xdr:nvPicPr>…</xdr:nvPicPr>
  <xdr:blipFill>
    <a:blip r:embed="rId1">
      <a:alphaModFix amt="65000"/>          <!-- transparency: 100% - amt/1000 -->
      <a:grayscl/>                          <!-- recolor: grayscale -->
      <!-- or <a:duotone><a:prstClr val="black"/><a:schemeClr val="…"/></a:duotone> (sepia/washout) -->
      <!-- or <a:biLevel thresh="50000"/> (black & white) -->
      <a:lum bright="20000" contrast="-10000"/>  <!-- brightness/contrast, thousandths of a % -->
    </a:blip>
    <a:srcRect l="10000" t="0" r="25000" b="0"/> <!-- crop, thousandths of a % per edge -->
    <a:stretch><a:fillRect/></a:stretch>
  </xdr:blipFill>
  <xdr:spPr>
    <a:xfrm rot="1200000" flipH="1">…</a:xfrm>   <!-- rotation in 60,000ths of a degree -->
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill>…</a:solidFill>                 <!-- fill: shows behind transparent image regions -->
    <a:ln w="19050">…</a:ln>                     <!-- border -->
    <a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" …/></a:effectLst>
  </xdr:spPr>
</xdr:pic>
```

Points that drive the design:

- **`spPr` styling is exactly the shapes/charts problem** — border, fill, shadow and the `a:xfrm`
  attributes go through spec 16's layer (plus the extensions below). Element order in `spPr` and the
  fill choice group are already solved there.
- **`blipFill` styling is new territory** — the `a:blip` child effects are a *sequence of effects*
  applied in order, and Excel is sensitive to their order (`alphaModFix` before recolor before `lum`
  as it writes them). This needs its own small writer with the same in-place discipline.
- **Rotation does not change the anchor.** Excel keeps the anchor/extent of the unrotated picture and
  rotates about the centre; `Width`/`Height`/`Left`/`Top` keep meaning the unrotated box. Same for
  flips. No bounding-box recomputation.
- **Recolor presets are fixed effect patterns**: Grayscale = `a:grayscl`; Black & white =
  `a:biLevel thresh="50000"`; Sepia and Washout = `a:duotone` with specific colour pairs Excel always
  uses. The reader maps those exact patterns back to the enum; any other pattern reads `Custom`.

## Design

### Public surface (additions to `IXLPicture`)

Flat nullable properties with per-property assignment tracking, following `IXLChartSeries` and spec
15's `IXLShape` — null means "not assigned, leave whatever the file has":

```csharp
// Border — same vocabulary as IXLShape
XLColor? BorderColor { get; set; }
double? BorderWidthPt { get; set; }
XLLineDashStyle? BorderDash { get; set; }         // shared enum with spec 15
bool NoBorder { get; set; }

// Fill behind transparent image regions
XLColor? FillColor { get; set; }
bool NoFill { get; set; }

// Orientation
double? Rotation { get; set; }                     // degrees, clockwise; a:xfrm/@rot
bool? FlipHorizontal { get; set; }
bool? FlipVertical { get; set; }

// Image effects
double? Transparency { get; set; }                 // 0–100 %, a:blip/a:alphaModFix
XLPictureRecolor? Recolor { get; set; }            // None|Grayscale|Sepia|Washout|BlackAndWhite; Custom read-only
double? Brightness { get; set; }                   // −100…100 %, a:lum/@bright
double? Contrast { get; set; }                     // −100…100 %, a:lum/@contrast

// Crop — thousandths-of-a-percent per edge in XML, exposed as 0–100 %
XLCrop? Crop { get; set; }                         // readonly record struct: Left, Top, Right, Bottom

// Shadow — single outer shadow
XLPictureShadow? Shadow { get; set; }              // record: Color, TransparencyPct, BlurPt, DistancePt, AngleDeg
```

New types: `XLPictureRecolor` enum, `XLCrop` and `XLPictureShadow` readonly record structs. Setting
`Recolor` to `Custom` throws (`Custom` is a read-back value only). Setting any property marks it
assigned; `null` assignment (e.g. `Rotation = null`) after a prior assignment means "remove the
modelled value on save" for removable elements and is defined per property in the implementation
notes of the PR — the chart precedent (`SetOutline`'s width handling) applies.

`Duplicate`/`CopyTo` copy assigned styling and, for loaded pictures, the unmodelled anchor content
they carry (the copy is a clone of the patched anchor, not a regeneration).

### Write path rework — patch-in-place

`AddPictureAnchor`'s replace-always behaviour is retired for loaded pictures:

- **Clean sheets are never touched** (spec 15's design rule 10, extended to pictures): if no picture
  on the sheet was added, edited, deleted, re-imaged or re-placed, the shape/picture pass does not
  materialise `WorksheetDrawing` and does not re-feed image binaries. This is a behavioural
  *improvement* — today every picture sheet is rewritten and every image re-fed on every save.
- **Loaded, edited pictures are patched**: located by image rel id (the existing
  `GetAnchorFromImageId` mechanism), geometry written into the *existing* anchor's markers/extents,
  `spPr` styling through the shared layer, `blipFill` styling through the new `BlipEffectsWriter`.
  Unassigned properties and unmodelled content are untouched.
- **The image binary is re-fed only when the stream changed** (new-picture or an explicit image
  replacement), not on every save.
- **New pictures** are generated via spec 16's `DrawingAnchorFactory` exactly as today, with assigned
  styling written through the same two writers — one implementation of each write, as in spec 15.
- **Placement change is the documented fallback to regeneration.** Changing `Placement` on a loaded
  picture changes the anchor *element type*, which cannot be patched. Unlike shapes (spec 15 throws),
  pictures have always supported this and callers rely on it — so it stays supported, implemented as
  regeneration of that one anchor with all *modelled* styling carried over, and documented as the one
  path where unmodelled anchor content is lost. A test pins exactly this semantics.
- **`RebaseNonVisualDrawingPropertiesIds` is retired for existing drawings.** Wholesale renumbering
  is incompatible with patch-in-place (and spec 15 already forbids it when shapes exist). It runs
  only when the entire drawing part is new. Behaviour change pinned by test.

### Read path

`LoadPictureFromAnchor` additionally reads, into unassigned model state (visible via the getters,
flagged only when the *user* sets them — mirroring spec 15's loaded-shape read): rotation/flips from
`a:xfrm`; border from `a:ln` (colour incl. `schemeClr`, width, `prstDash`); fill from `spPr`
`solidFill`/`noFill`; transparency from `alphaModFix`; recolor by pattern-matching the known preset
effect shapes (else `Custom`); brightness/contrast from `a:lum`; crop from `a:srcRect`; shadow from a
sole `a:outerShdw` (any other effect content → shadow reads null, is preserved).

### `BlipEffectsWriter` (new, `IO/DrawingML/`)

The `blipFill` counterpart of `ShapePropertiesWriter`: in-place, ordered writes of `a:alphaModFix`,
recolor effects, and `a:lum` inside the existing `a:blip`, plus `a:srcRect` maintenance beside the
existing fill-mode element (never replacing an existing `a:tile`/`a:stretch` choice). Effect order
mirrors what Excel writes, captured from fixtures. Lives in the shared layer because a future spec
(in-cell images, shape picture-fills) can reuse it; until then its only consumer is this spec, so it
is *created here with this spec's tests*, not in spec 16 (16 is extraction-only).

## Interplay with specs 15 and 16

- **From 16 (required):** `ShapePropertiesWriter`, `DrawingAnchorFactory`, harness + goldens.
- **Shared with 15, first-lander adds them:** the `a:noFill`-emission and `a:prstDash` operations in
  the shared layer, and the `XLLineDashStyle` enum. Both specs state this; whichever merges second
  rebases on the other's version.
- **Added here, reusable by shapes later:** the `a:outerShdw` write operation and the
  `a:xfrm` `@rot`/`@flipH`/`@flipV` attribute patching (attribute-only edits on an existing `a:xfrm`,
  offsets untouched). Spec 15 deferred shadow/rotation for shapes; when a shapes v2 picks them up,
  the layer operations already exist.
- **Conflict:** both 15 (task 4) and this spec rework `PictureWriter` orchestration. Sequential, not
  concurrent; either order works, second one rebases.

## Work plan

| # | Task | Size | Depends on |
|---|------|------|-----------|
| 1 | **Patch-in-place rework, no new API** — the fidelity fix alone: clean-sheet no-touch, patch geometry into existing anchors, conditional image re-feed, placement-change regeneration fallback, rebase retirement. Existing picture suite green; new harness-based tests prove an Excel-styled picture round-trips untouched (this fails on main today). | M | spec 16 |
| 2 | **Model + API**: properties, assignment flags, `XLPictureRecolor`/`XLCrop`/`XLPictureShadow`, public API files, `Duplicate`/`CopyTo` semantics. No IO. | S–M | — |
| 3 | **`spPr` styling writes**: border/fill via the shared layer (adding `noFill`/`prstDash` ops if spec 15 has not already), rotation/flip attribute patching, `outerShdw` operation. | M | 1, 2 |
| 4 | **`BlipEffectsWriter`**: transparency, recolor presets, brightness/contrast, crop — in-place and ordered. | M | 1, 2 |
| 5 | **Reader**: all modelled properties incl. recolor pattern-matching to presets. | S–M | 2 |
| 6 | **Tests + fixtures**: Excel-authored pictures covering every property and combinations (rotated+cropped+shadowed), recolor gallery fixtures for pattern-matching, change-set assertions throughout, manual opens-clean-in-Excel checks recorded. | M | 3, 4, 5 |
| 7 | **Docs**: picture styling page in `docs-website/`, and a correction to `docs/round-trip-fidelity.md` — it must state that picture styling was destroyed before this spec and is preserved/modelled after. | S | 6 |

Task 1 is deliberately standalone and PR-able first: it is the defect fix, valuable with zero new API,
and it is the risky part — everything after it is additive.

## Acceptance criteria

1. **The headline fix:** an Excel-authored workbook whose picture is rotated, flipped, bordered,
   shadowed, recolored, cropped and brightness-adjusted survives load → save → reopen with all of it
   intact — byte-identical drawing part when nothing was edited (clean-sheet rule), canonical-identical
   `xdr:pic` when only sheet content elsewhere changed. **This fails on main today** and the test
   asserting it is the spec's reason to exist.
2. Geometry-only edit (`MoveTo`/`WithSize`) on a styled loaded picture → change set is exactly the
   anchor markers/extents (+ `a:xfrm` off/ext); every styling element untouched.
3. Each modelled property: set via API → save (`validate: true`, zero errors) → reload → reads back;
   the full matrix opens clean in Excel 365 (manual check recorded in the PR).
4. Recolor: the four presets written by XLibur are element-identical to Excel's gallery output
   (fixture comparison); Excel-authored presets read back as the right enum member; an arbitrary
   duotone reads `Custom`, cannot be assigned, and survives save untouched.
5. Placement change on a loaded picture: regenerates that anchor, carries all modelled styling,
   drops unmodelled content — pinned by an explicit test and stated in the docs.
6. Image binaries are not re-fed for pictures whose stream did not change; unedited picture sheets
   are not rewritten at all. Write benchmark shows no regression; a save of a loaded
   many-pictures workbook with no edits gets measurably cheaper (record the numbers in the PR).
7. Full suite green on net8.0/net9.0/net10.0; grouped-picture and group-creation tests unmodified.

## Risks

- **Task 1 changes shipped save behaviour for every workbook with pictures.** Anchors that were
  regenerated (normalized) every save are now passed through. Files that depended on regeneration to
  *repair* odd anchors — if any exist — would surface as regressions. The existing picture suite plus
  golden fixtures over the picture test resources are the gate; run the full
  `TryToLoad` corpus through load-save as a sweep.
- **Anchor-form geometry patching.** Writing `MoveTo` into an existing `TwoCellAnchor` means
  recomputing from/to markers against current column widths/row heights — the regeneration path did
  this from scratch; the patch path must produce the same markers. Reuse the exact marker computation,
  assert equivalence in tests.
- **`a:blip` effect ordering.** The effect sequence is order-sensitive and underdocumented; mirror
  Excel-authored fixtures, and keep unrecognized effect elements exactly where they are relative to
  the ones being edited.
- **Recolor pattern-matching brittleness.** Sepia/Washout duotone colour pairs must match what current
  Excel writes; older files may differ slightly and will read `Custom` — acceptable (preserved), but
  document it.
- **Rotation and `editAs` interplay.** A rotated picture in a `TwoCellAnchor` whose row/column sizes
  change renders differently across anchor semantics; XLibur does not recompute anything — same
  posture as Excel, stated in docs.

## Out of scope for v1

Preset 3-D/bevel/glow/soft-edge effects, artistic effects (`a14:imgLayer`), picture styles gallery
(`xdr:style` on pictures), inner/perspective shadows (preserved, not modelled), tile fill mode
authoring (`a:tile` preserved if present), image replacement API (re-point `ImageStream` exists
already; a richer API is separate), per-picture hyperlinks (`a:hlinkClick` — preserved), and
compression/DPI-resampling of image binaries.
