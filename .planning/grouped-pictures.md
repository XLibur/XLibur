# Grouped pictures — full editable support

Tracking doc for turning grouped drawings (`xdr:grpSp`) from preserve-only into fully editable.

## Background / shipped

- **PR #111** (`fix:`) — preserve grouped pictures and shapes on round-trip (skip-load groups).
- **PR #112** (`feat:`) — load + resize pictures in a single top-level group, update in place on save.
  Internal model: `XLPicture.GroupInfo`/`IsInGroup`, `XLPictureGroup`.

### Key architecture constraints (must hold for every phase)

- The load `SpreadsheetDocument` is disposed after reading; **save re-parses a byte-copy** of the
  original package. Load and save DOMs are independent object graphs → link an in-memory picture to
  its `<xdr:pic>` by **stable `cNvPr` id**, never an object reference. Ids are unique across the whole
  drawing, so this works at any nesting depth.
- Save preserves unmodeled parts because the destination starts as a byte-copy, then is mutated in place.
- `NonVisualDrawingProperties` id rebasing must stay **skipped when a group is present** — it would
  break connector start/end refs (`a:stCxn/@id`).

## API direction (decided)

**Option B — first-class `IXLPictureGroup`.** Exposed via `ws.PictureGroups` and `IXLPicture.Group`;
membership ops (`Add`/`Remove`/`Ungroup`) on the group; `ws.Pictures.Group(p1, p2, …)` to create.
Move/resize stay on `IXLPicture`. Promote internal metadata to public + update `PublicAPI.Unshipped.txt`
(repo enforces the public-API analyzer under `TreatWarningsAsErrors`).

## Foundations (shared)

- **F1 — transform composition util.** Generalize child↔sheet math; support nesting.
  - child → sheet (per level): `sheet = off + (child − chOff)·(ext/chExt)`; `sheetExt = childExt·(ext/chExt)`.
  - sheet → child (inverse, for writing edits): `child = chOff + (sheet − off)·(chExt/ext)`.
  - nested: compose innermost-parent → outermost; effective scale = Π(ext/chExt).
  - store on picture: composed scale + offset, immediate parent group `cNvPr` id, ancestor id chain.

## Non-goals (v1, documented limitations)

- Rotation & flips (`Rotation`, `HorizontalFlip`/`VerticalFlip`) on groups/pictures — ignored.
- Validate each phase's round-trip with `OpenXmlValidator` in addition to structural assertions.

## Phases

Suggested order: 2.1 → 2.2 → 2.3 → 2.4 → 2.5 → 2.6. Each phase = its own branch/PR + fixture + tests.

- [x] **2.1 Nested groups (load + editable)** — branch `feat/grouped-pictures-nested` ✅
  - F1: composed (nesting-aware) scale on `XLPictureGroup` (`ScaleX/ScaleY` now total scale; `GroupId` stored).
  - Loader: `LoadGroupRecursive` walks `grpSp` at any depth, composing `ext/chExt` ratios.
  - Writer: `UpdateGroupedPicture` already keys off composed scale + locates `<xdr:pic>` by id (depth-independent) — no change needed.
  - Tests: `NestedGroupPictures.xlsx` fixture (outer 2× → inner 2×); load composed-scale geometry,
    unedited round-trip preserves both groups/extents/connector, deeply-nested resize round-trips. 6 grouped
    tests + 163 broader suite green.
- [x] **2.2 Moving grouped pictures (within group)** — branch `feat/grouped-pictures-nested` ✅
  - `XLPictureGroup` now carries the composed **affine** (`OffsetX/Y` EMU + `ScaleX/Y`) and load-time
    `LoadedLeftPx/TopPx`. Loader sets an A1-relative `TopLeft` marker so `Left`/`Top`/`MoveTo(l,t)`
    read & write **sheet-space** position for grouped pictures.
  - Writer: `UpdateGroupedPicture` handles size and position independently; on a move it inverts the
    affine to the child `a:off`. Group bbox (`ext`/`chExt`) kept **fixed** (documented decision).
  - Tests: Left/Top reflect sheet position; move round-trips single-level and deeply nested; group +
    sibling + connector preserved. 9 grouped tests + 166 broader suite green.
  - **Deferred:** moving the *whole group* (needs the group object) → folded into 2.6 public API.
- [x] **2.3 Removing a picture from a group** — branch `feat/grouped-pictures-nested` ✅
  - `pic.Delete()` / `ws.Pictures.Delete(name)` now removes a grouped picture for real. `XLPictures`
    routes grouped deletions to `DeletedFromGroups` (keyed by drawing id + rel id) instead of the
    rel-id-only `Deleted` set. `PictureWriter.RemoveGroupedPictures` removes just the matching
    `<xdr:pic>` (by id) and drops the image part only if no remaining blip references it. Group/siblings
    preserved; bbox fixed.
  - Gotcha fixed: the remove pass must be guarded on `Count > 0` so it doesn't materialize (and thus
    re-serialize) the drawing DOM for chart/form-control sheets — that briefly broke
    `PreserveChartsWhenSaving`/`FormControlsArePreserved`.
  - Tests: deleting a grouped picture leaves group + sibling + connector, drops only that image part,
    `Pictures.Count` decremented. 10 grouped tests + 167 broader suite green.
  - `IXLPictureGroup.Remove` alias arrives with the group API in 2.6.
- [x] **2.4 Adding a picture to a group** — branch `feat/grouped-pictures-nested` ✅
  - Internal entry point `XLPictures.AddToGroup(sibling, stream, name)` (full `IXLPictureGroup.Add`
    in 2.6) creates a FreeFloating picture tagged with the target group's composed transform + GroupId,
    marked `IsNew`. `PictureWriter.InsertGroupedPicture` allocates a drawing-wide id (max+1), a new
    image part (rel id via the generator, which already registers the drawing's existing rels so no
    collision), builds the `<xdr:pic>` with child `off`/`ext` from the requested sheet geometry via the
    inverse affine, and appends it to the group. Model is reset to "existing" afterwards for multi-save.
  - Tests: add a picture to a group, set size+position, save, reopen → 3 pictures in the group + 3
    image parts, geometry round-trips, connector/originals intact. 11 grouped tests + 168 broader green.
- [x] **2.5 Creating new groups** — branch `feat/grouped-pictures-nested` ✅
  - Internal `XLPictures.Group(params XLPicture[])` (full API in 2.6) validates members are same-sheet,
    free-floating, in-DOM, ≥2; computes the EMU bbox; tags members grouped (identity child space) and
    records an `XLPendingGroup`. `PictureWriter.CreateGroups` (run before the picture loop) builds the
    `grpSp` (id, `off`/`ext`/`chOff`/`chExt` = bbox) inside an `absoluteAnchor`, moves each member's
    existing `<xdr:pic>` into it with child `off`/`ext` = absolute sheet EMU, and removes the old
    top-level anchors.
  - Tests: two free-floating pictures → save → group → save → reopen yields a single group containing
    both, only the group's anchor at top level, positions preserved. 12 grouped tests + 169 broader green.
  - **Limitation (follow-up):** only free-floating, already-saved pictures can be grouped. Cell-anchored
    members need cumulative column/row→EMU geometry (not yet available); grouping brand-new unsaved
    pictures needs element insertion. Group anchor is `absoluteAnchor` (doesn't move with cells).
- [x] **2.6 Public API + docs** — branch `feat/grouped-pictures-nested` ✅
  - New public `IXLPictureGroup` (`Worksheet`, `Pictures`, `Add`, `Remove`). `IXLPicture` gains
    `IsInGroup` and `Group`. `IXLPictures.Group(params IXLPicture[])` creates a group; `IXLWorksheet`
    gains `PictureGroups`. Backed by `XLPictureGroupView` keyed by a stable per-worksheet
    `XLPictureGroup.GroupKey` (assigned at load/create, survives save unlike the cNvPr id).
  - `PublicAPI.Unshipped.txt` updated for all new symbols. XML docs on every new member.
  - Tests: public API end-to-end (PictureGroups, pic.Group, group.Pictures/Add/Remove, round-trip).
    13 grouped tests + 376 broader (2 pre-existing skips) green.

## Cross-cutting risks

- `cNvPr` id uniqueness on add/create (scan whole drawing incl. nested groups).
- Never renumber connector `a:stCxn`/`a:endCxn` ids.
- Group anchor variants (`oneCell`/`absolute`, `editAs`).
- Image-part sharing when removing.
- DOM re-serialization (namespace hoisting) is expected and lossless; assert semantic equality, not bytes.
