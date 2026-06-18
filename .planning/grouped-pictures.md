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
- [ ] **2.4 Adding a picture to a group**
  - `IXLPictureGroup.Add(stream/file, pos, size)`; new `cNvPr` id (drawing-wide max+1), new image part+rel,
    child off/ext via F1 inverse, insert `<xdr:pic>` into group; optional bbox grow.
- [ ] **2.5 Creating new groups** (most invasive)
  - `ws.Pictures.Group(params IXLPicture[])`; bbox → group off/ext, chOff/chExt = off/ext (scale 1);
    convert members into group children; remove their original top-level anchors; build `grpSp` + anchor.
- [ ] **2.6 Public API + docs**
  - Promote `IsInGroup`/`Group`/`IXLPictureGroup`/`ws.PictureGroups`; `PublicAPI.Unshipped.txt`; XML docs.

## Cross-cutting risks

- `cNvPr` id uniqueness on add/create (scan whole drawing incl. nested groups).
- Never renumber connector `a:stCxn`/`a:endCxn` ids.
- Group anchor variants (`oneCell`/`absolute`, `editAs`).
- Image-part sharing when removing.
- DOM re-serialization (namespace hoisting) is expected and lossless; assert semantic equality, not bytes.
