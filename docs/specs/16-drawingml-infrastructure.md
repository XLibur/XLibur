# Spec 16 — Shared DrawingML Infrastructure

**Area:** Architecture + Refactor
**Effort:** S–M (~1 week)
**Dependencies:** None. Conflicts with any in-flight work in `XLibur/Excel/IO/ChartFormatting.cs`
(spec 10's territory) and `XLibur/Excel/IO/PictureWriter.cs`. Spec 10's open 3D follow-on lives in
`ChartWriter.cs` and does **not** conflict with this spec.
**Status:** ✅ **Done** — PR #401 (task 1), PR #402 (tasks 2 and 3). See Results.
**Was a hard prerequisite for spec 15** (shapes & text boxes), which is now unblocked.

## Summary

Spec 15 needs machinery that already exists in the codebase in chart- or picture-specific form:
anchor construction, schema-correct DrawingML property writing, and a way to prove an edit changed
exactly what it meant to. Extracting it is refactoring of *shipped* behaviour — spec 10's chart
formatting and the picture save path — and deserves its own gate rather than landing interleaved with
new-feature PRs. This spec does only that: one new test instrument plus two behaviour-preserving
refactors, each a self-contained PR, each gated by a green suite and golden fixtures.

Who this serves, honestly: **spec 15 is the primary consumer.** The secondary value is real but
smaller — the change-set harness permanently strengthens the existing chart-patch tests, and any
future chart-*formatting* work inherits one correct DrawingML ordering implementation instead of
risking a second. (An earlier draft claimed spec 10's open 3D-chart follow-on as a consumer; that gap
is chart *group element* emission in `ChartWriter.cs` — see `AppendBar3DChart`,
`ChartWriter.cs:740` — and is unrelated to this layer.)

## Scope boundary: extraction only

This spec adds **no new DrawingML capability**. Operations spec 15 needs that charts never perform —
emitting an explicit `a:noFill` (chart code removes fills and stops; it never writes one), writing
`a:prstDash` (chart series don't model dash at all) — are **not** added here, because the whole
contract of this spec is "behaviour-preserving, gated by the existing suite", and the existing suite
cannot gate operations it never exercises. Spec 15's task 3 extends the layer with those operations
under its own change-set tests. If a PR in this spec adds a code path the chart and picture suites
don't reach, it is out of scope by definition.

## Why these extractions are justified (and what their risk is)

`xdr:sp/spPr` and a chart series' `c:spPr` are the **same schema type** — `a:CT_ShapeProperties`. The
rules for writing into it are subtle and already implemented once, correctly, in `ChartFormatting`:

- Fills are a *choice* group — setting a colour over an `a:gradFill` means removing the whole group
  first (`SetFill`, `ChartFormatting.cs:1263`, via `IsFillElement`), not appending beside it.
- The type is a *sequence* — `a:xfrm`, geometry, fill, `a:ln`, `a:effectLst`, 3-D, `a:extLst` — and
  the SDK does not order children for you (`InsertAfterLastOf`, line 1334, plus the
  `IsAfterFill`/`IsAfterOutline`/`IsEffectOrLater` predicates).
- Existing elements are mutated in place, never rebuilt — `SetOutline` (line 1282) edits the `a:ln`
  that is there, so unmodeled children (arrowheads, miter) survive.

A second implementation of these rules for shapes is how the two copies drift. The decision (taken
with the project owner, 2026-08-01) is **full extraction**: the setters move, not just the predicates.

The honest cost: this is *not* a mechanical move. The current signatures are chart-coupled —
`SetOutline(C.ChartShapeProperties, XLChartSeries, XLChartSeriesFormat)` takes the series and its
flags enum, and `EnsureShapeProperties` positions `spPr` after `C.SeriesText`/`C.Order`/`C.Index`,
which is knowledge about the *chart parent's* sequence, meaningless for other hosts. The extraction
therefore redesigns working code, under these rules:

1. **The shared layer takes values, never masks.** Its signatures are of the form "set fill to X",
   "clear fill", "ensure outline exists, correctly placed" — plain nullable values and explicit
   operations. `XLChartSeriesFormat` stays in `ChartFormatting`; spec 15's `AssignedProperties` flags
   stay in its patcher. Callers decide *whether* to write; the layer decides *how*. If a flags enum
   appears in a shared-layer signature, the abstraction has failed.
2. **Parent positioning stays with the parent.** `EnsureShapeProperties`' knowledge of where `spPr`
   sits inside a chart series element remains chart-side; `xdr:sp` never needs an equivalent because
   `spPr` is a *required* child of `CT_Shape`. The shared layer operates *within* an `spPr` it is
   handed.
3. **Colour building moves with the setters.** `BuildColor`/`MapSchemeColor`
   (`ChartFormatting.cs:742-748`) are part of "write a fill" and are outside the setter line range —
   they move too. They already handle theme colours (`A.SchemeColor` for `XLColorType.Theme`), so the
   layer's consumers get theme-colour writing for free.
4. **Behaviour-preserving, proven twice.** The chart formatting and picture tests pass unmodified,
   and output is byte-identical against golden fixtures captured from the pre-refactor code
   (task 1's corpus).

What does **not** move: the in-place paragraph-editing rules inside `SetRichText`
(`ChartFormatting.cs:257-287` — `rPr`-preserving run replacement, `endParaRPr`-stays-last,
`IsRunLevel`). They are the right bones for spec 15's text patching, but their only real consumer
today is the chart title path, and shaping a shared abstraction against spec 15's still-on-paper text
model is premature (decision 2026-08-01). **Spec 15's task 2 performs that extraction** when the
second consumer exists, relocating the rules into `DrawingML/` with the chart tests still gating.

## Current state

Verified against the tree at `dfd5f49b` (2026-08-01).

- **Anchor construction**: `PictureWriter.AddPictureAnchor` (`PictureWriter.cs:197-391`) builds all
  three anchor forms — `AbsoluteAnchor` (FreeFloating), `TwoCellAnchor` (MoveAndSize), `OneCellAnchor`
  (Move) — as three near-identical inline blocks differing only in the `xdr:pic` payload and marker
  sources. Marker fallbacks default to A1 (`PictureWriter.cs:284-301`).
- **Property writing**: `ChartFormatting.cs:1251-1348` (setters, predicates, ordered insertion) plus
  `BuildColor`/`MapSchemeColor` at 742-748, consumed by the series patching path (`ChartPatcher`).
- **Test instrument**: chart-patch tests assert selected properties of the output XML. Nothing asserts
  "this edit changed *only* these nodes", which is the property patch-in-place actually promises.

## Work plan

The harness lands **first** — it is the gate the two refactors are measured with.

| # | Task | Size | Gate |
|---|------|------|------|
| 1 | **XML change-set harness + golden corpus.** (a) Canonicalize a part before and after an operation (namespace-aware; attribute order and insignificant whitespace normalized) and assert the change set equals exactly an expected node set — nothing else added, dropped, or reordered. (b) Golden fixtures: capture the current code's save output for every chart fixture exercised by the formatting tests plus a representative multi-picture save, committed with a green test asserting current output matches golden — this is what lets the later refactor PRs prove byte-identity, since a test cannot run pre-refactor code after the fact. (c) Retrofit two existing chart-patch tests (one fill edit, one title edit) to change-set assertions. | M | Retrofitted tests pass and demonstrably fail when a deliberate extra mutation is introduced; golden test green on unmodified code. |
| 2 | **`DrawingAnchorFactory`** — extract anchor construction from `PictureWriter.AddPictureAnchor`: (placement, pixel geometry / markers, child element) → `AbsoluteAnchor`/`OneCellAnchor`/`TwoCellAnchor`, EMU conversion included. The A1 marker fallbacks move with it and are documented as part of the factory's contract, since future callers (spec 15) inherit them silently. `PictureWriter` becomes a caller. | S | Full picture test suite green; golden byte-identity holds. |
| 3 | **`IO/DrawingML/ShapePropertiesWriter`** — full extraction of `SetFill`/`SetOutline`/`BuildColor`/`MapSchemeColor`/ordering predicates/`InsertAfterLastOf`/`AppendBeforeExtensionList` under the redesign rules above, over `OpenXmlCompositeElement`. `ChartFormatting` keeps thin mask-driven adapters and its parent-positioning logic. No new operations (see scope boundary). | M | Chart formatting tests green, unmodified; golden byte-identity holds. |

Tasks 2 and 3 are independent of each other once 1 lands.

## Acceptance criteria

1. No public API change: `PublicAPI.Unshipped.txt` untouched (everything added is `internal`).
2. Full suite green on net8.0/net9.0/net10.0 after each PR, with no test modified except the two
   deliberately retrofitted in task 1.
3. Golden byte-identity: every fixture in task 1's corpus saves byte-identically after tasks 2 and 3.
   Any diff is a finding to investigate, never noise to re-baseline without a written explanation.
4. No flags enum (`XLChartSeriesFormat` or any successor) appears in a `DrawingML/` signature.
5. The harness absorbs the SDK's serialization quirks: a part whose DOM **was materialised and
   re-serialized** with no model edits canonicalizes to an **empty change set** (mere DOM loading
   re-serializes — see the comment at `PictureWriter.cs:138-140`; this criterion is what makes the
   harness usable for spec 15's fidelity tiers).
6. Scope boundary held: the extracted layer's observable capability set is unchanged — no operation
   exists in `DrawingML/` that the chart and picture suites do not exercise.

## Risks

- **This refactors shipped spec-10 behaviour for the benefit of code that does not exist yet.** The
  mitigation is structural: each task is a separate behaviour-preserving PR gated by suite + goldens,
  landed before spec 15 writes a line against it. If spec 15 is later descoped, the harness and the
  single-implementation ordering layer still pay for themselves in future chart-formatting work.
- **The value-only redesign may surface latent chart bugs** — e.g. an edit order the mask-driven code
  happened to make unobservable. The golden corpus (criterion 3) is the tripwire.
- **Canonicalization is deceptively hard** (namespace prefixes, `xml:space`, attribute ordering, the
  SDK's own normalization). Task 1 is sized M for this reason and its "deliberately broken mutation"
  check exists to prove the harness detects what it claims to.

## Results

**Done.** Task 1 shipped as PR #401; tasks 2 and 3 together as PR #402. `main` at `62b41f79`.

| Task | Delivered | PR |
|---|---|---|
| 1 | `XmlChangeSet`, `GoldenCorpus`, 11 golden fixtures, 13 self-tests, 2 retrofits | #401 |
| 2 | `DrawingAnchorFactory`, `DrawingAnchorGeometry`, `DrawingUnits`; `PictureWriter` net −129 lines | #402 |
| 3 | `ShapePropertiesWriter`; `ChartSeriesFormatXml` net −69 lines | #402 |

All six acceptance criteria met. `PublicAPI.Unshipped.txt` untouched — every added type is
`internal`. The nine `patch-*` and two `pictures-*` goldens are byte-identical after both refactors.
No test was modified except the two retrofitted in task 1, and the solution suite stayed at 29,146
across net8.0 and net10.0 — the same count before and after, so nothing was added, removed or
quietly skipped.

### Four things the spec got wrong, and what they cost

**1. The extraction targets had already moved.** The spec's line references are from `dfd5f49b`
(2026-08-01). Spec 22 deleted `ChartFormatting.cs` between the spec being written and being run, so
every `ChartFormatting.cs:NNNN` citation is dead. The setters live in
`Charts/ChartSeriesFormatXml.cs`; ordering is in `Charts/ChartElementOrder.cs`.

**2. A golden corpus already existed.** Spec 22 shipped `ChartGoldenCorpus` plus eight fixtures,
pinning the path that *builds a new chart*. This re-scoped task 1(b): the real gaps were the
**patch** path and the **drawing part**, which nothing pinned — precisely what tasks 3 and 2 rewrite.
Capturing what the formatting tests already exercised would have been near-worthless.

**3. Tasks 2 and 3 are not independent.** True of the design — they touch `PictureWriter` and
`ChartSeriesFormatXml` respectively — but false of the implementation. Both needed the same units
helper, task 2 introduced `DrawingUnits`, and task 3 extends it. Task 3 does not compile without
task 2. They were raised as one PR for this reason.

**4. The three anchor blocks were more alike than described.** The spec says they differ in "the
`xdr:pic` payload and marker sources". Diffed line by line, the payload regions are **byte-for-byte
identical** — only the anchor element varies. That is what allowed the picture to be built once and
passed in as `content`, and it is what makes the factory reusable for `xdr:sp` without change.

### Two items on the extraction list deliberately did not move

`AppendBeforeExtensionList` names `C.ExtensionList`, a **chart-namespace** type. Moving it into
`DrawingML/` would have put a chart type into a shared signature — the same category of error as a
flags enum, and a contradiction of criterion 4's intent.

`InsertAfterLastOf` could have moved, but `ChartElementOrder.InsertOrdered` — introduced by spec 22
*after* this spec was written — is a strictly better generalisation of both helpers, with per-parent
order tables the sibling chart modules already use. Moving the older helper would have left the tree
with **three** ordering implementations instead of two. The correct follow-up is chart-side: fold
`ChartSeriesFormatXml`'s two older helpers onto `InsertOrdered`. It is not free — it needs a
`MarkerChildOrder` table and care at eight call sites, because the two mechanisms differ in their
fallback (`InsertAfterLastOf` inserts at index 0, `InsertOrdered` appends).

### The harness was proven, not asserted

Task 1's brief required the change-set assertions to fail on a deliberate extra mutation. A real one
was introduced into `ChartPatcher.PatchChart` — flipping `c:plotVisOnly` on every patched chart. Both
retrofitted tests failed and **named the line**:

```
+ 3: modified /c:chartSpace[1]/c:chart[1]/c:plotVisOnly[1] @val: '1' -> '0'
```

Nine of the eleven goldens failed with it. The mutation was reverted; a permanent self-test covers
the same property so the guarantee does not depend on repeating the demonstration.

Two design decisions came out of building it:

- **The SDK does not re-serialize a part until something touches a child.** The criterion-5 test
  initially passed vacuously, because `new C.ChartSpace(xml).OuterXml` hands the raw input straight
  back. It now touches `ChildElements` and asserts the bytes really changed *before* asserting the
  change set is empty. That asymmetry — an unread part keeps its bytes, a read one does not — is
  itself why byte comparison cannot do this job, and it matters for spec 15's fidelity tiers.
- **Assertions compare text blocks, not collections.** An earlier version used TUnit's
  `IsEquivalentTo` and failed with `collection has 3 items but expected 2`, never naming the extra
  change — which defeats the instrument.

### Positional identity: a documented limitation, not a defect

Review of PR #401 asked for same-name sibling sequences to be aligned before ordinals are assigned,
keyed on something stable such as `c:ser/c:idx`. The behaviour is real: inserting or removing a
`c:ser` ahead of its siblings renumbers the rest, so each is reported as taking on its predecessor's
content plus one addition at the end of the run.

**The remedy was declined, and the reason is the harness's whole purpose.** Positional identity can
make a change set *louder* than the edit was; it cannot make one **empty**, because an empty change
set means every path in one document is present in the other with the same attributes, text and
child order. This harness gates refactors that must change nothing — over-reporting costs a reader
some time, under-reporting would be a false pass, and this cannot produce one.

`c:ser/c:idx` as a key is also circular: it keys identity on a value an edit is allowed to change, so
a `patch-*` fixture that legitimately renumbered a series would read as a series being replaced. And
a general-purpose diff that must also handle drawings and, later, shapes should not carry one
schema's knowledge. Proper alignment means a real sequence diff — a feature, not a fix.

Four tests now pin the behaviour, including the boundary that matters: a **trailing** sibling
insertion renumbers nothing, which is the case that arises when a drawing is appended to a part.

### Loose ends

- `SetOutlineColor(outline, null)` — clearing a series' line colour — has no test. Nothing in the
  suite sets `LineColor = null`. It was equally untested before the extraction, so no path was added,
  but it is now a named operation of a shared layer.
- The `default:` arm of `DrawingAnchorFactory.Create` is unreachable without casting an out-of-range
  integer to a three-member enum. A guard, not an operation.
- `XLibur/Excel/Drawings/XLPictures.cs:294` still holds a third copy of the pixels→EMU conversion.
- `XLibur.Excel.Coordinates.Emu` is **not** interchangeable with `DrawingUnits`: it stores `int` and
  rounds away from zero where the drawing conversion rounds to even, so substituting one shifts
  extents at `.5` boundaries. Both types cross-reference each other.
- `TestInfrastructure.cs` still claims the test project has nullable off. It does not —
  `Directory.Build.props:34` enables it unconditionally. The same false comment was removed from the
  two files this spec added.
- ~40 lines of resource plumbing are duplicated between `ChartGoldenCorpus` and `GoldenCorpus`.

### What this unblocks

Spec 15 named this spec a hard prerequisite and is now clear to start. What it inherits:

- `XmlChangeSet` for its fidelity tiers, including the empty-change-set-on-reserialization property.
- `DrawingAnchorFactory`, which takes its content element as a parameter — an `xdr:sp` needs no
  change to it. **Read its `<remarks>` first:** the A1 marker fallbacks are contract, and a drawing
  created `MoveAndSize` with neither marker set takes both at once.
- `ShapePropertiesWriter` over `OpenXmlCompositeElement`, so `xdr:sp/spPr` uses the same
  implementation as a chart series' `c:spPr`.

Per the scope boundary, the operations spec 15 needs that charts never perform — an explicit
`a:noFill`, `a:prstDash` — were **not** added here. Spec 15's task 3 adds them under its own
change-set tests. Spec 15's task 2 still owns the `SetRichText` extraction, now at
`Charts/ChartTitleXml.cs:144`.
