# Spec 16 — Shared DrawingML Infrastructure

**Area:** Architecture + Refactor
**Effort:** S–M (~1 week)
**Dependencies:** None. Conflicts with any in-flight work in `XLibur/Excel/IO/ChartFormatting.cs`
(spec 10's territory) and `XLibur/Excel/IO/PictureWriter.cs`. Spec 10's open 3D follow-on lives in
`ChartWriter.cs` and does **not** conflict with this spec.
**Status:** Proposed. **Hard prerequisite for spec 15** (shapes & text boxes).

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
