# XLibur Improvement Roadmap

Thirty-five prioritized, self-contained specs covering features, compatibility, architecture, and performance (memory + read/write times). Each spec is written to be handed to an independent agent/model: it states the problem with measured numbers, points at the exact files, prescribes a design, breaks the work into PR-sized tasks, and defines measurable acceptance criteria.

**Start a new performance effort at [spec 19](19-benchmark-hotspot-survey.md)**, not at this table. It re-ran the whole suite on 2026-08-07 and ranks what is actually slow now, which is not what specs 02–18 would predict — the biggest single number in the suite turns out to be the `CellsUsed()` enumeration, not parsing, packaging or styling. It also carries the current baselines for every benchmark and the run recipe.

Specs 01–10 are the original top-ten set; spec 11 is a follow-on that came out of implementing spec 03 (see below).

**This folder is the source of truth for specs.** The repo's `docs/specs` is a copy on its way out, kept in step by copying this folder over it as part of a spec's PR. **That sync copies the specs and the tasklists only — never `briefs/`.** The briefs are conductor dispatch records: they name local worktree paths, address a particular agent, and describe how work was handed out rather than documenting the library. They must never enter the repo. This is a standing exclusion, so a sync that copies the folder wholesale has to drop `briefs/` again.

Grounding: specs 01–10 were derived from a July 2026 survey of the codebase (architecture, feature inventory vs Excel, benchmark artifacts under `BenchmarkDotNet.Artifacts/results/`). Headline baselines: save 50K rows ≈ 1.0–1.1 s / **543 MB allocated**; load+read 250K×15 ≈ 5.6 s / 1.68 GB after PR #171; XLibur is already ~3× faster and ~6× leaner than upstream ClosedXML on save.

## The list

| # | Spec | Area | Effort | Status | Parallelizable? |
|---|------|------|--------|--------|-----------------|
| 01 | [Streaming write API](01-streaming-write-api.md) | Feature · Arch · Memory | L | ✅ **Done** (see Results) | Phase 1 refactor first, then independent |
| 02 | [Load-path allocation elimination](02-load-path-allocations.md) | Perf (read) | M | ✅ **Done** (#175) | 3 independent sub-tasks |
| 03 | [Save-path allocation reduction](03-save-path-allocations.md) | Perf (write) | M | In progress (see Results) | 7 small independent PRs |
| 04 | [Demand-driven formula evaluation](04-demand-driven-formula-eval.md) | Perf · Arch | L | Proposed | Single owner (correctness-critical) |
| 05 | [Structural-edit & bulk-style scalability](05-structural-edit-scalability.md) | Arch · Perf | L | ✅ **Done** (see Results; C1 declined) | 3 independent workstreams |
| 06 | [Workbook encryption (password files)](06-workbook-encryption.md) | Feature · Compat | L | ✅ **Done** (#245) | Container/crypto layers in parallel |
| 07 | [Formula function coverage (257→~420)](07-formula-function-coverage.md) | Feature | L | ✅ **Waves A–F done** (#252–#257) | **6 fully independent waves** |
| 08 | [LET / LAMBDA](08-let-lambda.md) | Feature | L | Proposed | Single owner (engine core) |
| 09 | [Threaded comments + round-trip fidelity](09-threaded-comments-roundtrip.md) | Feature · Compat | M | ✅ **Done** (#258) | Comments vs fidelity-audit split |
| 10 | [Chart formatting depth](10-chart-formatting-depth.md) | Feature | L | ✅ **Done** (PRs 1–4) | 4 PRs, 2–3 independent |
| 11 | [Create-path allocation reduction](11-create-path-allocations.md) | Perf (write) | M | ✅ **Tasks 1–4 done** | Task 4 lands in 11; 05 rebases |
| 12 | [Report templating (`XLibur.Report`)](12-report-templating.md) | Feature · Arch | L | ✅ **Done** (see Results; gauge corpus not ported) | 11 tasks; 4/5/6/10 parallel after 3 |
| 13 | [Public core surface for `XLibur.Report`](13-report-core-public-api.md) | Arch · API · Packaging | M | ✅ **Done** (#354) | Tasks 1 and 2 independent |
| 14 | [`Clear`/`CopyTo` scalability](14-clear-copyto-scalability.md) | Perf (edit) · Correctness | S | Proposed ([#271](https://github.com/XLibur/XLibur/issues/271)) | Task 1 first; 2/3/4 independent |
| 15 | [Shapes & text boxes](15-shapes-and-text-boxes.md) | Feature · Compat | L | Proposed (**16 has landed — unblocked**) | 1–3 one stream; then 4/5 parallel |
| 16 | [Shared DrawingML infrastructure](16-drawingml-infrastructure.md) | Arch · Refactor | S–M | ✅ **Done** (#401, #402; see Results) | 3 PRs; harness first, then 2/3 independent |
| 17 | [Picture styling & fidelity](17-picture-styling.md) | Feature · Compat · **Defect** | M–L | Proposed (**16 has landed — unblocked**) | Task 1 (fidelity fix) first and standalone; 3/4/5 parallel after 2 |
| 18 | [Template round-trip overhead](18-template-round-trip-overhead.md) | Perf (read + write) · **Defect** | M | Tasks 0–4 done (see Results) | Task 5 is the remaining cost; independent |
| 19 | [Benchmark hotspot survey (Aug 2026)](19-benchmark-hotspot-survey.md) | Perf (read + write) · Survey | L | Proposed | 5 areas; 1/2/3 fully independent |
| 20 | [Style key struct sizes](20-style-key-struct-size.md) | Perf (write · bulk styling) · Memory | M | Proposed | Task 0 first; 1→2 ordered; 3/4 independent |
| 21 | [Hot-path struct candidates](21-hot-path-struct-candidates.md) | Perf (read · enumeration) | M | ✅ **Done** (task 3 shipped; 1–2 declined on measurement) | Task 0 first; 1→2 ordered; 3 independent |
| 22 | [Chart IO: one module per chart concept](22-chart-concept-modules.md) | Arch · Refactor | M | ✅ **Done** (tasks 0–6; run before 16 — see Results) | Single owner; tasks sequential |
| 23 | [One implementation per style interface](23-single-style-facade.md) | Arch · Refactor · **Defect** | M | ✅ **Done** (#397; see Results) | Single owner; tasks sequential |
| 24 | [Worksheet element load gets one interface](24-worksheet-element-dispatch.md) | Arch · Refactor | S–M | ✅ **Done** (#395; see Results) | Single owner; tasks sequential |
| 25 | [Narrow the formula shifter's fallback](25-formula-shifter-seam.md) | Arch · **Correctness (masking)** | S | ✅ **Done** (#398) | Single owner; tasks sequential |
| 26 | [Give the grid one axis](26-grid-axis.md) | Arch · Refactor · **3 defects** | L | ✅ **Merged** (#409; see Results) | One `IGridAxis`, two adapters; 3 defects fixed; allocations down 12–50% |
| 27 | [One font conformance module](27-font-conformance-suite.md) | Test · Arch (seam) | S–M | Proposed | Single owner; **gates 34** |
| 28 | [One OOXML style decoder](28-single-style-decoder.md) | Arch · Refactor · **Defect (data loss)** | M | ✅ **Merged** (#411; see Results) | Single owner; 3 defects fixed; one premise disproved; load allocations flat or down |
| 29 | [One resolver per emitted element](29-write-path-resolvers.md) | Arch · **Correctness (divergence)** | M | 🟡 **PR open** ([#413](https://github.com/XLibur/XLibur/pull/413), 2026-08-27; see Results) | 10 commits, 28,358 green both TFMs; pane divergence fixed; **D18 found and open**; merge before 33; 31 unblocks *on merge*, not now |
| 30 | [Array application gets an interface](30-array-application-seam.md) | Arch · **Defect (241 functions)** | S–M | Proposed | Single owner; **before 32** |
| 31 | [Worksheet element writers get one interface](31-worksheet-element-writers.md) | Arch · Refactor | M–L | Proposed (**needs 29 merged**) | Single owner; tasks sequential. 29 is complete on a branch but **not merged** — starting 31 now means rebasing a structural sweep onto an unmerged branch |
| 32 | [Collapse the 61-overload registration](32-function-argument-spec.md) | Arch · Refactor | L | Proposed (**needs 30**) | Single owner; task 2 is a go/no-go gate |
| 33 | [Every sheet feature reacts through one seam](33-sheet-listener-seam.md) | Arch · **Defect (4 unshifted)** | M–L | 🟡 **PR open** ([#414](https://github.com/XLibur/XLibur/pull/414), 2026-08-27; see Results) | 11 commits, 28,444 green both TFMs. Shifter 222→65 lines, names no feature; 11 adapters (was 2); the 4 dead features move; **D15–D17 recorded, D15 and D17 live**; criterion 2 reported unreachable. **Merge #413 first** (shared `docs/specs` sync, textual conflicts, keep both). Does **not** unblock 34, which waits on 27 |
| 34 | [Split the font port: mechanism vs policy](34-font-port-split.md) | Arch · Refactor | M | Proposed (**needs 27**) | Single owner; tasks sequential |
| 35 | [Pivot table timelines](35-pivot-timelines.md) | Feature · Compat | M | ✅ **Done** (#406; see Results) | Task 1 (extraction) standalone; 2→3→4 ordered |

**Specs 26–34 came out of a second architecture review on 2026-08-24.** Their progress board,
dependency graph, conflict map and wave plan live in
[TASKLIST-architecture-deepening-2.md](TASKLIST-architecture-deepening-2.md).

All nine are the same shape — **one fact with two or more implementations, kept in agreement by
hand** — which is the shape round 1 found twice (23's style facades, 25's shifter) and round 2 found
nine times. The difference is that **five of these agreements have already failed in shipped code,
and nothing catches any of them**:

| Spec | Drift | Effect |
|---|---|---|
| 26 | `XLRow.cs:424-425` calls `IncrementColumnOutline`, copied from `XLColumn.cs:342-343` | `IncrementRowOutline` has zero callers; `@outlineLevelRow` is never emitted and `@outlineLevelCol` is inflated by row groups |
| 26 | `XLColumn.CellCount()` is character-identical to `XLRow.CellCount()` | Returns 1 instead of 1,048,576 |
| 28 | `LoadFont:202` searches the `<x:rPr>` element spellings while all three callers pass a `Font` | ✅ Fixed. A conditional-format font silently lost its **name, family numbering and charset** on load — confirmed by test: `FontName` came back `"Calibri"`, not `"Arial"`. Two further defects fixed alongside: an indent in a pivot dxf threw on load, and a duplicated `numFmtId` made a workbook unopenable |
| 29 | ~~`SheetViewWriter.cs:124` writes `frozenSplit`; `XLStreamingWorksheet.cs:502` writes `frozen`~~ | **Fixed 2026-08-27.** Both paths now resolve through `XLPaneSettings` and write `frozen`; the DOM path also stopped writing `xSplit="0"` for an unsplit axis |
| 30 | `FunctionDefinition.cs:106-118` builds `itemArg`, then calls `_function(ctx, args)` | `POWER({2,3,4},{1,2,3})` → `2,2,2` not `2,9,64`; **261 scalar functions** affected under array semantics, worksheet references included |

The missing test in each case is not an oversight — it is a consequence of the shape. Where a module
has one interface, the interface is the test surface; where the same fact has two implementations,
nothing sits at the seam to assert they agree. **Spec 30's defect is the clearest illustration.** Its
origin (`819528c9`, 2023) was correct; upstream `fc08037c` then deleted a wrapper and inlined it at
two call sites, applying the same replacement text to both — right at one, wrong at the other. The
fork's Sonar pass (#12) later extracted that loop into a helper *for testability*, gained no test
surface, and froze the bug in a method small enough to read as obviously fine. Extracting code
without giving it a test surface preserves its bugs exactly.

Spec 27 is the cheapest in the round and the one to start first: three adapters satisfy
`IXLFontEngine`, the two adapter test suites are 421 identical lines out of 434, **no file in the
repository references two font engines**, and the core autofit suite runs against V1 while the
shipped default is SkiaSharp. It touches no production code and it gates 34.

Two specs carry a **measurement gate empowered to stop the work**, following spec 21's precedent:
32's task 2 (an `ArgSpec[]` loop moves argument-shape resolution from compile time to the hot path)
and 34's task 6 (text measurement is on the autofit path). In both, a recorded measurement that
halts the spec is a real result.

The nine form four dependency chains plus one free spec, so they run as **five parallel streams, then
four**: 26→33, 27→34, 29→31, 30→32, with 28 independent.

**Specs 22–25 came out of an architecture review on 2026-08-23** that asked where modules are
*shallow* — interface nearly as wide as the implementation — rather than where they are slow. Their
progress board, conflict map and parallel-execution plan live in
[TASKLIST-architecture-deepening.md](TASKLIST-architecture-deepening.md). The four are **file-disjoint
from each other**, so 23, 24 and 25 can run as three concurrent streams today; 22 is blocked on 16.

Two of the four are defect reports as well as refactors. **23** found that every style interface has
two implementations — the ordinary facade and an `XLDeferred*` twin reached through
`IXLStyle.Batch` — and that they disagree on `InsideBorder`/`InsideBorderColor`: for a cell the
direct path is correctly a no-op (a 1×1 range has no interior edges) while the batch path sets all
four. **25** found that the shifter's regex fallback is selected by `catch (Exception)`, so a bug in
XLibur's own `ShiftPlan` is answered with a plausible result from the other implementation instead
of surfacing. Neither defect is new; both exist because two implementations must agree by hand, which
is what these specs remove.

Spec 21 came out of a review asking which hot-path classes could become structs, and its most useful
output is a negative: **almost everything that should already be a struct already is.** `Point`,
`Area`, `XLAddress`, `XLRangeAddress`, `ScalarValue`, `AnyValue`, `XLCellValue`, `XLUsedCell`, every
style key, and the load/save cell types are structs today — specs 02, 03, 05, 11, 18 and 19 got there
first. Ten types were reviewed and seven were rejected, each with a specific disqualifier recorded so
it is not re-proposed: `Formula` lives in a `ConditionalWeakTable` (`TValue : class`);
`TransitionEntry` is a class *on purpose*, so a cache slot fills with one atomic reference write and
cannot tear; `XLCell`, `XLCellFormula` and the `XL*Value` style objects are all tracked by reference
identity. What is left is the **slice enumerators** — `sealed class` with all-value-type fields,
reached through `IEnumerator<Point>`, so `SlicesEnumerator`'s k-way merge pays up to eight
non-inlinable interface calls per cell. 21 is a dispatch-and-inlining claim with no allocation
signature to fall back on, and **its task 4 is empowered to revert task 2** if four inline struct
enumerators cannot be kept from being copied.

**21 is now done, and its headline is a disproved premise.** Converting `Slice<TElement>.Enumerator`
to a struct is *free* (5,069 µs against a 5,053 µs baseline) — but **embedding** it by value in the
enclosing enumerator costs **+60%** on the walk, measured across five variants, with the wrapper
layer and the JIT's inlining decisions both ruled out as causes. The interface dispatch the spec was
written to remove had already been devirtualised by dynamic PGO, so it was never the cost. Task 1 was
implemented and reverted; task 2 was declined without being written, since it would embed four such
enumerators. **Task 3 shipped**: `XLRangeParameters` becomes a `readonly struct` and
`XLRangeBase.GetRange` stops materialising it and a style façade ahead of its own bounds check —
−13.6% time and 78.19 KB → 3 B per 1,000 sub-range calls, all of it on the repository-cache-hit path.
21's task 4 decision rule was itself mis-stated (it would have kept a 60% regression to remove an
88-byte per-enumeration constant) and is corrected in its Results — the fourth criterion in this spec
family to price work its task could not reach.

Spec 15 closes the last hole in the drawing surface: XLibur can add pictures and charts but no shape of
any kind, so a floating text box, callout or arrow cannot be created at all. Shapes that already exist in
a file do survive a round trip — save reopens the original package and rewrites only modelled parts, the
same mechanism `docs/round-trip-fidelity.md` documents — but they are invisible to the model, so their
text can be neither read nor changed. 15 adds `ws.Shapes` over DrawingML `xdr:sp` with a paragraph-aware
text model, and follows spec 10's chart precedent: new shapes are generated, loaded shapes are **patched
in place** so unmodelled XML (gradients, effects, theme styles) keeps surviving untouched.

Spec 16 was split out of 15 by review (2026-08-01): the machinery 15 needs already exists in chart- and
picture-specific form, and extracting it is refactoring of *shipped* behaviour that deserves its own
gate rather than landing interleaved with feature PRs. 16 delivers three internal pieces, harness
first: an XML change-set test harness with golden fixtures (proven by retrofitting two chart tests);
the anchor factory out of `PictureWriter.AddPictureAnchor`; and the DrawingML
fill/line/colour/element-ordering layer out of `ChartFormatting` — `xdr:sp/spPr` and a chart's `spPr`
are the same schema type, and one implementation of it beats two — with value-only signatures so
neither side's assigned-flags enum leaks in. 16 is strictly extraction: operations charts never
perform (`a:noFill` emission, `a:prstDash`) and the `SetRichText` paragraph-editing primitives are
added/extracted by 15 itself, where their real consumer shapes the code and its tests gate it.
**15 hard-depends on 16.** (Spec 10's open 3D follow-on is unrelated — that gap is chart group
emission in `ChartWriter.cs`, not this layer.)

Spec 17 started as "add the missing picture styling" and turned into a defect report: the picture save
path **destroys existing styling**. `PictureWriter.AddPictureAnchor` replaces every picture's anchor on
every save with a hardcoded `spPr` and `blipFill`, so rotation, borders, shadows, recolors and crops
authored in Excel are silently lost the first time XLibur saves the file — the round-trip-fidelity
guarantee never covered pictures, because it only protects what XLibur does not rewrite. 17's task 1 is
the standalone fix (patch-in-place for loaded pictures, clean sheets untouched, image binaries no longer
re-fed every save); the styling model — border, fill, rotation/flip, transparency, single outer shadow,
recolor presets, crop, brightness/contrast — builds on top of it through spec 16's shared layer plus a
new `BlipEffectsWriter`. **17 hard-depends on 16** and conflicts with 15 in `PictureWriter.cs`
(sequential, either order).

Spec 13 came out of preparing `XLibur.Report` to version independently of core. Report reads core
internals through an `InternalsVisibleTo` grant, which is safe only while the two ship as one
version off one tag. Building it against the released core package fails with 9 errors across two
files. Internals carry no compatibility contract, so a version floor over them would be a promise
core never made — a consumer referencing both packages gets core unified upward by NuGet and
Report breaks at runtime. 13 replaces the grant with two narrow public additions, a function
library callable without a grid and a pivot cache source that can be re-pointed, both expressed in
already-public types. **It is a hard prerequisite for Report's own version stream**, and ships in
core 0.107.0. Its task 3 — moving Report onto that surface — is the one piece of spec 12's package
still to write, and it supersedes the two `internal` core members spec 12 added
(`FunctionRegistry.Names`, `XLCalcEngine.Functions`).

Spec 11 was added after spec 03 landed: 03 halved the *save* phase and showed the rest of it is
`System.IO.Packaging`, leaving the *create* phase as 72% of the write benchmark. It is a follow-on,
not part of the original ten.

Spec 14 came out of implementing spec 12, the same way spec 11 came out of spec 03: spec 12's benchmark
criterion caught report generation scaling super-linearly, and the cause turned out to be a core-library
defect — `XLRangeBase.Clear` creates and deletes a data validation on every call even when the sheet has
none, which makes any range copy in a loop quadratic. Tracked as
[#271](https://github.com/XLibur/XLibur/issues/271); the two-line fix is prototyped and measured there
(`CopyTo` 420 µs → 13 µs at 30,000 rows).

Spec 12 (July 2026) is a feature spec outside the original survey: a report-templating package
(`XLibur.Report`) porting the ClosedXML.Report architecture with a Scriban expression engine,
an Excel-function bridge into `{{ }}` expressions, first-class chart/pivot/image handling
across range expansion, and an opt-in `XLibur.Report.DynamicLinq` package that runs
ClosedXML.Report's C#-expression template syntax unmodified. It is **done** bar the gaps its Results
section lists. Its core-side footprint grew beyond the `InternalsVisibleTo` grant it planned for — two
`internal` members for the function bridge, and a chart-reference fix the save path needed — all recorded
as deviations in spec 12 and all superseded or kept by spec 13.

Spec 02 delivered **−16.5% load time and −61.5% allocations** (4.750 s / 1020.92 MB → 3.968 s /
392.88 MB on the 250K×15 benchmark). It also produced two findings that change other specs: a
correction to spec 03's number-formatting task, and a reusable `XmlReader.ReadValueChunk` technique
for the IO layer. Both are recorded in spec 02's Results section.

Spec 10 landed in four PRs. Its PR 1 settled the question the spec flagged as its own hard part: the
chart writer never regenerates a chart it loaded, so unmodeled chart XML round-trips byte for byte,
and edits are patched into the existing part instead — PRs 2–4 extended that patcher rather than
adding a second write path. Along the way, turning the OpenXML validator on for the new chart tests
surfaced three long-standing schema violations in the chart writer, and the reader turned out to drop
one-cell/absolute anchored charts and every 3D or of-pie chart group. All fixed; see spec 10's four
Results sections. **Still open there:** the writer emits 3D pie/line/area and every surface type as
their 2D group elements, so XLibur's own 3D charts round-trip as 2D.

Spec 05 landed, and its main output is a correction to its own premise. It attributed the cost of
one-at-a-time row inserts to the range-shift pass materialising and sorting every live range; measured,
that is 8% of the workload its acceptance criterion names, and formula shifting is 68%. Two of its five
acceptance criteria were **disproved rather than met** — criterion 1 describes a workload dominated by a
fixed per-insert cost the spec never located, and criterion 3 asks for a complexity class that no
implementation can deliver. The prescribed spatial index over the range repository was declined on
evidence: a filter made unreachable ranges free, which showed the enumeration an index would remove was
never the expense. The real win was elsewhere — rewriting reference shifting onto `ClosedXML.Parser` cut
the combined workload 4,753 ms → 1,539 ms and fixed a reference-shifting bug (a deletion removing the
tail of a range dropped a surviving row: `A2:A8` with rows 5–9 deleted gave `A2:A3`, not `A2:A4`). See
spec 05's Results for the decomposition and for where the remaining 43% is.

Spec 01 landed and resolved the packaging hotspot spec 03 had deferred to it. The blocker was not
part *lifetime*, as 01 assumed, but part *buffering*: `System.IO.Packaging` opens a package
read/write, which is `ZipArchiveMode.Update`, and that holds every part's uncompressed bytes until
close. `XLStreamingWorkbook` therefore writes its OPC package straight over `ZipArchive` in Create
mode. Two capabilities fall out of owning the zip: output no longer has to be **seekable**, and
`CompressionLevel` became available — which also closed 01's Phase 3 for the *ordinary* save path,
since the SDK does expose `OpenXmlPackage.CompressionOption`. Async remains unimplemented, with the
reasons recorded in 01's Results.

## Why these ten (01–10)

**Performance (specs 02, 03, 04, 05).** The write cell-loop and the sheetData parse have both had a round of tuning; what remains, in measured order: per-cell string allocations on load (`<v>` + attributes + SST DOM), the ~543 MB formatted save (number formatting, inherited-style resolution, StyleKey hashing), the full-workbook-recalc cliff when reading one dirty formula cell (can build a 176 MB dependency tree to answer one read), and O(all-ranges·log) work per single row insert. Specs 02–05 attack each with concrete targets.

**Features (specs 01, 07, 08, 10).** The fork already leads upstream on charts, dynamic arrays, in-cell images, sparklines, WebP/SVG. The gaps that matter: no bounded-memory export path for huge files (01 — now done: `XLStreamingWorkbook` writes 1M×10 in 108 MB, or 14 MB with inline strings), ~250 missing formula functions with a clean registry to extend (07 — waves A–F now done; only the optional day-count-basis set A2 remains), no LET/LAMBDA (08 — still open, the one function family that needs engine work), and charts that couldn't be styled (10 — depth on the flagship differentiator, now done).

**Compatibility (specs 06, 09).** Both are now done. Password-encrypted files could not be opened or written at all — 06 added agile read/write and standard read (#245). Threaded comments were read lossily and downgraded on save — 09 gave them a real model and a write path (#258). 09's fidelity audit also **disproved its own premise**: chartsheets, form controls and slicers are *not* dropped on round-trip, because saving reopens the original package and rewrites only the parts XLibur models. See `docs/round-trip-fidelity.md`, which pins that behaviour with tests so a future rewrite cannot silently regress it.

**Architecture (specs 01, 04, 05).** The three structural debts the surveys flagged: no streaming seam in the IO layer, calc-engine all-or-nothing recalculation, and materialize-everything patterns for range shift and style propagation (plus two redundant spatial indexes — QuadTree and RBush).

## Suggested sequencing

```
Wave 1 (independent, start anytime):
  02 load allocations ✅ done · 03 save allocations (in progress) · 07 function waves A–F ✅ done
  09 threaded comments ✅ done
  11 Tasks 1–4 ✅ done (−28.8% on the write benchmark; bulk styling −86% per cell)
Wave 2 (after 03 lands, or coordinated):
  01 streaming write ✅ done (leaf serializers shared with 03's territory, not an enumeration seam)
  06 encryption ✅ done · 10 charts ✅ done (PRs 1–4: series formatting, data labels, legend/axes, reader gaps)
Wave 3 (single-owner, correctness-critical — don't parallelize internally):
  04 demand-driven eval · 05 structural edits ✅ done (C1 declined) · 08 LET/LAMBDA (08 after or alongside 04)

Remaining open work: 03 (finish), 04 demand-driven eval, 08 LET/LAMBDA, plus the optional 07
wave A2 (day-count-basis financial functions). Spec 05 leaves one follow-on it did not scope:
the ~665 ms of fixed per-insert cost paid on an empty sheet, which its Results section traces
to RelocateRange probing the range repository for XLRow instances that are never stored there.
```

**Read spec 02's Results section before starting 03** — it corrects 03's number-formatting task
and describes an allocation technique that applies to the rest of the IO layer.

Conflict map: 01↔03 (`SheetDataWriter`), 04↔08 (evaluation stack / `CalcContext`), 07 waves B↔C (`Statistical.cs`), 16↔any chart-*formatting* work (`ChartFormatting.cs` — 16 extracts its DrawingML property layer; spec 10's open 3D follow-on is in `ChartWriter.cs` and does not conflict), 15→16 and 17→16 (hard dependencies), 15↔17 (`PictureWriter.cs` save orchestration — sequential, either order), 15↔17 also share the shared-layer `noFill`/`prstDash` ops and `XLLineDashStyle` (first lander adds them). Everything else is disjoint. **Spec 05 must rebase onto spec 11**: 11's Task 4 rewrote bulk style propagation (`XLStylizedBase.ModifyStyle` / `SetStyle`), which is 05's territory.

Adding specs 22–25: **22→16 (hard dependency** — 16 extracts the DrawingML layer out of
`ChartFormatting.cs`, and 22 reorganises what is left; 16's change-set harness is also 22's gate),
**24↔18 task 5** (`XLWorkbook_Load.LoadWorksheetElements` — 24 first is recommended, since it leaves
18 one method to optimise rather than four across two modules), **23↔20** (`XL*Key.cs` field names,
read by the style facades' `with` expressions — sequential, either order). **25 conflicts with
nothing**, and 22↔23↔24↔25 are file-disjoint from each other. Full matrix in
[TASKLIST-architecture-deepening.md](TASKLIST-architecture-deepening.md).

Adding specs 26–34: four hard pairs, each sequential in the stated order — **26→33**
(`XLWorksheetRangeShifter.cs`, `XLWorksheet.cs` — 26 collapses the row/column duplication before 33
reorganises what is left), **27→34** (no shared file, but 34 moves metric computation across three
adapters and 27 is the only thing that would notice), **29→31** (`SheetViewWriter.cs`,
`ColumnWriter.cs` — never rebase a correctness fix onto a structural sweep) and **30→32**
(`FunctionDefinition.cs` — 32 removes the members 30's file reads). **28 is independent.** Against
the older specs: **31↔15/16/17** (`PictureWriter.cs`, `ChartWriter.cs` — hard; 31 waits or defers
that one writer), **32↔07 wave A2** (A2 would add registrations in the old form — decide the order
before starting either), **28↔24** (`WorksheetSheetDataReader.cs` — soft, 28 first is marginally
better), **28↔29** (`WorkbookStylesPartWriter.cs`, different regions — soft), **26↔14**
(`XLRangeBase.cs`, different methods — soft), **30↔04** (`CalculationVisitor.cs` — soft, 30 is much
smaller and goes first). **27 and 34 conflict with nothing.** Full matrix in
[TASKLIST-architecture-deepening-2.md](TASKLIST-architecture-deepening-2.md).

## Ground rules for implementing agents

- **Branch per spec/task; never commit to main.** Commit style follows the repo convention (`perf:`, `feat:`, `fix:`, `refactor:`).
- **Warnings are errors** (`TreatWarningsAsErrors=true`); nullable is enabled — new code must be null-annotated.
- **Do not upgrade SixLabors.Fonts** (license conflict).
- **No compound shell commands** (`&&`, `;`) in agent tool calls — repo convention (see `CLAUDE.md`).
- Build check: `dotnet build XLibur/XLibur.csproj -c Release -v q` · Tests: `dotnet test XLibur.Tests` · Benchmarks: `dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --filter '*Name*'` · Memory snapshots: same project with `-- profile` (writes dotMemory `.dmw` to `C:\profiles\`).
- **Perf PRs must include before/after numbers** (BenchmarkDotNet table + allocation delta) in the description, per the format of PR #171.
- Line numbers in specs are from the July 2026 survey — **verify against current code before editing**.
- File-format work (06, 09, 10) requires a recorded manual "opens clean in Excel" check plus automated reload-via-XLibur tests; Excel-authored test resources go in `XLibur.Tests/Resource/`.
