# XLibur Improvement Roadmap

Twenty prioritized, self-contained specs covering features, compatibility, architecture, and performance (memory + read/write times). Each spec is written to be handed to an independent agent/model: it states the problem with measured numbers, points at the exact files, prescribes a design, breaks the work into PR-sized tasks, and defines measurable acceptance criteria.

**Start a new performance effort at [spec 19](19-benchmark-hotspot-survey.md)**, not at this table. It re-ran the whole suite on 2026-08-07 and ranks what is actually slow now, which is not what specs 02–18 would predict — the biggest single number in the suite turns out to be the `CellsUsed()` enumeration, not parsing, packaging or styling. It also carries the current baselines for every benchmark and the run recipe.

Specs 01–10 are the original top-ten set; spec 11 is a follow-on that came out of implementing spec 03 (see below).

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
| 13 | [Public core surface for `XLibur.Report`](13-report-core-public-api.md) | Arch · API · Packaging | M | Proposed | Tasks 1 and 2 independent |
| 14 | [`Clear`/`CopyTo` scalability](14-clear-copyto-scalability.md) | Perf (edit) · Correctness | S | Proposed ([#271](https://github.com/XLibur/XLibur/issues/271)) | Task 1 first; 2/3/4 independent |
| 15 | [Shapes & text boxes](15-shapes-and-text-boxes.md) | Feature · Compat | L | Proposed (**needs 16 first**) | 1–3 one stream; then 4/5 parallel |
| 16 | [Shared DrawingML infrastructure](16-drawingml-infrastructure.md) | Arch · Refactor | S–M | Proposed | 3 PRs; harness first, then 2/3 independent |
| 17 | [Picture styling & fidelity](17-picture-styling.md) | Feature · Compat · **Defect** | M–L | Proposed (**needs 16 first**) | Task 1 (fidelity fix) first and standalone; 3/4/5 parallel after 2 |
| 18 | [Template round-trip overhead](18-template-round-trip-overhead.md) | Perf (read + write) · **Defect** | M | Tasks 0–4 done (see Results) | Task 5 is the remaining cost; independent |
| 19 | [Benchmark hotspot survey (Aug 2026)](19-benchmark-hotspot-survey.md) | Perf (read + write) · Survey | L | Proposed | 5 areas; 1/2/3 fully independent |
| 21 | [Hot-path struct candidates](21-hot-path-struct-candidates.md) | Perf (read · enumeration) | M | ✅ **Done** (task 3 shipped; 1–2 declined on measurement) | Task 0 first; 1→2 ordered; 3 independent |

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

## Ground rules for implementing agents

- **Branch per spec/task; never commit to main.** Commit style follows the repo convention (`perf:`, `feat:`, `fix:`, `refactor:`).
- **Warnings are errors** (`TreatWarningsAsErrors=true`); nullable is enabled — new code must be null-annotated.
- **Do not upgrade SixLabors.Fonts** (license conflict).
- **No compound shell commands** (`&&`, `;`) in agent tool calls — repo convention (see `CLAUDE.md`).
- Build check: `dotnet build XLibur/XLibur.csproj -c Release -v q` · Tests: `dotnet test XLibur.Tests` · Benchmarks: `dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --filter '*Name*'` · Memory snapshots: same project with `-- profile` (writes dotMemory `.dmw` to `C:\profiles\`).
- **Perf PRs must include before/after numbers** (BenchmarkDotNet table + allocation delta) in the description, per the format of PR #171.
- Line numbers in specs are from the July 2026 survey — **verify against current code before editing**.
- File-format work (06, 09, 10) requires a recorded manual "opens clean in Excel" check plus automated reload-via-XLibur tests; Excel-authored test resources go in `XLibur.Tests/Resource/`.
