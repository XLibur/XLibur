# XLibur Improvement Roadmap

Twelve prioritized, self-contained specs covering features, compatibility, architecture, and performance (memory + read/write times). Each spec is written to be handed to an independent agent/model: it states the problem with measured numbers, points at the exact files, prescribes a design, breaks the work into PR-sized tasks, and defines measurable acceptance criteria.

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
| 12 | [Report templating (`XLibur.Report`)](12-report-templating.md) | Feature · Arch | L | Proposed | 10 tasks; 4/5/6/10 parallel after 3 |

Spec 11 was added after spec 03 landed: 03 halved the *save* phase and showed the rest of it is
`System.IO.Packaging`, leaving the *create* phase as 72% of the write benchmark. It is a follow-on,
not part of the original ten.

Spec 12 (July 2026) is a feature spec outside the original survey: a report-templating package
(`XLibur.Report`) porting the ClosedXML.Report architecture with a Scriban expression engine,
an Excel-function bridge into `{{ }}` expressions, first-class chart/pivot/image handling
across range expansion, and an opt-in `XLibur.Report.DynamicLinq` package that runs
ClosedXML.Report's C#-expression template syntax unmodified. Its only touch on the core
library is an `InternalsVisibleTo` grant, so it does not conflict with the open perf specs.

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

Conflict map: 01↔03 (`SheetDataWriter`), 04↔08 (evaluation stack / `CalcContext`), 07 waves B↔C (`Statistical.cs`). Everything else is disjoint. **Spec 05 must rebase onto spec 11**: 11's Task 4 rewrote bulk style propagation (`XLStylizedBase.ModifyStyle` / `SetStyle`), which is 05's territory.

## Ground rules for implementing agents

- **Branch per spec/task; never commit to main.** Commit style follows the repo convention (`perf:`, `feat:`, `fix:`, `refactor:`).
- **Warnings are errors** (`TreatWarningsAsErrors=true`); nullable is enabled — new code must be null-annotated.
- **Do not upgrade SixLabors.Fonts** (license conflict).
- **No compound shell commands** (`&&`, `;`) in agent tool calls — repo convention (see `CLAUDE.md`).
- Build check: `dotnet build XLibur/XLibur.csproj -c Release -v q` · Tests: `dotnet test XLibur.Tests` · Benchmarks: `dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --filter '*Name*'` · Memory snapshots: same project with `-- profile` (writes dotMemory `.dmw` to `C:\profiles\`).
- **Perf PRs must include before/after numbers** (BenchmarkDotNet table + allocation delta) in the description, per the format of PR #171.
- Line numbers in specs are from the July 2026 survey — **verify against current code before editing**.
- File-format work (06, 09, 10) requires a recorded manual "opens clean in Excel" check plus automated reload-via-XLibur tests; Excel-authored test resources go in `XLibur.Tests/Resource/`.
