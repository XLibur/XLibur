# XLibur Improvement Roadmap — Top 10 Specs

Ten prioritized, self-contained specs covering features, compatibility, architecture, and performance (memory + read/write times). Each spec is written to be handed to an independent agent/model: it states the problem with measured numbers, points at the exact files, prescribes a design, breaks the work into PR-sized tasks, and defines measurable acceptance criteria.

Grounding: these were derived from a July 2026 survey of the codebase (architecture, feature inventory vs Excel, benchmark artifacts under `BenchmarkDotNet.Artifacts/results/`). Headline baselines: save 50K rows ≈ 1.0–1.1 s / **543 MB allocated**; load+read 250K×15 ≈ 5.6 s / 1.68 GB after PR #171; XLibur is already ~3× faster and ~6× leaner than upstream ClosedXML on save.

## The list

| # | Spec | Area | Effort | Status | Parallelizable? |
|---|------|------|--------|--------|-----------------|
| 01 | [Streaming write API](01-streaming-write-api.md) | Feature · Arch · Memory | L | Proposed | Phase 1 refactor first, then independent |
| 02 | [Load-path allocation elimination](02-load-path-allocations.md) | Perf (read) | M | ✅ **Done** (#175) | 3 independent sub-tasks |
| 03 | [Save-path allocation reduction](03-save-path-allocations.md) | Perf (write) | M | Proposed | 7 small independent PRs |
| 04 | [Demand-driven formula evaluation](04-demand-driven-formula-eval.md) | Perf · Arch | L | Proposed | Single owner (correctness-critical) |
| 05 | [Structural-edit & bulk-style scalability](05-structural-edit-scalability.md) | Arch · Perf | L | Proposed | 3 independent workstreams |
| 06 | [Workbook encryption (password files)](06-workbook-encryption.md) | Feature · Compat | L | Proposed | Container/crypto layers in parallel |
| 07 | [Formula function coverage (257→~420)](07-formula-function-coverage.md) | Feature | L | Proposed | **6 fully independent waves** |
| 08 | [LET / LAMBDA](08-let-lambda.md) | Feature | L | Proposed | Single owner (engine core) |
| 09 | [Threaded comments + round-trip fidelity](09-threaded-comments-roundtrip.md) | Feature · Compat | M | Proposed | Comments vs fidelity-audit split |
| 10 | [Chart formatting depth](10-chart-formatting-depth.md) | Feature | L | Proposed | 4 PRs, 2–3 independent |
| 11 | [Create-path allocation reduction](11-create-path-allocations.md) | Perf (write) | M | ✅ **Tasks 1–4 done** | Task 4 lands in 11; 05 rebases |

Spec 11 was added after spec 03 landed: 03 halved the *save* phase and showed the rest of it is
`System.IO.Packaging`, leaving the *create* phase as 72% of the write benchmark. It is a follow-on,
not part of the original ten.

Spec 02 delivered **−16.5% load time and −61.5% allocations** (4.750 s / 1020.92 MB → 3.968 s /
392.88 MB on the 250K×15 benchmark). It also produced two findings that change other specs: a
correction to spec 03's number-formatting task, and a reusable `XmlReader.ReadValueChunk` technique
for the IO layer. Both are recorded in spec 02's Results section.

## Why these ten

**Performance (specs 02, 03, 04, 05).** The write cell-loop and the sheetData parse have both had a round of tuning; what remains, in measured order: per-cell string allocations on load (`<v>` + attributes + SST DOM), the ~543 MB formatted save (number formatting, inherited-style resolution, StyleKey hashing), the full-workbook-recalc cliff when reading one dirty formula cell (can build a 176 MB dependency tree to answer one read), and O(all-ranges·log) work per single row insert. Specs 02–05 attack each with concrete targets.

**Features (specs 01, 07, 08, 10).** The fork already leads upstream on charts, dynamic arrays, in-cell images, sparklines, WebP/SVG. The gaps that matter: no bounded-memory export path for huge files (01), ~250 missing formula functions with a clean registry to extend (07), no LET/LAMBDA (08 — the one function family that needs engine work), and charts that can't be styled (10 — depth on the flagship differentiator).

**Compatibility (specs 06, 09).** Password-encrypted files can't be opened or written at all (06 — hard blocker, zero code exists). Threaded comments are read lossily and silently downgraded on save; chartsheets, form controls, and slicers are dropped on round-trip (09).

**Architecture (specs 01, 04, 05).** The three structural debts the surveys flagged: no streaming seam in the IO layer, calc-engine all-or-nothing recalculation, and materialize-everything patterns for range shift and style propagation (plus two redundant spatial indexes — QuadTree and RBush).

## Suggested sequencing

```
Wave 1 (independent, start anytime):
  02 load allocations ✅ done · 03 save allocations · 07 function waves A–F · 09 threaded comments
  11 Tasks 1–4 ✅ done (−28.8% on the write benchmark; bulk styling −86% per cell)
Wave 2 (after 03 lands, or coordinated):
  01 streaming write (Phase 1 seam shared with 03's territory) · 06 encryption · 10 charts PR1
Wave 3 (single-owner, correctness-critical — don't parallelize internally):
  04 demand-driven eval · 05 structural edits · 08 LET/LAMBDA (08 after or alongside 04)
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
