# Spec 18 — Template Round-Trip Overhead (open → small edit → save)

**Area:** Performance (read + write time, memory)
**Effort:** M (task 1 is the bulk of the win; 2–4 are independent follow-ons)
**Dependencies:** Touches `WorksheetPartWriter` and the style façades. Coordinate with Spec 03 (both touch the save path) and Spec 11 (create-path allocations); no overlap with the `SheetDataWriter` cell loop, which is already tuned.
**Status:** Task 0 done; task 1 in progress.

## Summary

The cost of *touching* an existing workbook — open it, change little or nothing, save it again — is dominated by work that scales with the workbook's structure and stored contents rather than with what the caller changed. A request-per-export service pays this on every request no matter how little it writes.

The trigger was a profiled export in a consuming application: ~300 ms per request went on opening and re-saving a 124 KB template (~10 sheets, ~20 defined names, ~26 data validations) purely to write one column of lookup values.

The headline defect: **saving a workbook that was loaded from a file re-materialises every stored row and cell as an OpenXML DOM, then throws it away.** Re-saving an untouched 20,000 × 21 workbook costs 1.73 s and 334 MB, of which roughly 1.33 s and 278 MB is the save half — for zero changes.

## Measured baselines

BenchmarkDotNet, `XLibur.Benchmarks.TemplateRoundTripBenchmarks`, net10.0 Release:

| Benchmark | Mean | Allocated |
|---|---:|---:|
| `Open` (10 sheets × 100 rows, 20 names, 26 validations) | 4.67 ms | 1.67 MB |
| `OpenAndSaveUnchanged` (same fixture) | 10.42 ms | 3.89 MB |
| `RefreshLookupColumn` (1,000 values) | 11.49 ms | 4.55 MB |
| `LoadRowHeavy` (20,000 × 21, stored) | 402.4 ms | 55.40 MB |
| **`OpenAndSaveRowHeavyUnchanged`** | **1,729.6 ms** | **333.68 MB** |

Reference point for task 1: writing those same 420,000 cells into a workbook whose *stored* part is empty saves in **428 ms** (`profile template`, "grid save"). The gap to the 1,330 ms save half above is the discarded DOM, not the writing.

Reproduce with:

```
dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- --filter "*TemplateRoundTripBenchmarks*"
dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile template
dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile template loop save   # for a profiler
```

## Findings and work plan

| # | Task | Status | Size |
|---|------|--------|------|
| 0 | Benchmarks + decomposition probe that expose all of this | ✅ Done | S |
| 1 | Stop materialising `<sheetData>` into a DOM on save | 🔵 In progress | M |
| 2 | Per-cell styling costs ~2.3× because the wrapper caches never hit | ⬜ Proposed | M |
| 3 | Lookup-column refresh costs ~3× per cell versus the grid path — cause unknown | ⬜ Needs investigation | S |
| 4 | Re-verify and file the remainder | ⬜ Proposed | S |

---

### Task 0 — Measurement tooling ✅

- `XLibur.Benchmarks/TemplateRoundTripBenchmarks.cs` — the trustworthy numbers.
- `XLibur.Benchmarks/TemplateRoundTripProfile.cs` — `profile template [path.xlsx]` decomposes the round trip; `profile template loop <open|save|roundtrip>` runs one phase so a profiler can attach.
- `XLibur.Benchmarks/TemplateFixture.cs` — the shared synthetic template.

**Do not use the probe's timings to claim a change.** On the reference machine they move by tens of percent between runs of identical code — the same order as the effects being chased. The probe locates cost and reports exact allocation; BenchmarkDotNet proves movement. An early iteration of this work reported a 30% win that BenchmarkDotNet then showed to be ~0%.

### Task 1 — Stop materialising `<sheetData>` into a DOM on save 🔵

**Current state.** `XLibur/Excel/IO/WorksheetPartWriter.cs`, `GetWorksheetDom` (~line 42):

```csharp
using var reader = OpenXmlReader.Create(worksheetPart);
if (!reader.Read()) throw new ArgumentException(...);
worksheet = (Worksheet)reader.LoadCurrentElement()!;
```

`LoadCurrentElement()` on `<worksheet>` builds the **entire** worksheet DOM, including every `<row>` and `<c>` in `<sheetData>`, as OpenXML objects. `StreamToPart` (~line 239) then walks the top-level children and, for the `SheetData` child, **ignores it** and calls `SheetDataWriter.StreamSheetData` instead. Every one of those row and cell objects is discarded.

The intent is already stated in the code (~line 108): *"Sheet data is not updated in the Worksheet DOM here, because it is later being streamed directly to the file without an intermediate DOM representation. This is done to save memory, which is especially problematic for large sheets."* That intent is defeated whenever the part is non-empty — i.e. on every save of a workbook that was loaded, and on every save after the first.

A trace of `profile template loop save` puts `GetWorksheetDom` at **41% of `SaveAs`** even on the 10-sheet fixture whose sheets hold only 100 rows each.

**Design.** Build the detached DOM child by child instead of in one call, substituting an empty `SheetData`:

1. Position the reader on `<worksheet>`; create `new Worksheet()`.
2. Copy the root's namespace declarations and attributes from the reader.
3. Walk the top-level children. For `SheetData`, append `new SheetData()` and skip the subtree with `ReadNextSibling()` — one raw `XmlReader.Skip`, no per-row object. For everything else, `LoadCurrentElement()` and append.

Everything downstream is satisfied by an empty `SheetData`: `worksheet.Elements<SheetData>().First()` still finds it, `XLWorksheetContentManager` still anchors on it, and `StreamToPart`'s `child is SheetData` test still fires in the right position.

This is the save-side twin of the load-side fix already made in `XLWorkbook_Load.LoadWorksheetElements`, where `<sheetData>` was being skipped one row at a time.

**Risks.**
- **Root element identity.** `StreamToPart` writes the root from `worksheet.Prefix`, `LocalName`, `NamespaceUri`, `NamespaceDeclarations` and `GetAttributes()`. A hand-built `Worksheet` must reproduce all of these. Prefix is the awkward one: most producers use the default namespace, but a `x:worksheet` root exists in the wild and would be re-emitted unprefixed. Same effective namespace, different bytes — acceptable for correctness, but confirm against the round-trip corpus.
- **Unknown/`AlternateContent` top-level children** must still round-trip; they go through `LoadCurrentElement()` unchanged.
- **Empty part** (`partIsEmpty`) path is untouched.

**Acceptance criteria.**
1. `OpenAndSaveRowHeavyUnchanged` allocation reduced ≥ 50% (333.68 MB → ≤ 165 MB); mean time reduced ≥ 40%.
2. `OpenAndSaveUnchanged` allocation not regressed.
3. Saved output semantically identical to main across the full corpus — `ExcelDocsComparerTests` and the round-trip suites green.
4. All tests green; no public API change.

### Task 2 — Per-cell styling costs ~2.3× ⬜

Corroborated twice:

- `profile template` — per-cell vs per-column number format is **2.1–2.6× time and a stable 2.3× allocation** across 1,000/5,000/20,000 rows.
- `profile alloc` — `CreateFormattedAndSave` allocates **183.3 MB** in the create phase against **25.1 MB** for the same rows unformatted.

**Mechanism.** `XLStylizedBase` caches its `XLStyle` façade in `_cachedStyle`, and `XLStyle` caches `_cachedNumberFormat`; `XLStyleValue` additionally memoises key transitions (`StoreTransition`). The machinery is already designed to avoid this cost — but `worksheet.Cell(r, c)` returns a **fresh `XLCell` on every call**, so those per-object caches are always cold. Each formatted cell therefore allocates the cell plus a fresh façade chain.

Options, in rough order of appeal: cache the façade per worksheet keyed by nothing (re-target the wrapper's container on access); pool/cache `XLCell` instances from `Cell(r, c)`; or add a bulk styling API and document per-column styling as the fast path. Needs a design decision before implementation — note that `Cell(r,c)` returning a stable instance has semantics implications well beyond styling.

### Task 3 — Lookup refresh costs ~3× per cell ⬜

`profile template`, lookup refresh: 10,000 values cost ~38 ms above the round-trip floor, ≈3.8 µs per cell, where the grid path runs ≈1.3 µs per cell for write + save. Candidate causes, none yet confirmed: 10,000 *unique* strings pressuring the shared-string table where the grid repeats `"Yes"`/`"No"`; or extending the used range from 100 to 10,000 rows interacting with the 26 data validations and 20 defined names on the fixture.

**Measure before designing.** This is the one finding here with no established mechanism.

### Task 4 — Re-verify and file the remainder ⬜

After tasks 1–3, re-run the benchmark set and the `profile template` decomposition. File follow-ups for anything ≥ 5% of what remains. Note but do not chase `System.IO.Packaging` buffering and `XmlWriter` internals — that is Spec 01/03 territory.

## Already ruled out

- **The two-pass loader is not worth collapsing.** `LoadWorksheetElements` reads structural elements with the SDK reader and `LoadSheetDataRaw` re-opens the part for a raw pass over `<sheetData>`. The second stream open plus rescan measures **0.05–0.65 ms per load** across sheet shapes, because everything ahead of `<sheetData>` is small. The waste was entirely in *getting past* `<sheetData>` in pass 1, which is fixed.
- **The `SheetDataWriter` cell loop.** Spec 03 tasks 3–6 are implemented: raw slice enumerator (no `XLCell` per cell), table-totals guard, single-entry style memo, blank short-circuit.
- **`SaveAs` is not single-use** and does not dispose the source stream, contrary to a note inherited with the original harness. It adopts its destination as the workbook origin, so a second `SaveAs` throws `ObjectDisposedException` only if the caller disposed that previous destination. Behaviour is coherent; only the diagnostic is poor.

## Measurement protocol

Every PR in this spec carries a before/after BenchmarkDotNet table for at least `OpenAndSaveRowHeavyUnchanged`, `OpenAndSaveUnchanged` and `LoadRowHeavy`. A/B a library change by stashing only the library (`git stash push -- XLibur/`) so the benchmark project is byte-identical across both runs.
