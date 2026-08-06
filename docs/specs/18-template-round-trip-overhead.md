# Spec 18 — Template Round-Trip Overhead (open → small edit → save)

**Area:** Performance (read + write time, memory)
**Effort:** M (task 1 is the bulk of the win; 2–4 are independent follow-ons)
**Dependencies:** Touches `WorksheetPartWriter` and the style façades. Coordinate with Spec 03 (both touch the save path) and Spec 11 (create-path allocations); no overlap with the `SheetDataWriter` cell loop, which is already tuned.
**Status:** Tasks 0–4 done — see [Results](#results). Task 5 open (the dominant remaining cost).

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

```bash
dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- --filter "*TemplateRoundTripBenchmarks*"
dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile template
dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile template loop save   # for a profiler
```

## Findings and work plan

| # | Task | Status | Size |
|---|------|--------|------|
| 0 | Benchmarks + decomposition probe that expose all of this | ✅ Done | S |
| 1 | Stop materialising `<sheetData>` into a DOM on save | ✅ Done (`26b248d9`) | M |
| 2 | Per-cell styling: CPU cost in the façade setters | ✅ Done (`4d98f127`) — allocation share found inherent; **win restated at −13% by task 4** | M |
| 3 | Lookup-column refresh costs ~3× per cell versus the grid path | ✅ Explained — inherent, no fix | S |
| 4 | Re-verify and file the remainder | ✅ Done | S |
| 5 | Each structurally empty worksheet costs ~2.2 ms / ~0.24 MB to round-trip | ⬜ Proposed | M |

---

### Task 0 — Measurement tooling ✅

- `XLibur.Benchmarks/TemplateRoundTripBenchmarks.cs` — the trustworthy numbers.
- `XLibur.Benchmarks/TemplateRoundTripProfile.cs` — `profile template [path.xlsx]` decomposes the round trip; `profile template loop <open|save|roundtrip>` runs one phase so a profiler can attach.
- `XLibur.Benchmarks/TemplateFixture.cs` — the shared synthetic template. `sheetCount` built one sheet
  too many until review caught it (`Math.Max(1, sheetCount - 1)` lookup sheets, so `sheetCount: 1` gave
  two). Only the `sheets=1` row was affected — 10 and 40 were always right — and the task 5 figures below
  are the corrected ones.

**Do not use the probe's timings to claim a change.** On the reference machine they move by tens of percent between runs of identical code — the same order as the effects being chased. The probe locates cost and reports exact allocation; BenchmarkDotNet proves movement. An early iteration of this work reported a 30% win that BenchmarkDotNet then showed to be ~0%.

**BenchmarkDotNet is not automatically trustworthy either** — it is only trustworthy when its warnings are clear. Task 2's win was overstated by 2.5× because its benchmark ran 2.7 ms iterations against the recommended 100 ms and divided by a baseline that scattered 2.7–6.1 ms; BDN said so in a `MinIterationTime` warning that went unread. Before believing a `Ratio`, check three things: that the `MinIterationTime` warning is absent or the iteration is comfortably long, that the *baseline's* removed outliers do not span an appreciable fraction of its mean, and — for anything with `[IterationSetup]` — that the previous iteration's garbage is collected in `[IterationCleanup]` rather than inside the next measurement.

**A/B the controls, not just the target.** The reliable signal in the task 2 re-run was not the two benchmarks that improved; it was the five that did not move while they did. A change that appears to shift every variant in the table, or that shifts variants it cannot reach, is measuring the machine.

### Task 1 — Stop materialising `<sheetData>` into a DOM on save ✅

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

**Acceptance criteria.** All met — see [Results](#results).
1. ✅ `OpenAndSaveRowHeavyUnchanged` allocation reduced ≥ 50% (333.68 MB → ≤ 165 MB); mean time reduced ≥ 40%.
2. ✅ `OpenAndSaveUnchanged` allocation not regressed.
3. ✅ Saved output semantically identical to main across the full corpus.
4. ✅ All tests green; no public API change.

**Outcome.** The part is copied into memory with `<sheetData>` reduced to an empty element, and the DOM is built from that copy. Rows are tokenised once by a raw reader rather than becoming objects nobody reads.

Assembling the root by hand and appending each top-level child — the design sketched above — was implemented first and is faster still, but it is **not byte-faithful**, and the risk anticipated above turned out to be the wrong one. The root's own prefix reproduced fine; what broke was descendants. A child parsed on its own records the namespace declarations it needs rather than inheriting the root's, so a part whose root uses a *default* namespace round-trips its `mc:AlternateContent` subtrees as `<controls xmlns="…">` where the file had `<x:controls>` (caught by `SavingTests.FormControlsArePreserved`). Two further traps found on the way, worth not re-deriving:

- `reader.Prefix` on `OpenXmlPartReader` is the *element's resolved* prefix, not the one the file used. It falls back to the SDK's canonical `x` for the spreadsheet namespace even when the part declared none.
- A *default* namespace declaration never appears in `reader.NamespaceDeclarations`. That list is filled by testing for the `xmlns` prefix, which `xmlns="…"` does not have, so it surfaces among the attributes under the reserved local name `xmlns`. `AddNamespaceDeclaration` cannot put it back either — it rejects an empty prefix.

Copying the part keeps the SDK doing one parse of one document, so the emitted bytes are unchanged by construction. The copy covers everything except `<sheetData>`; it measured at 991.7 ms against 917.4 ms for the unfaithful version, inside that version's own 76 ms standard deviation.

## Results

BenchmarkDotNet, net10.0 Release, before/after task 1:

| Benchmark | before | after | Δ time | Δ allocated |
|---|---|---|---|---|
| `OpenAndSaveRowHeavyUnchanged` | 1,729.6 ms / 333.68 MB | 991.7 ms / 88.83 MB | −43% | **−73%** |
| `OpenAndSaveUnchanged` | 10.42 ms / 3.89 MB | 8.42 ms / 3.17 MB | −19% | −19% |
| `RefreshLookupColumn` | 11.49 ms / 4.55 MB | 10.08 ms / 3.83 MB | −12% | −16% |
| `LoadRowHeavy` | 402.4 ms / 55.40 MB | 406.4 ms / 55.40 MB | — | — |

`LoadRowHeavy` is load-only and unaffected, as expected; it is listed as the control.

### Task 2 — Per-cell styling costs ~2.3× ✅ (partly; see below)

The symptom was corroborated twice:

- `profile template` — per-cell vs per-column number format is **2.1–2.6× time and a stable 2.3× allocation** across 1,000/5,000/20,000 rows.
- `profile alloc` — `CreateFormattedAndSave` allocates **183.3 MB** in the create phase against **25.1 MB** for the same rows unformatted.

**The mechanism first recorded here was wrong**, and it is worth saying why rather than quietly deleting it. It claimed the façade objects were the cost, because `Cell(r, c)` returns a fresh `XLCell` whose `_cachedStyle` is therefore always cold. Two things are wrong with that. `XLCellsCollection.GetCell` already keeps a direct-mapped `XLCell` cache; and, more importantly, the façades barely allocate.

`CellStylingBenchmarks` splits the path. Per 20,000 cells, against an unstyled write of the same values:

| step | time | allocated |
|---|---:|---:|
| `XLStyle` façade | 21 ns/cell | 76 B/cell |
| `XLNumberFormat` façade | 6 ns/cell | 31 B/cell |
| the `.Format = x` setter | **163 ns/cell** | 23 B/cell |
| *(for scale: assigning a pre-built `IXLStyle` instead)* | 59 ns/cell | 123 B/cell |

So the whole façade chain is ~2% of the styling **allocation** and ~14% of its **time**. The allocation is the style-slice write, and it is inherent: a distinct style per cell is N slice entries, exactly as `BulkStyleBenchmarks` already concluded for spec 05's criterion 3. **There is no allocation win available here.** Per-column styling stays 2.3× leaner because it writes one style, not 20,000.

What *was* available was CPU, in the setter. Every façade applied a component key by interning it and then handing it straight back:

```csharp
private void SetKey(XLNumberFormatKey newKey)
{
    Key = newKey;                    // XLNumberFormatValue.FromKey -> repository lookup
    _style.ModifyNumberFormat(Key);  // hashes the same key again
}
```

The lookup is waste in both directions `ModifyXxx` can go: on a transition-cache hit it never needs the component value, and on a miss it interns the component anyway inside `XLStyleValue.FromKey`. All six façades shared the pattern. Fixed in `4d98f127` by applying the modification first and taking the interned value back off the resulting style.

The fix is real but smaller than first reported. **Task 4 re-ran the A/B and the original figures did not
reproduce; they are corrected here.** What was recorded — 2.38 → 1.57, ≈328 → ≈219 ns per styled cell —
came from a benchmark whose iterations ran 2.7–5.8 ms against BenchmarkDotNet's recommended 100 ms, and
whose `ValueOnly` baseline scattered across 2.7–6.1 ms. Since every `Ratio` in the table divides by that
baseline, the reported win was mostly baseline movement. A first re-run disagreed with itself in opposite
directions on two benchmarks, which is what exposed it.

Re-measured at 100,000 rows with a forced gen2 collection in `IterationCleanup` (so the previous
iteration's workbook is reclaimed outside the measured region), A/B against `4d98f127~1` with the
benchmark project byte-identical across both arms:

| Benchmark | setter per cell | before | after | Δ |
|---|---|---:|---:|---:|
| `StyleFacadePerCell` | 1× | 48.26 ms | **43.76 ms** | −4.50 |
| `TwoPropertiesPerCell` | 2× | 58.95 ms | **54.23 ms** | −4.72 |
| `ValueOnly` *(control)* | — | 24.57 ms | 23.05 ms | −1.52 |
| `FacadeStyleOnly` *(control)* | — | 42.00 ms | 42.46 ms | +0.46 |
| `FacadeChainOnly` *(control)* | — | 40.22 ms | 40.37 ms | +0.15 |
| `StyleAssignedPerCell` *(control)* | — | 43.46 ms | 43.28 ms | −0.18 |
| `StyleOnColumn` *(control)* | — | 18.11 ms | 18.13 ms | +0.02 |

The five controls never call a façade setter and scatter within ±1.5 ms in both directions; the two that
do call it both improve, with non-overlapping error bars. Measured against each run's own baseline the
styling overhead falls **237 → 207 ns/cell** for one property (−13%) and 344 → 312 ns/cell for two
(−9%). Allocation is unchanged, as expected for a pure lookup elision.

**The per-step split above is regime-dependent — do not read it as fixed.** Re-measuring at 100,000 rows
inverts it: `FacadeStyleOnly` rises to ratio 1.85 and `FacadeChainOnly` to 1.76, i.e. building the façade
and setting *nothing* costs nearly as much as the whole styled write, and `FacadeChainOnly` comes out
*cheaper* than `FacadeStyleOnly` despite doing strictly more work. The GC columns explain it: every styled
variant runs 2× the baseline's Gen0 collections plus Gen1 collections the baseline never triggers, so past
~20K styled cells the measurement is dominated by reclaiming the extra allocation rather than by the style
code. The 21/6/163 ns attribution holds at 20K rows and nowhere else. Allocation ratios, by contrast, are
stable across both sizes (1.52 / 1.49 / 1.74 / 1.31 / 1.43 / 0.65) — that part of the conclusion was always
sound.

**Still open.** Roughly 207 ns/cell of styling overhead remains against 20 ns/cell for the value write.
Whether any of it is reachable is unmeasured; the transition machinery is the obvious suspect, but see the
regime warning above before trusting any per-step number that motivates chasing it.

**Noted while here, not fixed:** `XLStyleValue.GetTransition` matches on the 32-bit transition hash alone and never compares the key, so two component keys sharing a hash *and* a cache slot would hand back the wrong style. The window is small — an 8-entry direct-mapped cache — and the risk predates this work, but it is a correctness bug rather than a performance one and deserves its own issue.

### Task 3 — Lookup refresh costs ~3× per cell ✅ Explained — not a defect

`profile template`, lookup refresh: the 1,000 → 10,000 step costs ≈5.4 µs per cell where the grid path runs ≈1.3 µs per cell for write + save. Two causes were proposed and neither had evidence: shared-string pressure, because the lookup values are all distinct; or the sheet's geometry, because a single column pays whatever a row costs once per cell instead of once per twenty.

`SheetGeometryBenchmarks` crosses the two — 20,000 string cells in every variant, so a difference is a difference *per cell*:

| variant | mean | allocated |
|---|---:|---:|
| `TallNarrow_Unique` | 40.73 ms | 17.42 MB |
| `TallNarrow_Repeated` | 29.71 ms | 9.96 MB |
| `ShortWide_Unique` | 30.74 ms | 12.85 MB |
| `ShortWide_Repeated` | 17.70 ms | 5.50 MB |

**Both effects are real, roughly equal, independent, and they compound.** Geometry costs 25–40%; string uniqueness costs 27–42%; together the worst quadrant is 2.3× the best in time and 3.2× in allocation. Neither candidate was "the" cause, which is why crossing them mattered — either comparison alone would have confounded the two and produced a confident wrong answer.

Splitting the geometry penalty by phase (unique strings) locates it:

| phase | tall narrow | short wide | penalty |
|---|---:|---:|---:|
| write only | 10.45 ms / 7.45 MB | 8.38 ms / 6.15 MB | +2.07 ms — 19% |
| save (by subtraction) | 30.98 ms / 9.97 MB | 22.12 ms / 6.70 MB | **+8.86 ms — 81%** |

So it is mostly the `<row>` element and its bookkeeping, ~443 ns per row, not the row storage. The string effect is a flat ~7.4 MB per 20,000 distinct values (~373 B each) whichever way the sheet is shaped, exactly as one shared-string entry plus one `<si>` per distinct value should behave.

**Conclusion: there is nothing to fix.** Both costs are inherent to the data's shape — a row costs what a row costs, and a distinct string has to be stored and written once. The lookup refresh simply sits in the worst quadrant of both while the grid sits near the best: 21 columns wide, and only about a quarter of its cells are unique strings (the rest are numbers, dates and a repeated `"Yes"`/`"No"`). That accounts for the ~4× without any defect. `SheetDataWriter.WriteStartRow` was inspected and is already tight — it early-outs before touching row attributes when no `XLRow` exists.

**Guidance rather than a code change:** cost scales with rows and with *distinct* values, not with cells. A caller refreshing a tall single-column lookup is on the most expensive path there is, and the lever available to them is the data, not the library.

### Task 4 — Re-verify and file the remainder ✅

**Tasks 1 and 3 reproduce exactly.** All five `TemplateRoundTripBenchmarks` match their recorded post-fix
values with allocation identical to the byte (`OpenAndSaveRowHeavyUnchanged` 921.3 ms / 88.83 MB,
`LoadRowHeavy` 389.8 ms / 55.4 MB, `Open` 4.22 ms / 1.67 MB, `OpenAndSaveUnchanged` 7.89 ms / 3.17 MB,
`RefreshLookupColumn` 9.91 ms / 3.83 MB). All six `SheetGeometryBenchmarks` land within noise of theirs,
allocation again identical to the byte. Times sit 5–7% below the recorded figures across the board, which
is machine variance, not a new win — identical allocation says it is the same code path.

**Task 2 did not reproduce and has been corrected above**, along with the benchmark that produced it.

**Filed as task 5:** per-worksheet round-trip cost, the dominant remaining term.

**Neither defined names nor data validations allocate measurably**, the two plausible suspects from the
profiled application. Against a 10-sheet fixture: 0 names → 3.2 MB, 20 names → 3.3 MB; then 0 → 26 → 100
validations holds flat at 3.3 MB. Allocation is reported exactly by the probe, so this part is solid.

**Their time cost is not established, and one run should not be trusted here.** Three decompositions over a
byte-identical fixture (22,725 B every time) put 20 names at 28.5 ms against 30.6 ms for none, then 46.3
against 27.2, then 25.9 against 27.6. Two of the three say free and one says 70% dearer; nothing changed
but the machine. Anyone wanting a time answer for names or validations needs a BenchmarkDotNet case, not
this probe — see the measurement warning under task 0, which exists because this spec has already published
figures that did not survive re-measurement.

What the sheet-count slope shows is robust because it is read across three points that all move together
and is corroborated by allocation. The rest of the table is not.

Left unchased by scope, as instructed: `System.IO.Packaging` buffering and `XmlWriter` internals (Spec 01/03).

### Task 5 — A worksheet costs ~1.0 ms and ~0.19 MB structurally, before any row on it ⬜ (attributed)

> **Correction.** This task first claimed ~2.2 ms per *structurally empty* worksheet. It was measuring
> sheets that each held 101 rows: `TemplateFixture` hard-coded `LookupRows = 100`, so sheet count and row
> count moved together and the slope charged the rows to the sheet. `LoadRowXml` and `LoadCellXml` being
> 18.6% of a trace of a fixture described as having no data is what exposed it. `lookupRows` is now a
> parameter and the two are separated below.

`profile template`, two sheet-count sweeps differing only in rows per lookup sheet:

| sheets | header-only | 100 rows/sheet |
|---:|---:|---:|
| 1 | 6.3 ms / 1.1 MB / 6,437 B | 7.5 ms / 1.1 MB / 6,437 B |
| 10 | 15.1 ms / 2.8 MB / 12,766 B | 27.6 ms / 3.2 MB / 22,581 B |
| 40 | 46.8 ms / 8.5 MB / 33,749 B | 94.4 ms / 10.6 MB / 76,074 B |

| | per sheet, time | per sheet, alloc |
|---|---:|---:|
| **structural (header-only)** | **~1.0 ms** | **~0.19 MB** |
| the 100 rows on top | ~1.2 ms | ~0.05 MB |
| *(total, as previously reported)* | ~2.2 ms | ~0.24 MB |

Both slopes hold across both intervals — 0.98 and 1.06 ms/sheet header-only, 2.23 and 2.23 with rows — so
the split is real. At `sheets=1` the two sweeps build a byte-identical fixture (6,437 B) and still measured
6.3 against 7.5 ms, which sets the probe's noise floor at ~1.2 ms here; the slopes come from differences of
8.8 and 31.7 ms, well clear of it.

**The structural half is the interesting one.** A header-only worksheet adds ~703 B to the file and ~0.19 MB
of allocation — roughly **270× its on-disk size**, for a sheet holding one cell. Allocation is where the
split is starkest: rows contribute only ~21% of the per-sheet bytes. Unlike rows and distinct strings
(task 3), this scales with nothing the caller can reduce.

#### Attribution

`dotnet-trace --format speedscope` over `profile template loop roundtrip` (10 sheets, 20 names, 26
validations), 21,253 ms sampled. Inclusive shares of the whole trace:

| | share | what it is |
|---|---:|---|
| `LoadWorksheetElements` | 23.7% | per sheet — **XLibur** |
| `GenerateWorksheetPartContent` | 10.2% | per sheet — **XLibur** |
| zip deflate/inflate + the `Monitor` contention under it | ~8.5% | per part — packaging |
| `Path.GetTempFileName` + `CreateFile` | 7.8% | **once per package**, not per sheet |

**The per-sheet cost is roughly 4:1 XLibur's own work over packaging, so task 5 is in scope** rather than
Spec 01/03 territory as originally guessed. Splitting the two was the point of the trace and it came out
the opposite way to the assumption.

Two things worth recording separately:

- **Array growth dominates.** `Thread.PollGCWorker` is 30.3% of the trace, and its call sites are almost
  entirely `Array.Resize` / `Array.Copy` / `Buffer.Memmove` — `LoadCellXml` resizing (11.5%),
  `Slice<T>.Lut<T>.SetValue` resizing (10.7%), and `XmlTextReaderImpl.AllocNode` growing the reader's node
  buffer (10.1%). Time parks on the GC poll at safepoints inside those copy loops, so this is the
  allocation the benchmarks measure showing up as wall clock. Pre-sizing the per-sheet slices is the
  obvious first move, and it is squarely XLibur's code.
- **The temp file is the SDK's, and it is per workbook.** `WriteableStreamExtensions.EnableWriteableStream`
  calls `Path.GetTempFileName` to make the package writeable — 7.8% of the trace, but fixed per round trip
  rather than scaling with sheets, so it is not part of this slope. Worth its own note; `GetTempFileName`
  on Windows probes for an unused name and degrades as `%TEMP%` fills.

## Already ruled out

- **The two-pass loader is not worth collapsing.** `LoadWorksheetElements` reads structural elements with the SDK reader and `LoadSheetDataRaw` re-opens the part for a raw pass over `<sheetData>`. The second stream open plus rescan measures **0.05–0.65 ms per load** across sheet shapes, because everything ahead of `<sheetData>` is small. The waste was entirely in *getting past* `<sheetData>` in pass 1, which is fixed.
- **The `SheetDataWriter` cell loop.** Spec 03 tasks 3–6 are implemented: raw slice enumerator (no `XLCell` per cell), table-totals guard, single-entry style memo, blank short-circuit.
- **`SaveAs` is not single-use** and does not dispose the source stream, contrary to a note inherited with the original harness. It adopts its destination as the workbook origin, so a second `SaveAs` throws `ObjectDisposedException` only if the caller disposed that previous destination. Behaviour is coherent; only the diagnostic is poor.

## Measurement protocol

Every PR in this spec carries a before/after BenchmarkDotNet table for at least `OpenAndSaveRowHeavyUnchanged`, `OpenAndSaveUnchanged` and `LoadRowHeavy`. A/B a library change by stashing only the library (`git stash push -- XLibur/`) so the benchmark project is byte-identical across both runs.
