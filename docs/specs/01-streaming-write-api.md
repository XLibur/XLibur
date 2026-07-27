# Spec 01 — Streaming (Forward-Only) Write API for Large Workbooks

**Area:** Feature + Architecture + Performance (memory)
**Effort:** L (2–4 weeks)
**Dependencies:** None (but coordinate with Spec 03 — both touch `SheetDataWriter`)
**Status:** ✅ **Done** — see [Results](#results-2026-07-27) below.

## Summary

Add a forward-only, row-at-a-time streaming writer (analogous to POI's SXSSF or OpenXmlWriter) so callers can produce arbitrarily large .xlsx files with bounded memory. Today the entire workbook must be materialized in the in-memory slice model before `SaveAs` runs; a 5M-row export costs GBs of resident memory even though the write path itself is already streaming XML.

## Motivation

- `SheetDataWriter.StreamSheetData` already streams `<sheetData>` via raw `XmlWriter` — but it can only read from a fully materialized `XLWorksheet`. The storage, not the serializer, is the memory ceiling.
- Benchmarks: 50K rows × 3 cols save allocates ~543 MB; resident slice memory is ~16–32 bytes/cell plus styles. At 5M+ rows the in-memory model is the blocker, not CPU.
- Competing libraries (EPPlus streaming, MiniExcel, POI SXSSF) win large-export scenarios on this feature alone.

## Current state (file pointers)

- `XLibur/Excel/IO/SheetDataWriter.cs` (750 lines) — hot loop; takes concrete `XLWorksheet`, `SaveContext`, `SaveOptions`. Signature: `StreamSheetData(XmlWriter, XLWorksheet, SaveContext, SaveOptions)`. No abstraction seam.
- `XLibur/Excel/IO/WorksheetPartWriter.cs` — builds detached OpenXML DOM for non-sheetData elements, then re-emits via its own `XmlWriter` over the part stream, splicing in `StreamSheetData`.
- `XLibur/Excel/Cells/SharedStringTable.cs` (232 lines) — workbook-global, refcounted; SST is written after sheets by `SharedStringTableWriter.cs`.
- `XLibur/Excel/XLWorkbook_Save.cs` (1002 lines) — `CreateParts` orchestration.
- All IO classes are `internal static`. No async, no `CompressionLevel` control (zip goes through the OpenXML SDK / `System.IO.Packaging`).

## Design

### Phase 1 — Extract the data-source seam (internal refactor)

Introduce an internal interface consumed by `SheetDataWriter`:

```csharp
internal interface IXLSheetDataSource
{
    // Forward-only enumeration of non-empty rows in ascending row order.
    IEnumerable<XLRowData> EnumerateRows();
}
```

where `XLRowData` exposes row number, row properties (height/style/hidden), and a forward-only cell enumerator yielding `(column, XLCellValue value, styleId, formula?)`. Adapt the existing `SlicesEnumerator` path into an implementation backed by `XLWorksheet`. **Zero behavior change; all existing tests must pass unchanged.** Benchmark `CreateAndSave` before/after — regression budget ≤ 2% time, 0% allocation.

### Phase 2 — Public streaming writer

New public entry point (new file area `XLibur/Excel/Streaming/`):

```csharp
public sealed class XLStreamingWorkbook : IDisposable
{
    public static XLStreamingWorkbook Create(Stream output, XLStreamingOptions? options = null);
    public XLStreamingWorksheet AddWorksheet(string name);
    public void Finish();   // writes SST, styles, workbook part, closes package
}

public sealed class XLStreamingWorksheet
{
    public IXLStyle RowStyle { get; }           // style applied to next appended row
    public void AppendRow(params XLCellValue[] values);
    public void AppendRow(ReadOnlySpan<XLCellValue> values, IXLStyle? style = null);
    public void SkipRows(int count);
    public void Complete();                      // closes the worksheet part
}
```

Constraints (document clearly): rows append-only in ascending order; one worksheet open at a time; no formula recalculation (formula strings pass through verbatim); no reading back.

Implementation notes:
- Reuse `SheetDataWriter` cell-serialization internals via the Phase-1 seam — a streaming worksheet is just another `IXLSheetDataSource` whose enumeration is driven by the caller.
- SST strategy is an option: `SharedStrings` (default, dictionary built incrementally, written at `Finish()`) or `InlineStrings` (zero SST memory, larger files). The dictionary is the only unbounded memory in shared mode — call this out in docs.
- Styles: accept `IXLStyle` and intern to `XLStyleValue` via the existing repositories; styles part written at `Finish()`.
- Worksheet parts must be written before the workbook part is finalized — the OpenXML SDK supports adding parts incrementally; if part-stream lifetime under `System.IO.Packaging` proves awkward, fall back to writing sheet XML to temp `FileStream`s and copying into the package at `Finish()` (POI approach). Prototype this risk first (task P2.1).

### Phase 3 (stretch) — `CompressionLevel` + async save

While in this area: expose `SaveOptions.CompressionLevel` if feasible via the SDK, and add `SaveAsAsync`/`FinishAsync`. If `System.IO.Packaging` blocks both, record findings in the spec and split into a follow-up.

## Work plan

| # | Task | Size |
|---|------|------|
| P1.1 | Define `IXLSheetDataSource`/`XLRowData`, adapt `SheetDataWriter` to consume it | M |
| P1.2 | `XLWorksheet`-backed implementation over `SlicesEnumerator`; benchmark parity | M |
| P2.1 | Spike: incremental part writing under OpenXML SDK 3.x (or temp-file fallback) | S |
| P2.2 | `XLStreamingWorkbook`/`XLStreamingWorksheet` public API + options | M |
| P2.3 | SST incremental build + inline-strings mode | S |
| P2.4 | Style interning for streaming rows | S |
| P2.5 | Tests: round-trip via full `XLWorkbook.Load`, Excel-opens-clean validation, 1M-row memory test (assert peak working set bound) | M |
| P2.6 | Benchmark class `StreamingWriteBenchmarks` (vs `CreateAndSave` and vs raw OpenXML SDK) | S |
| P3.1 | CompressionLevel / async spike + implementation or writeup | M |

## Acceptance criteria

1. Writing 1,000,000 rows × 10 cols via the streaming API stays under **150 MB peak managed heap** (measure with the `MemoryProfile.cs` harness pattern).
2. Output opens in Excel with no repair dialog; `XLWorkbook.Load` round-trips values, styles, and formula strings.
3. Streaming write of the 50K-row benchmark scenario is ≥ as fast as `CreateAndSave` and allocates ≤ 25% of it.
4. No public-API breaks; `PublicApiAnalyzers` updated for additions only.
5. Existing test suite green; new tests in `XLibur.Tests/Excel/Streaming/`.

## Risks

- OpenXML SDK part-stream lifetime may force the temp-file fallback (adds disk I/O; still bounded memory). De-risk with task P2.1 first.
- API design is permanent — keep surface minimal (append-only) for v1; no cell-level random access.

## References

- Architecture survey: `SheetDataWriter` has no seam; `ISlice` is a shift contract, not a data-source abstraction.
- Benchmarks: `XLibur.Benchmarks/XLiburWorkbookBenchmarks.cs`, results under `BenchmarkDotNet.Artifacts/results/`.

## Results (2026-07-27)

All five acceptance criteria met. Measured on net10.0 with
`dotnet run -c Release --project XLibur.Benchmarks -f net10.0 -- profile streaming` and
`-- --filter "*StreamingWriteBenchmarks*"`.

### Acceptance criteria

| # | Criterion | Result |
|---|---|---|
| 1 | 1M × 10 under 150 MB peak managed heap | **107.9 MB** shared strings, **13.9 MB** inline ✅ |
| 2 | Opens clean; `XLWorkbook.Load` round-trips values, styles, formulas | `OpenXmlValidator` clean, round-trip tests green ✅ |
| 3 | 50K scenario ≥ as fast as `CreateAndSave`, ≤ 25% of its allocations | **1.59× faster**, **20.1%** of allocations ✅ |
| 4 | No public-API breaks; `PublicApiAnalyzers` additions only | Additions only ✅ |
| 5 | Suite green; new tests in `XLibur.Tests/Excel/Streaming/` | 7432 tests, 25 new ✅ |

### 50K × 3 workload

| Writer | Mean | Allocated |
|---|---|---|
| `XLWorkbook.CreateAndSave` | 250.5 ms | 67.46 MB |
| **`XLStreamingWorkbook`** | **157.9 ms** | **13.59 MB** |
| `XLStreamingWorkbook`, inline strings | 144.3 ms | 8.70 MB |
| `XLStreamingWorkbook`, `CompressionLevel.Fastest` | 91.9 ms | 17.00 MB |
| Raw OpenXML SDK (practical floor for SDK-based libraries) | 205.2 ms | 142.26 MB |

The streaming writer is faster and ~10× leaner than the raw SDK baseline, because it does not
go through `System.IO.Packaging` at all.

### 1M × 10, peak managed heap

| Writer | Peak heap | Elapsed | File |
|---|---|---|---|
| `XLStreamingWorkbook`, shared strings | 107.9 MB | 10.6 s | 36.1 MB |
| `XLStreamingWorkbook`, inline strings | 13.9 MB | 8.8 s | 35.5 MB |
| `XLWorkbook` **at 100K × 10** (a tenth of the size) | 126.7 MB | 2.2 s | 3.5 MB |

Every row in this workload carries a distinct string, the worst case for the shared string
table — the 108 MB is almost entirely that dictionary, and the inline figure is what the
documented escape hatch buys. Retained (post-collection) heap is ~30 KB at 400K rows either way.

### What changed against the plan

**Phase 1 — the `IXLSheetDataSource` seam was not built.** `StreamSheetData` is a single pass
over a struct slice enumerator with no per-cell allocation, the product of specs 03 and 11. An
`IEnumerable<XLRowData>` seam would add an interface dispatch per cell and a row object per row
to the existing save path, missing this spec's own ≤2% time / 0% allocation parity budget — and
the abstraction would have to be wide enough to carry formula, misc metadata, rich text and
table-totals membership, none of which the streaming side has. Instead the *leaf* serializers
(row start, cell start, value writers, type mapping) were extracted into `CellXmlWriter` and are
shared by both producers. Same reuse, no change to the hot loop.

**Phase 2 — the package is written by hand, not through the OpenXML SDK.** Task P2.1 asked
whether parts can be written incrementally under SDK 3.x. They can, and the spike confirmed it —
but only for *correctness*. Memory told a different story: `System.IO.Packaging` opens a package
read/write, which maps to `ZipArchiveMode.Update`, and that mode buffers every part's
**uncompressed** bytes until the archive closes. Measured 48 MB of live heap for 100K × 10,
growing linearly — the exact failure the API exists to prevent, and the same ~96 MB hotspot
[spec 03](03-save-path-allocations.md) flagged and deferred here. Opening the package write-only
*does* select `ZipArchiveMode.Create`, but the SDK enumerates parts and throws
`Cannot retrieve parts of writeonly container`.

So `StreamingPackageWriter` assembles the OPC package directly over `ZipArchive` in Create mode.
This is a different resolution than the spec's predicted fallback (staging sheet XML in temp
files): temp files would not have helped, because the copy into the package would have been
buffered just the same. Writing the zip ourselves also means no temp disk at all, and two
capabilities fall out for free — the output need not be **seekable** (a workbook can be written
straight to a network stream, which `XLWorkbook.SaveAs` cannot do) and `CompressionLevel` is
directly available.

**Two design points settled before implementation**, both because the spec's API sketch could
not express what its own constraints promised:

- The row builder `XLStreamingRow` was added alongside `AppendRow`. The sketched surface had no
  way to write a **formula** (despite the stated "formula strings pass through verbatim"
  constraint) or a **per-cell style**, and `AppendRow(params XLCellValue[])` allocates an array
  per row, working against the memory goal.
- Streamed worksheets support column widths, freeze panes and an autofilter range, not just a
  name. All are O(1) memory and written around `sheetData`; without them a streamed export
  cannot be a usable report.

**Style and shared-string ids are inputs, not outputs.** A cell's style id is fixed the moment
the cell is written, long before the styles part exists, and that XML is already in the package
by then. `WorkbookStylesPartWriter.GenerateStreamingContent` therefore emits one `cellXf` per
interned style in intern order, with no dedup and no remap, making index *i* the *i*-th style by
construction. Shared strings are handed out densely for the same reason, so no `SstMap`
translation is needed — unlike `SharedStringTable`, which is refcounted and leaves gaps.

**Value-driven style rules are now shared.** A date, a duration and a number are all stored as a
serial number, so only the number format distinguishes them on the way back in; a leading
apostrophe and a line break are likewise carried by the style. `XLWorksheet` applied these rules
on cell assignment. They moved to `XLValueStyleRules` so the cell setter and the streaming writer
share one definition — found because a streamed `DateTime` initially round-tripped as a `double`.

### Phase 3 — findings

**`CompressionLevel`: implemented on both paths.** The SDK does expose
`OpenXmlPackage.CompressionOption` (undocumented in the spec's survey), so
`SaveOptions.CompressionLevel` now works for ordinary saves, mapping onto
`System.IO.Packaging.CompressionOption`. It applies to parts a save creates; re-saving a loaded
file leaves its existing parts at whatever level they were written with. The streaming writer
takes `System.IO.Compression.CompressionLevel` directly. `Fastest` is 1.7× quicker than
`Optimal` on the 50K workload for a ~25% larger file.

**Async: not implemented, and deliberately so.** Split out rather than guessed at.

- For `XLWorkbook.SaveAs`, it is blocked. `System.IO.Packaging` is entirely synchronous and the
  SDK has no async save or package API — the only `*Async` methods in 3.4.1 are
  `OpenXmlWriter.WriteStartElementAsync` and friends, which do not help since the package write
  underneath them still blocks. A `SaveAsAsync` here could only be `Task.Run(() => SaveAs(...))`,
  which does not free a thread, it just moves the blocked one — an anti-pattern in a library API.
- For the streaming writer it *is* achievable, since we own the zip: `XmlWriterSettings.Async`
  plus `ZipArchive` entry `WriteAsync`. But it would need an async twin of every write method
  (`AppendRowAsync`, `Cell`…), roughly doubling a permanent public surface, for a workload that
  is CPU-bound on deflate rather than I/O-bound. The benefit is confined to not holding a thread
  while flushing to a slow sink — and the non-seekable-stream support already covers the main
  scenario that motivates it (writing to an HTTP response). Worth a follow-on spec with a real
  use case behind it, not speculative API surface now.
