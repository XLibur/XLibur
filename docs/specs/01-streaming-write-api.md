# Spec 01 — Streaming (Forward-Only) Write API for Large Workbooks

**Area:** Feature + Architecture + Performance (memory)
**Effort:** L (2–4 weeks)
**Dependencies:** None (but coordinate with Spec 03 — both touch `SheetDataWriter`)
**Status:** Proposed

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
