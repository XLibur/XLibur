# Spec 02 — Load-Path Allocation Elimination (sheetData, SST, styles)

**Area:** Performance (read time + memory)
**Effort:** M (1–2 weeks, three independent sub-tasks)
**Dependencies:** None. Builds on PR #171 (raw `XmlReader` sheetData pass).
**Status:** ✅ Implemented in PR #175 — Tasks A, C1, C2 landed; Task B deferred (see Results)

## Results

`XLiburReadBenchmarks.LoadWorkbook`, 250K×15, net8.0, `InProcessEmitToolchain`, Ryzen 9 5950X:

| Stage | Mean | Allocated |
|---|---:|---:|
| Baseline | 4.750 s | 1020.92 MB |
| + Task A (SST raw reader) + C1 (dense style cache) | 3.897 s | 775.78 MB |
| + Task C2 (chunked attribute/value reads) | **3.968 s** | **392.88 MB** |

**−16.5% wall time, −61.5% allocations**, against acceptance targets of ≥15% and ≥35%. The
allocation target was beaten comfortably; the time target was met but only just. The final time is
within noise of the intermediate reading (Error ±0.03 s) — Task C2 buys allocation, not CPU. The
"≤ 3.2 s" aspiration in the Summary was **not** reached: after these three tasks, load time is no
longer dominated by garbage, so further gains need a different lever (see Follow-ups).

### What shipped

- **Task A** — `SharedStringReader` streams `<si>` entries with a raw `XmlReader`. Rich/phonetic
  entries still need the DOM, and richness is only knowable after the first child is consumed (a
  leading `<t>` may be followed by `<rPh>`), so their subtree is re-serialized rather than read via
  a plain `ReadOuterXml`.
- **Task C1** — `StyleValueCache`: flat array indexed by `cellXfs` index, scoped to the *workbook*
  rather than per-worksheet, since `styles.xml` is workbook-global and resolution is a pure
  function of the index.
- **Task C2** — `XmlReader.ReadValueChunk` into a reusable 64-char buffer for cell *and row*
  attributes and `<v>` content, parsing from the span, with a `StringBuilder` fallback for content
  wider than the buffer. `String`, `Error` and `Date` values still materialize a string, either
  because the text *is* the value or because the type is too rare to justify a span parser.

### Task B was deferred, deliberately

The remap targets the `Dictionary<Text,int>` probe in `SharedStringTable.IncreaseRef`: ~750K probes
at ~25–30 ns on the benchmark sheet, i.e. **under 1% of load time and zero allocations**. That does
not justify the risk to shared-string reference counting, which drives `DecreaseRef`, free-list
reuse, and what gets written back to the SST on save. Revisit only if a string-dominated workload
profiles as probe-bound.

### Findings for other specs

1. **`ReadValueChunk` works on attribute nodes**, not just text, and removed 99.8% of the
   allocation for that work (103.8 → 0.2 bytes/cell) when measured in isolation. This was verified
   with a throwaway probe before being built on. The technique generalizes to any `XmlReader` hot
   path in the IO layer.
2. **Doubles are saved with `"G15"`** by deliberate policy in `ObjectExtensions.ToInvariantString`,
   so values wider than 15 significant digits already lose precision on save today. **Spec 03's
   span-based `TryFormat` rewrite must target `G15`, not `"R"`/shortest-round-trip**, or it will
   silently change output for every workbook. Spec 03 task 1 currently says "R"/G17 — treat this
   as the correction.

### Follow-ups this opened

- Formula cells still allocate: `<f>` content is retained (unavoidable) but the `f` element's
  attributes (`t`, `ref`, `si`, `aca`, …) still go through `reader.Value`. Same technique applies.
- Load remains single-threaded; with garbage no longer dominant, per-sheet parallelism is the
  next structural lever rather than further allocation trimming.

## Summary

PR #171 moved `<sheetData>` parsing off `OpenXmlPartReader` (load: 6504 ms / 2357 MB → 4218 ms / 1017 MB on the 250K×15 benchmark). Three large allocation sources remain, each independently fixable: (1) per-cell transient strings for attributes and `<v>` content, (2) the shared-string table still being read through the full OpenXML DOM, and (3) per-cell dictionary probes for shared-string re-interning and style resolution. Target: **load ≤ 3.2 s / ≤ 600 MB** on the 250K×15 benchmark.

## Current state (verify before starting — line numbers drift)

- `XLibur/Excel/IO/WorksheetSheetDataReader.cs` (~1232 lines):
  - `LoadCellXml` (~line 234): reads `reader.Value` (string alloc) for attributes `r`, `s`, `t`, `ph`, `cm`, `vm` — `r`/`s`/`cm`/`vm` are immediately parsed to ints and discarded.
  - `LoadCellContentXml` (~line 340): `reader.ReadElementContentAsString()` on `<v>` allocates a string per numeric cell, parsed to `double` and thrown away. ~3.75M transient strings on the benchmark sheet. **This is the single biggest remaining load allocation.**
  - `SetSharedStringCellValue` (~line 933): parses SST index from the `<v>` string, then `SetCellValueDuringLoad` → `_sst.IncreaseRef(text)` — a `Dictionary<string,int>` probe per cell to re-intern a string already unique in the file's SST.
  - `ResolveCachedStyleValue` (~line 543): `Dictionary<int, XLStyleValue>` probe per styled cell; `GetInheritedStyleFast` (~line 514) runs per cell. styleIndex is dense and small — an array fits.
- `XLibur/Excel/IO/SharedStringReader.cs`: builds the entire `SharedStringTablePart.SharedStringTable` OpenXML DOM, then iterates `Elements<SharedStringItem>()`. Pre-sized `SharedStringEntry[]` exists, but the DOM is still fully materialized and discarded.

## Design

### Task A — SST via raw `XmlReader` (biggest bang for string-heavy files)

Rewrite `SharedStringReader` with the same raw-`XmlReader` treatment as PR #171:
- Stream `<si>` items; plain `<t>` items → `string` direct; rich-text `<si>` (has `<r>` runs) → fall back to parsing runs into `XLImmutableRichText` (keep existing rare-path behavior; if complex, `ReadOuterXml` + DOM for rich items only is acceptable).
- Honor `uniqueCount` for pre-sizing (already done — keep).

### Task B — File-SST → workbook-SST id remap

Build an `int[] sstRemap` once after reading the file SST: `sstRemap[fileId] = workbookSstId`, populated via one `IncreaseRef` pass over unique strings (refcount 0 initially, incremented per cell reference by array math instead of dictionary probes). Then `SetSharedStringCellValue` becomes: parse index → `sstRemap[index]` → store id + bump refcount directly. Replaces ~3.75M dictionary probes with array lookups.
- Requires an internal `SharedStringTable` API to increase refcount by id (exists conceptually — check `IncreaseRef` overloads) and to bulk-reserve capacity.

### Task C — Dense style array + numeric `<v>` without string

1. Replace `Dictionary<int, XLStyleValue> StyleList` with `XLStyleValue?[]` indexed by styleIndex (size = cellXfs count, known up front).
2. Attributes: for integer-valued attributes use `reader.MoveToAttribute` + a small helper that parses via `XmlReader.ReadContentAsInt()` where the implementation avoids the intermediate string, or parse `reader.Value` once into a reused position. Measure — if `ReadContentAsInt` still allocates internally on this reader, keep strings for attributes (they're smaller than `<v>`) and document why.
3. `<v>` numeric content: try `reader.ReadElementContentAsDouble()` first and measure allocations (it may bypass the string on `XmlTextReaderImpl`). If it doesn't, the fallback design is a UTF-8 side-parse: read the raw part stream through a custom minimal scanner only for the `<v>` hot path. **Do not build a full custom XML parser in this task** — if `ReadElementContentAsDouble` doesn't pay off, record measurements and stop; the UTF-8 tokenizer becomes its own follow-up spec.

## Work plan

| # | Task | Size | Independent? |
|---|------|------|--------------|
| A | SST raw-XmlReader rewrite | M | Yes |
| B | SST id remap array (needs A's entry array, coordinate) | S | After A |
| C1 | Dense style array | S | Yes |
| C2 | `ReadElementContentAsDouble` / attribute-alloc measurements + fixes | M | Yes |
| D | Update `XLiburReadBenchmarks` results + `MemoryProfile` snapshots; record before/after in PR | S | Last |

## Measurement protocol (required for every PR)

1. `dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --filter '*XLiburReadBenchmarks*'`
2. dotMemory snapshots via `dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- profile` (writes `.dmw` to `C:\profiles\`).
3. Report time + allocated + peak in the PR description, per the format used in PR #171.

## Acceptance criteria

1. 250K×15 `LoadAndReadAllCells`: load phase allocations reduced ≥ 35% from current main; wall time reduced ≥ 15%. No regression in `CreateAndSave`.
2. All existing tests green, including rich-text SST round-trip tests, phonetics, and inline-string tests.
3. Behavior identical for: missing `uniqueCount`, out-of-range SST index (must keep current error/tolerance behavior — check tests first), duplicate strings in file SST (legal per spec).

## Risks

- File SSTs can legally contain duplicate strings; the remap must map two file ids to one workbook id without breaking refcounts.
- `XmlReader` typed-content methods differ across reader implementations — measure, don't assume.

## References

- PR #171 (`bcb180b4`) for the established pattern and PR-description format.
- Perf survey notes: SST DOM, per-cell probes, `<v>` string allocs identified as the top three remaining load costs.
