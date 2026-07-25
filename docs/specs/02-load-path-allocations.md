# Spec 02 — Load-Path Allocation Elimination (sheetData, SST, styles)

**Area:** Performance (read time + memory)
**Effort:** M (1–2 weeks, three independent sub-tasks)
**Dependencies:** None. Builds on PR #171 (raw `XmlReader` sheetData pass).
**Status:** Proposed

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
