# Architecture

This document explains the internal architecture of XLibur — how the library
is structured, the design patterns it uses, and how it reads and writes Excel
files. It is aimed at contributors and anyone who wants to understand what
happens under the hood.

---

## 1. Overview

XLibur is a .NET library for reading, manipulating, and writing Excel 2007+
(`.xlsx`, `.xlsm`) files. It wraps the OpenXML SDK with a fluent, object-oriented
API so that callers never need to touch raw XML or the Open Packaging Convention
directly.

| Property           | Value                                                       |
|--------------------|-------------------------------------------------------------|
| Target frameworks  | net8.0, net9.0, net10.0                                     |
| License            | MIT                                                         |
| Core dependency    | DocumentFormat.OpenXml 3.4.1                                |
| Parser             | ClosedXML.Parser 2.0.0                                      |
| Font handling       | SixLabors.Fonts 1.0.1                                       |
| Spatial indexing   | RBush.Signed 4.0.0                                          |
| Number formatting  | ExcelNumberFormat 1.1.0                                     |

---

## 2. Layer Architecture

```
┌─────────────────────────────────────────────────────┐
│                    User Code                        │
│          var wb = new XLWorkbook("file.xlsx");       │
│          wb.Worksheets.Add("Sheet1");                │
└───────────────────────┬─────────────────────────────┘
                        │
┌───────────────────────▼─────────────────────────────┐
│              Public API  (IXL* interfaces)           │
│  IXLWorkbook, IXLWorksheet, IXLCell, IXLRange, ...   │
└───────────────────────┬─────────────────────────────┘
                        │
┌───────────────────────▼─────────────────────────────┐
│            Object Model  (XL* classes)               │
│  XLWorkbook, XLWorksheet, XLCell, XLRange,           │
│  SharedStringTable, XLStyleValue, XLCalcEngine, ...  │
└───────────────────────┬─────────────────────────────┘
                        │
┌───────────────────────▼─────────────────────────────┐
│               IO Layer  (Readers / Writers)          │
│  WorksheetSheetDataReader, SheetDataWriter,          │
│  WorkbookPartWriter, ChartReader, ChartWriter, ...   │
└───────────────────────┬─────────────────────────────┘
                        │
┌───────────────────────▼─────────────────────────────┐
│             DocumentFormat.OpenXml SDK               │
│  SpreadsheetDocument, OpenXmlReader/Writer, ...      │
└───────────────────────┬─────────────────────────────┘
                        │
┌───────────────────────▼─────────────────────────────┐
│                 .xlsx ZIP package                    │
│  xl/worksheets/sheet1.xml, xl/sharedStrings.xml, ... │
└─────────────────────────────────────────────────────┘
```

**Separation of concerns:**

- **Public API** — Stable interface types (`IXL*`) that callers program against.
  Changes here are breaking changes.
- **Object Model** — Internal `XL*` classes that hold in-memory state. They know
  nothing about XML or packaging.
- **IO Layer** — Static reader/writer classes in `XLibur/Excel/IO/` that translate
  between the object model and the OpenXML DOM or `XmlWriter` streams.
- **OpenXML SDK** — Handles the ZIP package, XML parsing, and relationship
  management.

---

## 3. Object Model

### 3.1 Workbook Hierarchy

```
XLWorkbook  (partial class: XLWorkbook.cs, _Load.cs, _Save.cs, _ImageHandling.cs)
├── XLWorksheets  (IXLWorksheets)
│   └── XLWorksheet  (partial: XLWorksheet.cs, XLWorksheetInternals.cs)
│       ├── XLCellsCollection  (4 parallel slices)
│       ├── XLRangeFactory
│       ├── XLColumns / XLRows
│       ├── XLTables
│       ├── XLCharts
│       ├── XLPictures
│       ├── XLConditionalFormats
│       ├── XLDataValidations
│       ├── XLAutoFilter
│       ├── XLDefinedNames  (worksheet-scoped)
│       └── XLPivotTables
├── XLDefinedNames  (workbook-scoped)
├── XLPivotCaches
├── SharedStringTable
├── XLCalcEngine
├── XLTheme
└── XLCustomProperties
```

`XLWorkbook` is a `partial class` split across four files:

| File                               | Responsibility                                |
|------------------------------------|-----------------------------------------------|
| `XLWorkbook.cs`                    | Properties, constructors, static defaults     |
| `XLWorkbook_Load.cs`              | `LoadSpreadsheetDocument` and load helpers     |
| `XLWorkbook_Save.cs`              | `CreatePackage` / `CreateParts` and save helpers |
| `XLWorkbook_ImageHandling.cs`     | In-cell image support                         |

**Base class hierarchy for stylized elements:**

```
XLStylizedBase  (abstract — holds StyleValue, propagates style changes)
├── XLRangeBase  (abstract — range address, cell enumeration, sorting)
│   ├── XLStoredRangeBase  (stores XLRangeAddress directly)
│   │   ├── XLRange
│   │   ├── XLRangeRow / XLRangeColumn
│   │   └── XLTable
│   ├── XLRow  (computes address from row number — avoids 48-byte overhead)
│   └── XLColumn  (computes address from column number)
└── XLCell  (lightweight facade — delegates to slices)
```

### 3.2 Cell Storage: Slice Architecture

Cell data is **not** stored inside `XLCell`. Instead, `XLCellsCollection` holds
four parallel sparse slices, each covering the full worksheet grid
(1,048,576 rows × 16,384 columns):

```
XLCellsCollection
├── ValueSlice       — cell values (numbers, booleans, text SST IDs, errors, dates)
├── FormulaSlice     — formula ASTs, types (Normal/Array/DataTable), dirty flags
├── StyleSlice       — Slice<XLStyleValue?>  (null = inherit from row/column/sheet)
└── MiscSlice        — Slice<XLMiscSliceContent>  (comments, hyperlinks, misc flags)
```

`XLCell` is a **lightweight facade** (two fields: `_cellsCollection` +
`_point`) that reads from and writes to the slices. It is created on demand
and never cached — the slices are the source of truth.

**Slice internals — `Slice<T>` + `Lut<T>`:**

`Slice<T>` is a sparse array indexed by `(row, column)`. Internally it uses a
two-level lookup table (`Lut<T>`):

```
Lut<RowData>  (top level — one entry per row)
└── LutBucket  (bottom level — 32-element array, bit-masked occupancy)
    └── RowData  (another Lut<T> for columns within that row)
```

- Top-level array grows by doubling; bottom buckets hold up to 32 elements.
- 5-bit split: `topIdx = index >> 5`, `bottomIdx = index & 0x1F`.
- Unused slots return `default(T)` without allocating.
- Column usage tracked in a `Dictionary<int, int>` for fast `UsedColumns`.

The `ISlice` interface provides uniform shift/insert/delete operations so that
inserting or deleting rows and columns is consistent across all four slices.

### 3.3 Addressing

| Type              | Description                                               | Packing                                   |
|-------------------|-----------------------------------------------------------|-------------------------------------------|
| `XLSheetPoint`    | A point in a worksheet (row + column). Used everywhere internally. | `ulong`: bits 0-13 = column (0-based), bits 14-33 = row (0-based) |
| `XLAddress`       | Cell address with optional worksheet, `$` flags, and cached string. | `ulong`: bits 0-14 = column+1, bits 15-35 = row+1, bit 36 = fixedRow, bit 37 = fixedColumn |
| `XLSheetRange`    | A rectangular range within a sheet (two `XLSheetPoint`s). | Two packed `ulong` values |
| `XLRangeAddress`  | A range address that can span worksheets, with validity flag. | Holds worksheet reference + two `XLAddress` |
| `XLBookArea`      | Sheet name + `XLSheetRange`. For cross-sheet references.   | `string Name` + `XLSheetRange Area` |
| `XLBookPoint`     | Sheet name + `XLSheetPoint`. For cross-sheet cell references. | `string Name` + `XLSheetPoint Point` |

Bit-packing avoids alignment padding and enables fast equality checks and
row-major sorting via a single `ulong` comparison.

### 3.4 Shared String Table

`SharedStringTable` (`XLibur/Excel/Cells/SharedStringTable.cs`) provides
workbook-level string deduplication. Each unique text (plain `string` or
`XLImmutableRichText`) is assigned an integer ID.

```
SharedStringTable
├── _table: List<Entry>        — ID → (Text, RefCount)
├── _reverseDict: Dictionary   — Text → ID  (for dedup lookup)
└── _freeIds: List<int>        — recycled IDs from deleted strings
```

- **Reference counting:** Each cell that uses a shared string increments
  the ref count. When a cell value changes, the old string is dereferenced.
  When `RefCount` reaches 0, the ID is added to `_freeIds`.
- **Free-list recycling:** New strings preferentially reuse IDs from
  `_freeIds` before appending to `_table`, keeping the table compact.
- **Bulk load optimization:** `EnsureCapacity(n)` pre-sizes internal
  collections; `TrimExcess()` releases slack after load completes.

---

## 4. Styles System

### 4.1 Immutable Values + Repository Deduplication

Styles follow a **Key → Value** pattern. Each style component is split into:

- **Key struct** (e.g., `XLStyleKey`, `XLFontKey`) — a value type holding the
  raw data. Keys are compared by value.
- **Value class** (e.g., `XLStyleValue`, `XLFontValue`) — an immutable
  reference type constructed from the key. Holds resolved sub-component
  references.

Deduplication happens via `XLRepositoryBase<TKey, TValue>`, a
`ConcurrentDictionary<TKey, WeakReference>` that ensures only one `TValue`
instance exists per unique `TKey`. When the value is no longer referenced, the
`WeakReference` allows GC to reclaim it.

```
XLStyleValue
├── XLAlignmentValue    (from XLAlignmentKey)
├── XLBorderValue       (from XLBorderKey)
├── XLFillValue         (from XLFillKey)
├── XLFontValue         (from XLFontKey)
├── XLNumberFormatValue (from XLNumberFormatKey)
├── XLProtectionValue   (from XLProtectionKey)
└── IncludeQuotePrefix  (bool)
```

`XLStyleValue.FromKey(ref key)` is the single entry point: it checks the
repository and returns the existing instance or creates a new one.

### 4.2 Style Inheritance

Cells can inherit styles from their row, column, or worksheet:

```
Cell  →  Row  →  Column  →  Worksheet
```

`XLStylizedBase` is the abstract base class. Setting a style propagates to
`Children` (e.g., setting a row style applies to all explicitly created cells
in that row). `XLCell` stores its style in the `StyleSlice` — a `null` entry
means "inherit".

The `IXLStylized` interface and `XLDeferredStyle` / `XLDeferredFont` / etc.
classes provide lazy wrappers that resolve the inherited chain on read.

### 4.3 Transition Cache

`XLStyleValue` has an 8-slot direct-mapped transition cache for bulk style
operations. When applying the same style change to thousands of cells (e.g.,
"make bold"), the first cell computes the new `XLStyleKey`, looks it up in the
repository, and caches the `(hash → result)` mapping. Subsequent cells with the
same base style hit the cache directly.

```csharp
private const int TransitionCacheSize = 8;
private const int TransitionCacheMask = TransitionCacheSize - 1;

// Slot = transitionHash & 0x7
// Collisions simply evict; benign races cause misses, never incorrect results.
```

---

## 5. Load Pipeline

Loading starts from the `XLWorkbook` constructor (file path or stream) and
flows through `LoadSpreadsheetDocument`:

```
XLWorkbook(path/stream)
  └─ Load(path/stream)
       └─ LoadSheets(path/stream)
            └─ SpreadsheetDocument.Open(...)
                 └─ LoadSpreadsheetDocument(dSpreadsheet)
```

**Steps in `LoadSpreadsheetDocument`:**

| #  | Step                         | Description                                             |
|----|------------------------------|---------------------------------------------------------|
| 1  | Load Shared Strings          | `SharedStringReader.Read(...)` → `SharedStringEntry[]`; pre-size workbook SST |
| 2  | Load Theme                   | `LoadWorkbookTheme(...)` → `XLTheme` colors             |
| 3  | Load Rich Data               | `RichDataReader.LoadRichData(...)` for in-cell images   |
| 4  | Load Properties              | Custom/extended file properties, file sharing, protection |
| 5  | Load Calculation Settings    | Calculate mode, reference style, precision settings     |
| 6  | Load Styles                  | Stylesheet → `StylesheetData` (number formats, fills, borders, fonts, dxf) |
| 7  | Detect Normal Style          | If the workbook's "Normal" built-in style is customized, update default column width |
| 8  | **Sheets Pass 1**            | Create empty `XLWorksheet` objects (name, position, visibility) |
| 9  | **Sheets Pass 2**            | Load all sheet data (cells, merges, drawings, etc.)     |
| 10 | Load Active Tab              | Set the active worksheet                                |
| 11 | Load Defined Names           | `DefinedNameReader.LoadDefinedNames(...)` — workbook and worksheet scoped |
| 12 | Load Pivot Cache Definitions | `PivotTableCacheDefinitionPartReader.Load(...)`         |
| 13 | Load Pivot Tables            | Associate pivot tables with their worksheets            |

**Why two passes for sheets?**

Pass 1 creates all `XLWorksheet` objects before any data is loaded. This avoids
costly calculation invalidation during loading — formulas that reference other
sheets always find their target sheet already exists. It also enables future
optimizations like parallel sheet loading.

---

## 6. Save Pipeline

Saving flows from `SaveAs` / `Save` through `CreatePackage` → `CreateParts`:

```
SaveAs(path) / Save()
  └─ CreatePackage(path/stream, docType, options)
       └─ SpreadsheetDocument.Create/Open(...)
            └─ CreateParts(document, options)
```

**Steps in `CreateParts`:**

| #  | Step                         | Description                                            |
|----|------------------------------|--------------------------------------------------------|
| 1  | Delete removed worksheets    | Clean up sheets the user deleted (including pivot caches) |
| 2  | Initialize RelId generator   | Collect existing relationship IDs to avoid collisions   |
| 3  | Generate workbook-level parts | Extended properties, workbook XML, shared string table, styles |
| 4  | Prepare pivot caches         | Ensure cache definitions and records are up-to-date     |
| 5  | Ensure dynamic array metadata | Add `XLDAPR` metadata if any cell uses dynamic arrays   |
| 6  | Ensure rich value image parts | For in-cell images                                      |
| 7  | **Per-worksheet loop:**      |                                                        |
|    | — Comments & VML             | Generate comment parts and VML drawings                 |
|    | — Tables                     | Generate table parts                                    |
|    | — Worksheet content          | `WorksheetPartWriter.GenerateWorksheetPartContent(...)`  |
|    | — Pivot tables               | Generate pivot table definition parts                   |
| 8  | Supplementary parts          | Calculation chain, custom properties, etc.              |
| 9  | Validate (optional)          | `OpenXmlValidator` if `options.ValidatePackage` is set  |

**Streaming `<sheetData>` with raw `XmlWriter`:**

`SheetDataWriter.StreamSheetData(...)` is the performance-critical hot path. It
bypasses the OpenXML SDK's `OpenXmlPartWriter` by extracting the underlying
`XmlWriter` via reflection and writing `<row>` / `<c>` elements directly:

```csharp
// Steal the XmlWriter from OpenXmlPartWriter for raw streaming
var xml = (XmlWriter)XmlWriterFieldInfo.GetValue(writer)!;
xml.WriteStartElement("sheetData", Main2006SsNs);
```

This avoids the overhead of creating OpenXML DOM objects for every cell and
enables streaming writes with minimal allocations. Cell references are formatted
into a reusable `char[]` buffer.

---

## 7. Formula Engine

### 7.1 Parsing & AST

Formulas are parsed by `ClosedXML.Parser` (external NuGet) into an AST. The
`FormulaParser` class wraps it:

```
string formula  →  FormulaParser.GetAst(...)  →  Formula (AST root)
                                                    └── ValueNode tree
```

`ExpressionCache` caches parsed ASTs keyed by formula string, so identical
formulas (common in spreadsheets) are parsed only once.

### 7.2 Evaluation

`CalculationVisitor` walks the AST and produces results:

- **`ScalarValue`** — A discriminated union (`readonly struct`) with 5 variants:
  Blank, Logical (bool), Number (double), Text (string), Error (XLError).
- **`AnyValue`** — Extends `ScalarValue` with Array and Reference variants for
  multi-cell results.
- **`CalcContext`** — Evaluation context holding the workbook, current cell, and
  culture.

**Built-in functions:** 232 functions registered across 10 category files:

| Category      | Functions | Examples                                    |
|---------------|-----------|---------------------------------------------|
| MathTrig      | 74        | SUM, SUMIF, ROUND, ABS, MOD, RAND           |
| Statistical   | 34        | AVERAGE, COUNT, MAX, MIN, STDEV, PERCENTILE  |
| Text          | 31        | CONCATENATE, LEFT, MID, FIND, SUBSTITUTE     |
| DateAndTime   | 23        | DATE, TODAY, YEAR, MONTH, DATEDIF            |
| Lookup        | 19        | VLOOKUP, INDEX, MATCH, INDIRECT, OFFSET      |
| Information   | 17        | ISBLANK, ISERROR, ISNUMBER, TYPE, CELL       |
| Engineering   | 12        | DEC2HEX, HEX2DEC, CONVERT, BITAND           |
| Database      | 12        | DSUM, DAVERAGE, DCOUNT, DGET                |
| Logical       | 7         | IF, AND, OR, NOT, IFERROR, IFS              |
| Financial     | 3         | PMT, PV, FV                                 |

### 7.3 Dependency Tracking

`DependencyTree` tracks which cells depend on which ranges so that dirty
flags propagate correctly when values change:

```
DependencyTree
├── _dependencies: Dictionary<XLCellFormula, FormulaDependencies>
├── _sheetTrees:   Dictionary<string, SheetDependencyTree>
│   └── SheetDependencyTree  (uses RBush spatial index)
└── _visitor:      DependenciesVisitor  (extracts precedents from AST)
```

- **RBush spatial indexing** — Each formula's precedent ranges are inserted as
  rectangles into a per-sheet R-tree (`RBush`). When a cell value changes,
  a spatial query finds all formulas whose precedent ranges overlap the
  changed cell.
- **Dirty propagation** — Marking a cell dirty triggers a cascade: all
  dependent formulas are also marked dirty. `XLCalculationChain` orders
  evaluation to minimize recalculation.
- The tree is built lazily on first evaluation and rebuilt when the workbook
  structure changes (sheets added/renamed/deleted, tables resized, etc.).

---

## 8. Ranges & Navigation

### Range Hierarchy

```
XLRangeBase  (abstract — address, cell enumeration, styling, sorting)
├── XLStoredRangeBase  (stores XLRangeAddress as a field)
│   ├── XLRange          — general rectangular range
│   ├── XLRangeRow       — single-row range within a parent range
│   ├── XLRangeColumn    — single-column range within a parent range
│   └── XLTable          — structured table (extends range with headers/totals)
├── XLRow                — full worksheet row (address computed from row number)
└── XLColumn             — full worksheet column (address computed from col number)
```

### XLRangeFactory

Each worksheet owns an `XLRangeFactory` that creates range objects from
`XLRangeKey` descriptors. The factory dispatches by `XLRangeType`:

```csharp
public XLRangeBase Create(XLRangeKey key) => key.RangeType switch
{
    XLRangeType.Range       => CreateRange(key.RangeAddress),
    XLRangeType.Column      => CreateColumn(...),
    XLRangeType.Row         => CreateColumn(...),
    XLRangeType.RangeColumn => CreateRangeColumn(...),
    XLRangeType.RangeRow    => CreateRangeRow(...),
    XLRangeType.Table       => CreateTable(...),
    ...
};
```

### Named Ranges (Defined Names)

`XLDefinedNames` holds named ranges at both workbook and worksheet scope.
Each `XLDefinedName` stores a formula string that resolves to a range
reference (e.g., `Sheet1!$A$1:$C$10`). The `DefinedNameReader` loads them
from the workbook XML during pass 11 of the load pipeline.

---

## 9. Sheet Features

### Tables

`XLTable` extends `XLStoredRangeBase` with structured reference support.
Key capabilities:
- Auto-filter integration (`IXLAutoFilter`)
- Header row and totals row
- `XLTableField` for column-level operations (totals functions, structured references)
- `TableNameGenerator` / `TableNameValidator` for name management
- Dedicated `TablePartWriter` for serialization

### Conditional Formatting

`XLConditionalFormat` (`IXLConditionalFormat`) supports:
- Color scales (2- and 3-color with min/mid/max)
- Data bars
- Icon sets
- Cell value rules, formula-based rules, etc.
- `ConditionalFormattingWriter` / `ConditionalFormatReader` for IO

### Data Validation

`XLDataValidation` (`IXLDataValidation`) provides:
- Whole number, decimal, date, time, text length criteria
- List validation
- Custom formula validation
- `XLValidationCriteria` base with typed subclasses (`XLWholeNumberCriteria`, `XLDecimalCriteria`, etc.)

### AutoFilters

`XLAutoFilter` (`IXLAutoFilter`) implements column-level filtering:
- Regular value filters (`XLFilterColumn`)
- Custom filters with operators (`XLCustomFilteredColumn`)
- `XLFilter` / `XLFilterConnector` for composing filter conditions

---

## 10. Charts & Drawings

### Charts

`XLChart` extends `XLDrawing<IXLChart>` and supports 78 chart types
(defined in the `XLChartType` enum), spanning areas, bars, columns, lines,
pies, donuts, scatter, radar, surface, bubbles, and 3D variants including
cone, cylinder, and pyramid shapes. Five extended types (BoxWhisker, Funnel,
Sunburst, Treemap, Waterfall) use the Office 2016+ `cx` namespace.

```
XLDrawing<T>  (base — name, description, position, visibility, shape ID)
└── XLChart   (chart type, title, series, bar orientation/grouping)
    ├── Series:          XLChartSeriesCollection  (primary)
    ├── SecondarySeries:  XLChartSeriesCollection  (secondary axis)
    └── SecondPosition:   XLDrawingPosition  (two-cell anchor)
```

- `ChartReader` / `ChartWriter` handle the chart XML parts
- `IsNew` flag distinguishes programmatically-created charts (which need
  full serialization) from loaded charts (which are passed through)

### Pictures

`XLPicture` (`IXLPicture`) handles image embedding:
- Supports multiple formats via `IXLGraphicEngine`
- Two-cell anchor positioning via `XLMarker` (TopLeft / BottomRight)
- Placement modes: `MoveAndSize`, `Move`, `Absolute`
- `PictureWriter` for serialization

---

## 11. Pivot Tables

The pivot table subsystem is the most complex area of the codebase, spanning
**65 `.cs` files** across three subdirectories:

```
PivotTables/
├── Areas/              — field/axis model (13 files)
│   ├── XLPivotTableAxis, XLPivotTableAxisField
│   ├── XLPivotDataField, XLPivotDataFields
│   ├── XLPivotFieldBase, XLPivotFieldAxisItem, XLPivotFieldItem
│   └── XLPivotTableFilters, XLPivotTablePageField
├── PivotStyleFormats/  — style formatting (12 files)
│   ├── XLPivotStyleFormat, XLPivotStyleFormatBase
│   └── XLPivotTableStyleFormats, XLPivotValueStyleFormat
├── PivotValues/        — value calculations (6 files)
│   ├── XLPivotValueCombination, XLPivotValueFormat
│   └── IXLPivotValue, IXLPivotValues
└── (root)              — core types (34 files)
    ├── XLPivotTable, XLPivotTableField, XLPivotTableEnums
    ├── XLPivotCache, XLPivotCacheValues, XLPivotCacheSharedItems
    ├── XLPivotCaches
    └── XLPivotArea, XLPivotConditionalFormat, XLPivotFormat
```

### Cache / Definition Separation

- **`XLPivotCache`** — Holds the source data: field names, shared items, and
  cache records. A cache can be shared by multiple pivot tables. Tied to a
  source range (worksheet area, external workbook, or consolidation).
- **`XLPivotTable`** — The layout definition: which fields go on which axis
  (row, column, data, page/filter), formatting, and conditional formats.

### Dedicated IO

Pivot tables have their own reader/writer pairs:

| Reader                                  | Writer                                     |
|-----------------------------------------|--------------------------------------------|
| `PivotTableCacheDefinitionPartReader`   | `PivotTableCacheDefinitionPartWriter`      |
|                                         | `PivotTableCacheRecordsPartWriter`         |
| `PivotTableDefinitionPartReader`        | `PivotTableDefinitionPartWriter2`          |

---

## 12. Key Design Patterns

1. **Slice storage** — Cell data split into parallel sparse arrays instead of
   per-cell objects. Eliminates millions of small allocations for large sheets.

2. **Style deduplication** — Immutable `XLStyleValue` instances deduplicated
   via `WeakReference`-backed `ConcurrentDictionary`. Same pattern repeated for
   font, fill, border, alignment, number format, and protection components.

3. **Transition cache** — 8-slot direct-mapped cache on `XLStyleValue` for
   amortizing repeated style transitions during bulk operations.

4. **Streaming save** — `SheetDataWriter` bypasses the OpenXML DOM and writes
   raw XML via `XmlWriter` for the `<sheetData>` element (the largest part of a
   worksheet). Reusable `char[]` buffers avoid per-cell string allocations.

5. **Bit-packed structs** — `XLSheetPoint` and `XLAddress` pack row, column,
   and flags into a single `ulong` for fast equality, hashing, and comparison.

6. **Spatial indexing** — `DependencyTree` uses `RBush` R-trees for O(log n)
   lookup of formulas affected by a cell change, instead of scanning all formulas.

7. **Two-pass loading** — Sheets are created (pass 1) before any data is loaded
   (pass 2), avoiding expensive recalculation invalidation during load.

8. **Shared String Table** — Workbook-level text deduplication with reference
   counting and free-list recycling. Pre-sized during load for performance.

9. **Two-level LUT** — `Lut<T>` provides sparse storage with 32-element
   buckets and bit-masked occupancy tracking. Memory-efficient for large
   worksheets with scattered data.

10. **Facade pattern** — `XLCell` is a lightweight facade (two fields) that
    delegates all storage to the underlying slices. Created on demand, never
    cached as the source of truth.

---

## 13. Project Layout

```
XLibur/
└── Excel/
    ├── AutoFilters/           — IXLAutoFilter, XLAutoFilter, filter columns
    ├── Caching/               — XLRepositoryBase (WeakRef dedup)
    ├── CalcEngine/            — Formula engine, dependency tree, 232 functions
    │   ├── Exceptions/        — Calc-specific exception types
    │   ├── Functions/         — 10 category files (MathTrig, Lookup, etc.)
    │   └── Visitors/          — CalculationVisitor, DependenciesVisitor
    ├── Cells/                 — XLCell, Slice<T>, ValueSlice, SharedStringTable
    ├── Charts/                — XLChart, 78 chart types, series collections
    ├── Columns/               — XLColumn, XLColumns
    ├── Comments/              — XLComment, XLComments
    ├── ConditionalFormats/    — Color scales, data bars, icon sets
    ├── ContentManagers/       — Part relationship management
    ├── Coordinates/           — XLAddress, XLSheetPoint, XLBookArea, etc.
    ├── CustomProperties/      — Document custom properties
    ├── DataValidation/        — Validation criteria (whole, decimal, date, time, text)
    ├── DefinedNames/          — Named ranges (workbook + worksheet scope)
    ├── Drawings/              — XLDrawing, XLPicture, anchor positioning
    ├── Exceptions/            — Domain-specific exceptions
    ├── Hyperlinks/            — Cell/range hyperlinks
    ├── InsertData/            — Bulk data insertion helpers
    ├── IO/                    — 42 reader/writer files (the full IO layer)
    ├── Misc/                  — Helper types
    ├── PageSetup/             — Print settings, margins, headers/footers
    ├── Patterns/              — Pattern matching utilities
    ├── PivotTables/           — 65 files: cache, definition, areas, styles, values
    ├── Protection/            — Sheet/workbook protection
    ├── Ranges/                — XLRange, XLRangeBase, factory, sorting, indexing
    │   ├── Index/             — Range indexing helpers
    │   └── Sort/              — Sort comparers
    ├── RichText/              — XLImmutableRichText, rich text formatting
    ├── Rows/                  — XLRow, XLRows
    ├── Sparkline/             — Sparkline support
    ├── Style/                 — XLStyleValue, keys, deferred styles, colors
    │   └── Colors/            — XLColor, theme colors, indexed colors
    └── Tables/                — XLTable, structured references, table fields
```
