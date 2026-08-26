# Spec 29 — One resolver per element the two write paths both emit

**Area:** Architecture · **Correctness (divergence)**
**Effort:** M (~4–5 days)
**Dependencies:** None hard. **Conflicts with spec 31** — both work inside `SheetViewWriter.cs` and
`ColumnWriter.cs`. See Conflicts.
**Status:** Proposed.

## Goal

Make the two write paths agree on the XML they both emit, by extracting the *decision* — which
attributes, at which defaults — into value-in/value-out resolvers that both consume. Deliver a
cross-path agreement test that fails the moment they diverge again.

The agreement test is the artefact this spec is really delivering. The resolvers are how the
agreement becomes structural rather than coincidental.

## Why this spec exists

XLibur has two write paths. They share exactly one seam and re-implement everything above it.

| Element | Ordinary (DOM) path | Streaming path |
|---|---|---|
| `<pane>` | `SheetViewWriter.cs:113-147` `SetupPane` | `XLStreamingWorksheet.cs:484-509` `WriteSheetViews` |
| `<cols>` / `<col>` | `ColumnWriter.cs:20-61` `WriteColumns` | `XLStreamingWorksheet.cs:519-556` `WriteColumns` |
| `<sheets>` | `WorkbookPartWriter.cs:31-40` + `:112-207` | `XLStreamingWorkbook.cs:248-258` |
| `styles.xml` | `WorkbookStylesPartWriter.cs:18-85` `GenerateContent` | `:100-180` `GenerateStreamingContent` |
| cell / row leaves | `CellXmlWriter.cs` | `CellXmlWriter.cs` — **shared** |

Four elements have two implementations that must agree by hand. **One of them already disagrees in
shipped code.**

### The pane state disagrees today

Ordinary path, `SheetViewWriter.cs:124`:

```csharp
        pane.State = PaneStateValues.FrozenSplit;
```

Unconditional. Every pane the DOM path writes carries `state="frozenSplit"`.

Streaming path, `XLStreamingWorksheet.cs:502`:

```csharp
            xml.WriteAttributeString("state", "frozen");
```

So `FreezeRows(2)` produces `state="frozenSplit"` from `XLWorkbook.SaveAs` and `state="frozen"` from
`XLStreamingWorkbook`. Same model state, two spellings.

### Which one is wrong

Excel writes `state="frozen"`. Scanning every `.xlsx`/`.xlsm` under `XLibur.Tests/Resource/`
(395 workbooks, 30 `<pane>` tags found in `xl/worksheets/*.xml`):

| Spelling | Pane tags | Distinct files |
|---|---|---:|
| `state="frozen"` | 27 | 10 |
| `state="frozenSplit"` | 3 | 2 |
| `state="split"` | 0 | 0 |
| no `state` attribute | 0 | 0 |

Every one of the 30 came from an input/loaded fixture; no `output.xlsx` expected-output fixture
contains a `<pane>` at all, so the corpus is evidence about Excel, not about XLibur.

The three `frozenSplit` tags are the documented split-then-freeze case: two live inside
`<customSheetView>` elements in `Resource/TryToLoad/LO/xlsm/tdf111974.xlsm`, whose *main*
`<sheetView>` on the same sheet uses `state="frozen"`. Representative Excel-authored tags:

```xml
<pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/>
<pane xSplit="2" ySplit="1" topLeftCell="C7" activePane="bottomRight" state="frozen"/>
<pane xSplit="2" topLeftCell="C1" activePane="topRight" state="frozen"/>
```

**The DOM path is the bug; the streaming path is right.** Excel also omits the unused axis rather
than writing `0` — the third tag above has no `ySplit` at all. Only a LibreOffice-authored file in
the corpus writes `xSplit="0"` explicitly.

### Nothing catches it

`grep -rn 'frozenSplit|PaneStateValues' XLibur.Tests` returns **zero hits**. Across the whole tree
the two spellings appear in exactly three places: `SheetViewWriter.cs:124`,
`XLStreamingWorksheet.cs:502`, and the reader.

The reader is why it hides. `WorksheetElementReader.cs:76-77`:

```csharp
        if (new[] { PaneStateValues.Frozen, PaneStateValues.FrozenSplit }.Contains(pane?.State?.Value ??
                PaneStateValues.Split))
```

Both spellings map back to the same model. So the one test that exercises a streamed freeze,
`StreamingWriteTests.cs:401-419`, round-trips through `XLWorkbook` and asserts only the model:

```csharp
        await Assert.That(view.SplitRow).IsEqualTo(2);
        await Assert.That(view.SplitColumn).IsEqualTo(1);
```

That passes on both spellings. A load-and-compare test structurally cannot see this class of defect;
only a test that reads the bytes can.

### What does *not* diverge

Worth stating, because the brief for this spec assumed more divergence than the code has:

- **`activePane` agrees.** `SheetViewWriter.GetActivePaneValue` (`:149-158`) and
  `XLStreamingWorksheet.ResolveActivePane` (`:511-517`) are written differently but produce the same
  value for the same freeze: rows-only → `bottomLeft`, columns-only → `topRight`, both →
  `bottomRight`. The DOM path additionally honours an active cell through
  `GetActivePaneForActiveCell` (`:162-177`); the streaming API has no active-cell concept, so that is
  a capability gap, not a disagreement.
- **`topLeftCell` agrees.** DOM builds `GetColumnLetterFromNumber(SplitColumn + 1) + (SplitRow + 1)`
  (`:136-137`); streaming builds `new Point(_freezeRows + 1, _freezeColumns + 1).ToString()` (`:500`).
  `Point` is `(row, column)` (`Point.cs:29`), so both yield `B3` for a 2-row / 1-column freeze.
- **`<selection>`** is emitted by the DOM path (`SetupSelections`, `:188-223`) and never by streaming.
  Capability gap again — streaming exposes no selection API.

**One premise this spec has not confirmed:** the DOM path assigns `pane.HorizontalSplit = hSplit` and
`pane.VerticalSplit = ySplit` unconditionally (`:128-129`), including zero, so it should write
`xSplit="0"` where the streaming path omits the attribute entirely (`:494-498`). That is read off the
source, not off a produced file. **Task 1 confirms or disproves it**; if it is disproved, task 3's
resolver drops the split-omission rule and says so.

### `<cols>` is the partial precedent

The streaming path already calls into the DOM path for one decision. `XLStreamingWorksheet.cs:536`:

```csharp
                xml.WriteAttribute("width", ColumnWriter.GetColumnWidth(column.Width.Value).SaveRound());
```

`ColumnWriter.GetColumnWidth` (`ColumnWriter.cs:185-188`) is the shared width rule. That is exactly
the shape this spec generalises — one decision, two emitters — applied to one attribute out of six.
The other five (`customWidth`, `style`, `hidden`, `outlineLevel`, `collapsed`) are decided twice.

### styles.xml: mostly shared already

`WorkbookStylesPartWriter.cs` is 1,211 lines. `GenerateContent` is 68 of them (`:18-85`);
`GenerateStreamingContent` is 81 (`:100-180`). Between them they share five private helpers —
`ResolveNumberFormats`, `ResolveFonts`, `ResolveFills`, `ResolveBorders`, `BuildCellFormat` — so the
per-font, per-fill, per-border and per-`cellXf` emission is already single-implementation. What the
two do differently is **id assignment**, and that difference is the reason the streaming path exists:
its remarks at `:90-99` record that ids are the *input* rather than the output, because a cell was
handed a style id before the styles part existed. That is a real, documented divergence in behaviour,
not duplicated logic. Task 6 assesses whether anything is left to fold in.

### A leak worth naming

`XLStreamingWorkbook.cs:225`:

```csharp
        WorkbookStylesPartWriter.GenerateStreamingContent(stylesheet, Styles.OrderedStyles, new SaveContext());
```

`SaveContext` (`XLWorkbook_Save.NestedTypes.cs:14`) is a ten-member bag: differential formats, colour
filter dxf ids, a rel-id generator, table ids and names, pivot cache ids, an SST map, three style
dictionaries. `GenerateStreamingContent` reads three of them. The caller constructs one inline,
never binds it, and discards it — everything the call writes into it is thrown away.

That is evidence the shared code is taking the wrong parameter, not that the streaming path is
missing state. The same applies to `ColumnWriter.WriteColumns`, which takes a `SaveContext`
(`:25`) and reads `context.SharedStyles` and nothing else.

### The working precedent

`CellXmlWriter.cs` (234 lines) is the one seam the two paths genuinely share, with three callers:
`SheetDataWriter`, `XLStreamingWorksheet`, and `StreamingSharedStringTable`. Spec 01's Results record
why it took that shape rather than a wide one:

> **Phase 1 — the `IXLSheetDataSource` seam was not built.** … An `IEnumerable<XLRowData>` seam would
> add an interface dispatch per cell and a row object per row to the existing save path … Instead the
> *leaf* serializers (row start, cell start, value writers, type mapping) were extracted into
> `CellXmlWriter` and are shared by both producers. Same reuse, no change to the hot loop.

This spec applies that same rule one level up: share the leaf *decision*, not the traversal.

## Non-goals

- **Not merging the two write paths.** Two adapters over one resolved value is the target; one write
  path is not. Both serializers are kept deliberately. The DOM path exists so unmodelled markup
  survives a round trip — it reopens the loaded package and rewrites only the parts XLibur models
  (`docs/round-trip-fidelity.md`). The streaming path exists for bounded memory and owns its own
  `ZipArchive` in Create mode, because `System.IO.Packaging` opens read/write and holds every part's
  uncompressed bytes until close (spec 01 Results). Neither can become the other.
- **Not a performance spec.** But streaming's bounded-memory guarantee must not regress — task 7
  confirms it against spec 01's recorded 107.9 MB / 14.0 MB at 1M × 10.
- **Not touching `CellXmlWriter`.** It already is the shape this spec generalises, and it is close to
  spec 03's territory.
- **Not touching `SheetDataWriter` internals.** Spec 01 / spec 03 own those.
- **Not touching `WorksheetElementReader.LoadSheetViewPane`.** It must keep accepting both spellings —
  files in the wild carry both, as the corpus above shows.
- **No public API change.**

## Current state

Verified against the tree at `1b41cadd` (2026-08-24).

- `SheetViewWriter.cs` — 280 lines. `WriteSheetViews` `:62-98`; `SetupPane` `:113-147`, with
  `pane.State = PaneStateValues.FrozenSplit` at `:124`; `GetActivePaneValue` `:149-158`;
  `GetActivePaneForActiveCell` `:162-177`; `SetupSelections` `:188-223`
- `XLStreamingWorksheet.cs` — 575 lines. `FreezePanes` `:111-119`; `WriteSheetViews` `:484-509`, with
  `state="frozen"` at `:502`; `ResolveActivePane` `:511-517`; `WriteColumns` `:519-556`, reusing
  `ColumnWriter.GetColumnWidth` at `:536`
- `ColumnWriter.cs` — 282 lines. `WriteColumns` `:20-61` (takes `SaveContext` at `:25`);
  `BuildColumnElement` `:127-167`; `GetColumnWidth` `:185-188`
- `WorkbookStylesPartWriter.cs` — 1,211 lines. `GenerateContent` `:18-85`;
  `GenerateStreamingContent` `:100-180`; `ResolveFonts` `:1042-1077` (the one shared helper still
  taking a whole `SaveContext`)
- `WorkbookPartWriter.cs` — 424 lines. `GenerateContent` sheets block `:31-40`;
  `UpdateExistingSheets` `:112`; `AppendNewSheets` `:126`; `ReorderSheets` `:157`
- `XLStreamingWorkbook.cs` — 366 lines. `WriteStylesPart` `:222-231` with `new SaveContext()` at
  `:225`; `WriteWorkbookPart` `:233-261` with the `<sheets>` block at `:248-258`
- `SaveContext` — `XLWorkbook_Save.NestedTypes.cs:14-…`, ten members
- `WorksheetElementReader.LoadSheetViewPane` — `:73-89`, accepts both spellings at `:76-77`
- `StreamingWriteTests.cs` — 749 lines, 28 tests. The freeze test is `:401-419`
- `RoundTripFidelityTests.ReadPart` — `:159-169`, the raw-part-reading helper task 1 copies

## File structure

```
XLibur/Excel/IO/XLPaneSettings.cs              new — pane resolver + XLPaneState/XLPaneCorner
XLibur/Excel/IO/XLColumnSettings.cs            new — column resolver
XLibur/Excel/IO/SheetViewWriter.cs             modified — SetupPane consumes the resolver
XLibur/Excel/IO/ColumnWriter.cs                modified — BuildColumnElement consumes the resolver;
                                                          SaveContext parameter narrowed
XLibur/Excel/IO/WorkbookStylesPartWriter.cs    modified — ResolveFonts narrowed; streaming overload
                                                          drops its SaveContext parameter
XLibur/Excel/Streaming/XLStreamingWorksheet.cs modified — both writers consume the resolvers
XLibur/Excel/Streaming/XLStreamingWorkbook.cs  modified — the fabricated SaveContext deleted
XLibur.Tests/Excel/IO/WritePathAgreementTests.cs  new — the cross-path harness
```

Nothing is deleted. Both serializers stay.

## The design

A resolver is a `readonly struct` plus a static factory. It takes the model's values and returns the
attribute values, with every default already applied. It knows nothing about XML, the OpenXML SDK, or
either writer.

```csharp
/// <summary>Which corner of a split view holds the active pane.</summary>
internal enum XLPaneCorner { TopLeft, TopRight, BottomLeft, BottomRight }

/// <summary>
/// <c>ST_PaneState</c>. XLibur only ever writes <see cref="Frozen"/>; the other two exist because
/// files in the wild carry them and the reader accepts them.
/// </summary>
internal enum XLPaneState { Frozen, FrozenSplit, Split }

/// <summary>
/// The <c>&lt;pane&gt;</c> attributes both write paths emit, with defaults applied.
/// </summary>
/// <remarks>
/// This type owns the decision; the two writers own only the emission. Before spec 29 the decision
/// was made twice and the two copies disagreed on <c>state</c> — the DOM path wrote
/// <c>frozenSplit</c> for every pane while the streaming path wrote <c>frozen</c>, and the reader
/// mapped both back to the same model so no round-trip test could see it.
/// </remarks>
internal readonly struct XLPaneSettings
{
    /// <summary><c>xSplit</c>, or <c>null</c> when the axis is not split and the attribute is omitted.</summary>
    internal required int? SplitColumn { get; init; }

    /// <summary><c>ySplit</c>, or <c>null</c> when the axis is not split.</summary>
    internal required int? SplitRow { get; init; }

    internal required string TopLeftCell { get; init; }
    internal required XLPaneCorner ActivePane { get; init; }
    internal required XLPaneState State { get; init; }

    /// <summary><c>false</c> when no <c>&lt;pane&gt;</c> element should be written at all.</summary>
    internal bool HasPane => SplitColumn is not null || SplitRow is not null;

    /// <param name="paneTopLeftCell">
    /// An explicit pane scroll position, or <c>null</c> to anchor at split + 1. The streaming path
    /// always passes <c>null</c>; it exposes no equivalent API.
    /// </param>
    /// <param name="activeCell">
    /// The active cell, or <c>null</c>. When set it decides which corner owns the active pane;
    /// otherwise the split shape does. Streaming always passes <c>null</c>.
    /// </param>
    internal static XLPaneSettings Resolve(
        int splitColumn, int splitRow, XLAddress? paneTopLeftCell, XLAddress? activeCell);
}
```

`XLPaneState` and `XLPaneCorner` are XLibur's own enums rather than the SDK's `PaneStateValues` /
`PaneValues`, so the streaming path does not take an SDK dependency to read a decision it renders as
a raw string. Each adapter maps at the point of emission:

```csharp
// SheetViewWriter
pane.State = settings.State switch
{
    XLPaneState.Frozen => PaneStateValues.Frozen,
    XLPaneState.FrozenSplit => PaneStateValues.FrozenSplit,
    _ => PaneStateValues.Split,
};

// XLStreamingWorksheet
xml.WriteAttributeString("state", settings.State switch
{
    XLPaneState.Frozen => "frozen",
    XLPaneState.FrozenSplit => "frozenSplit",
    _ => "split",
});
```

The column resolver is the same shape:

```csharp
/// <summary>The <c>&lt;col&gt;</c> attributes both write paths emit, with defaults applied.</summary>
internal readonly struct XLColumnSettings
{
    internal required uint Min { get; init; }
    internal required uint Max { get; init; }
    internal required uint? StyleId { get; init; }

    /// <summary>Already through <see cref="ColumnWriter.GetColumnWidth"/> and <c>SaveRound</c>.</summary>
    internal required double? Width { get; init; }

    internal required bool Hidden { get; init; }
    internal required bool Collapsed { get; init; }

    /// <summary>0 means the attribute is omitted.</summary>
    internal required byte OutlineLevel { get; init; }

    /// <summary><c>customWidth</c> accompanies a width and is omitted without one.</summary>
    internal bool CustomWidth => Width is not null;

    internal static XLColumnSettings Resolve(
        uint min, uint max, uint? styleId, double? rawWidth,
        bool hidden, bool collapsed, int outlineLevel);
}
```

**What the resolvers do not own:** *which* columns get written. The DOM path expands every column in
`[min, max]`, back-fills 1..min-1 and max+1..16384 with the worksheet style, and collapses equal
neighbours into runs (`ColumnWriter.cs:80-183`). The streaming path writes one `<col>` per registered
range and does no filling or collapsing. Those are different products for different purposes and stay
where they are. Only the per-`<col>` attribute decision is shared.

## Global constraints

- Warnings are errors (`TreatWarningsAsErrors=true`); nullable enabled; new code must be
  null-annotated.
- Branch per spec; never commit to main. Commit prefixes `test:` for tasks 1 and 7, `fix:` for task 2,
  `refactor:` for tasks 3–6.
- No compound shell commands (`&&`, `||`, `;`) in agent tool calls.
- **Do not use `sed -i` on tracked files.** `.gitattributes` checks out CRLF and Git Bash's `sed -i`
  rewrites the file as LF, turning a one-line change into a whole-file diff. Use the Edit/Write tools
  and verify with `git diff --numstat`: a file whose changed-line count is near its total line count
  has been rewritten, not edited.
- Test filtering uses `--treenode-filter`, never `--filter`. Exit 5 = invalid option; exit 8 = zero
  tests matched. Never filter at solution level — name the `.csproj`.
- Pass `-f net10.0` for iteration; run without it before opening the PR.
- Build: `dotnet build XLibur/XLibur.csproj -c Release -v q`
- Tests: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
- Tests use TUnit: `await Assert.That(actual).IsEqualTo(expected)`. Assertions are awaitable and a
  missing `await` passes silently, so treat CS4014 as an error. `[Test]`, `[Arguments(...)]`,
  `[MethodDataSource(...)]`. The suite is serial (`[assembly: NotInParallel]`).
- `required` members need C# 11+, available on net8.0/net9.0/net10.0. If the language version is
  pinned lower, drop `required` and validate in `Resolve` instead.

## Work plan

| # | Task | Size | Gate |
|---|---|---|---|
| 1 | Cross-path agreement harness | M | New test file compiles and runs; the pane-state case **FAILS** |
| 2 | Fix the pane state towards `frozen` | S | Task 1's pane case green; full suite green |
| 3 | `XLPaneSettings`; both paths onto it | M | Harness green; full suite green |
| 4 | `XLColumnSettings`; both paths onto it | M | Harness green; full suite green |
| 5 | Narrow the fabricated `SaveContext` | S | `grep 'new SaveContext()' Streaming/` returns nothing |
| 6 | Assess `<sheets>` and styles.xml | S | A recorded decision either way |
| 7 | Confirm streaming's bounded memory | S | Within spec 01's 107.9 MB / 14.0 MB |

Task 1 is the important one. Tasks 3 and 4 are behaviour-preserving once task 2 has landed, and task 1
is the only thing that can prove that.

---

### Task 1 — The cross-path agreement harness

**Landed:** `7c57efd8` — failing on the pane state, as intended.

Write the same workbook down both paths, read the bytes out of both packages, and compare the named
attributes. **This test lands failing**, on the pane state. That is the point: it is a defect report
that runs.

**Files:**
- Create: `XLibur.Tests/Excel/IO/WritePathAgreementTests.cs`

**Interfaces:**
- Produces: `Both_write_paths_agree_on_the_pane`, `Both_write_paths_agree_on_a_column` — the gate for
  tasks 2, 3 and 4.

- [x] **Step 1: Write the harness**

```csharp
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using XLibur.Excel;
using XLibur.Excel.Streaming;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// XLibur has two write paths — the ordinary DOM save and <see cref="XLStreamingWorkbook"/> — and
/// both emit <c>&lt;pane&gt;</c> and <c>&lt;col&gt;</c>. A load-and-compare test cannot see a
/// disagreement between them, because the reader normalises both spellings back to one model: that
/// is exactly how <c>state="frozenSplit"</c> vs <c>state="frozen"</c> shipped unnoticed. These tests
/// read the bytes.
/// </summary>
public class WritePathAgreementTests
{
    #region Pane

    [Test]
    [Arguments(1, 0)]
    [Arguments(0, 2)]
    [Arguments(2, 1)]
    public async Task Both_write_paths_agree_on_the_pane(int freezeRows, int freezeColumns)
    {
        using var dom = SaveViaDom(freezeRows, freezeColumns);
        using var streamed = SaveViaStreaming(freezeRows, freezeColumns);

        var domPane = PaneTag(dom);
        var streamedPane = PaneTag(streamed);

        await Assert.That(domPane).IsNotEmpty();
        await Assert.That(streamedPane).IsNotEmpty();

        foreach (var name in new[] { "state", "activePane", "topLeftCell", "xSplit", "ySplit" })
            await Assert.That(Attribute(streamedPane, name)).IsEqualTo(Attribute(domPane, name));
    }

    #endregion Pane

    #region Column

    [Test]
    public async Task Both_write_paths_agree_on_a_column()
    {
        using var dom = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            var column = ws.Column(2);
            column.Width = 33;
            column.OutlineLevel = 1;
            ws.Cell("A1").Value = "x";
            wb.SaveAs(dom);
        }

        using var streamed = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(streamed))
        {
            var sheet = wb.AddWorksheet("S");
            var column = sheet.Column(2);
            column.Width = 33;
            column.OutlineLevel = 1;
            sheet.AppendRow("x");
            wb.Finish();
        }

        var domCol = ColTag(dom, min: 2);
        var streamedCol = ColTag(streamed, min: 2);

        await Assert.That(domCol).IsNotEmpty();
        await Assert.That(streamedCol).IsNotEmpty();

        foreach (var name in new[] { "width", "customWidth", "hidden", "outlineLevel", "collapsed" })
            await Assert.That(Attribute(streamedCol, name)).IsEqualTo(Attribute(domCol, name));
    }

    #endregion Column

    #region Helpers

    private static MemoryStream SaveViaDom(int freezeRows, int freezeColumns)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("S");
            ws.SheetView.Freeze(freezeRows, freezeColumns);
            ws.Cell("A1").Value = "x";
            wb.SaveAs(ms);
        }

        return ms;
    }

    private static MemoryStream SaveViaStreaming(int freezeRows, int freezeColumns)
    {
        var ms = new MemoryStream();
        using (var wb = XLStreamingWorkbook.Create(ms))
        {
            var sheet = wb.AddWorksheet("S");
            sheet.FreezePanes(freezeRows, freezeColumns);
            sheet.AppendRow("x");
            wb.Finish();
        }

        return ms;
    }

    private static string PaneTag(MemoryStream package)
        => Match(ReadSheet1(package), "<pane\\b[^>]*>");

    private static string ColTag(MemoryStream package, uint min)
        => Match(ReadSheet1(package), $"<col\\b[^>]*\\bmin=\"{min}\"[^>]*>");

    private static string Match(string xml, string pattern)
    {
        var match = Regex.Match(xml, pattern);
        return match.Success ? match.Value : string.Empty;
    }

    /// <summary>The attribute's value, or <c>null</c> when the attribute is absent.</summary>
    private static string? Attribute(string tag, string name)
    {
        var match = Regex.Match(tag, $"\\b{name}=\"([^\"]*)\"");
        return match.Success ? match.Groups[1].Value : null;
    }

    private static string ReadSheet1(MemoryStream package)
    {
        package.Position = 0;
        using var archive = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true);
        var entry = archive.Entries.First(e =>
            e.FullName.Equals("xl/worksheets/sheet1.xml", StringComparison.OrdinalIgnoreCase));

        using var entryStream = entry.Open();
        using var reader = new StreamReader(entryStream);
        return reader.ReadToEnd();
    }

    #endregion Helpers
}
```

`ReadSheet1` is `RoundTripFidelityTests.ReadPart` (`:159-169`) narrowed to one part. **Confirm both
paths name the first sheet's part `xl/worksheets/sheet1.xml`** — the streaming path builds it through
`XLStreamingWorksheet.EntryName(Index)`, and its relationship target is written as
`EntryName(worksheet.Index)["xl/".Length..]` (`XLStreamingWorkbook.cs:272-273`). If the names differ,
resolve the part through the relationship rather than hardcoding it, and say so in the test.

If the DOM path's `<col min="2">` is collapsed into a wider run so no tag has `min="2"`, widen
`ColTag` to match the run containing column 2 and record what it matched — do not weaken the
assertions.

- [x] **Step 2: Run it and record exactly what fails**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WritePathAgreementTests/*"`

Expected: **FAIL** on `state` for all three pane arguments — `frozenSplit` from the DOM path against
`frozen` from streaming.

Record every other attribute that fails, verbatim. Two in particular:

- **`xSplit` / `ySplit`.** The premise in "Why this spec exists" is that the DOM path writes `xSplit="0"`
  for a rows-only freeze while streaming omits the attribute. This task is what decides whether that
  premise is true. If both omit it, say so and drop the split-omission rule from task 3's resolver —
  a disproved premise is a result, not a problem.
- **Column attributes.** If `Both_write_paths_agree_on_a_column` passes on the first run, that is also
  a result: it means `<col>` is already in agreement for this case, and task 4 becomes a refactor with
  no defect behind it. Say so and keep task 4 anyway — the point is that agreement stops being
  coincidental.

- [x] **Step 3: Land the test failing, marked**

Do not fix anything here. Add `[Skip("Fails on the pane state; spec 29 task 2 fixes it.")]` to
`Both_write_paths_agree_on_the_pane` **only if** the repo's CI will not accept a red test on the
branch; otherwise leave it red and let task 2 turn it green in the same PR chain. Check with:

Run: `git log --oneline -20`

- [x] **Step 4: Commit**

```bash
git add XLibur.Tests/Excel/IO/WritePathAgreementTests.cs
git commit -m 'test(io): compare the two write paths byte for byte (spec 29 task 1)'
```

---

### Task 2 — Fix the pane state

**Landed:** `bce00355`; the four fixtures it turned red regenerated in `4f9f92cf`.

Excel writes `state="frozen"` for a pure freeze. The streaming path is right; the DOM path is wrong.

**Files:**
- Modify: `XLibur/Excel/IO/SheetViewWriter.cs:124`

- [x] **Step 1: Change the spelling**

```csharp
        // Excel writes state="frozen" for a pure freeze; frozenSplit is what it writes when a pane
        // was frozen from an existing manual split, which XLibur never produces. Of the 30 <pane>
        // tags in the test corpus, 27 (10 files) are "frozen" and 3 (2 files) are "frozenSplit",
        // and the frozenSplit ones sit in <customSheetView> alongside a "frozen" main sheetView.
        // The reader accepts both (WorksheetElementReader.LoadSheetViewPane) and must keep doing so.
        pane.State = PaneStateValues.Frozen;
```

- [x] **Step 2: Run the agreement test**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WritePathAgreementTests/*"`
Expected: the `state` assertion passes for all three arguments. `xSplit`/`ySplit` may still fail —
task 3 owns that.

- [x] **Step 3: Run the whole suite**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS.

`grep -rn 'frozenSplit|PaneStateValues' XLibur.Tests` returned nothing before this change, so no test
pins the old spelling. If one fails anyway, it is asserting the state indirectly — through a file
comparison or an `output.xlsx` fixture. **Do not revert.** A reference `output.xlsx` that contains
`frozenSplit` is an XLibur-authored artefact of this defect, not evidence about Excel; regenerate it
and note which fixture changed.

- [x] **Step 4: Confirm the change is one line**

Run: `git diff --numstat XLibur/Excel/IO/SheetViewWriter.cs`
Expected: a small changed-line count against a 280-line file. A count near 280 means the file was
rewritten with LF endings — discard and redo with the Edit tool.

- [x] **Step 5: Commit**

```bash
git add XLibur/Excel/IO/SheetViewWriter.cs
git commit -m 'fix(io): write state="frozen" for a frozen pane, matching Excel (spec 29 task 2)'
```

---

### Task 3 — `XLPaneSettings`

**Landed:** `ba6a206b`. The `xSplit="0"` premise held, so the split-omission rule stayed.
`activeCell` is `Point?`, not `XLAddress?` — see Results.

**Files:**
- Create: `XLibur/Excel/IO/XLPaneSettings.cs`
- Modify: `XLibur/Excel/IO/SheetViewWriter.cs:113-177`
- Modify: `XLibur/Excel/Streaming/XLStreamingWorksheet.cs:484-517`

**Interfaces:**
- Produces: `XLPaneSettings`, `XLPaneState`, `XLPaneCorner`,
  `XLPaneSettings.Resolve(int, int, XLAddress?, XLAddress?) → XLPaneSettings`.
- Removes: `SheetViewWriter.GetActivePaneValue`, `SheetViewWriter.GetActivePaneForActiveCell`,
  `XLStreamingWorksheet.ResolveActivePane`.

- [x] **Step 1: Write the resolver**

```csharp
    internal static XLPaneSettings Resolve(
        int splitColumn, int splitRow, XLAddress? paneTopLeftCell, XLAddress? activeCell)
    {
        var corner = ResolveCorner(splitColumn, splitRow, activeCell);

        return new XLPaneSettings
        {
            // Excel omits the unused axis rather than writing 0.
            SplitColumn = splitColumn > 0 ? splitColumn : null,
            SplitRow = splitRow > 0 ? splitRow : null,
            TopLeftCell = paneTopLeftCell is { IsValid: true } p
                ? p.ToStringRelative(false)
                : XLHelper.GetColumnLetterFromNumber(splitColumn + 1) + (splitRow + 1),
            ActivePane = corner,
            // XLibur never produces a split-then-frozen pane, so the state is always Frozen.
            State = XLPaneState.Frozen,
        };
    }

    private static XLPaneCorner ResolveCorner(int splitColumn, int splitRow, XLAddress? activeCell)
    {
        if (activeCell is not { } active)
        {
            if (splitRow == 0 && splitColumn == 0) return XLPaneCorner.TopLeft;
            if (splitRow == 0) return XLPaneCorner.TopRight;
            return splitColumn == 0 ? XLPaneCorner.BottomLeft : XLPaneCorner.BottomRight;
        }

        var bottom = splitRow > 0 && active.RowNumber > splitRow;
        var right = splitColumn > 0 && active.ColumnNumber > splitColumn;

        if (bottom && right) return XLPaneCorner.BottomRight;
        if (bottom) return XLPaneCorner.BottomLeft;
        return right ? XLPaneCorner.TopRight : XLPaneCorner.TopLeft;
    }
```

The two branches are `GetActivePaneValue` (`SheetViewWriter.cs:149-158`) and
`GetActivePaneForActiveCell` (`:162-177`) moved verbatim. **Check the active-cell property names**:
`GetActivePaneForActiveCell` reads `active.Row` and `active.Column` off whatever
`xlWorksheet.ActiveCell` is; use the same members here rather than the ones written above if they
differ.

**If task 1 disproved the `xSplit="0"` premise**, drop the `> 0 ? … : null` on the two split
properties, make them plain `int`, and record the correction in this task's checkbox.

- [x] **Step 2: Put the DOM path onto it**

```csharp
    private static Pane? SetupPane(SheetView sheetView, XLSheetViewContentManager svcm, XLWorksheet xlWorksheet)
    {
        var settings = XLPaneSettings.Resolve(
            xlWorksheet.SheetView.SplitColumn,
            xlWorksheet.SheetView.SplitRow,
            xlWorksheet.SheetView.PaneTopLeftCellAddress,
            xlWorksheet.ActiveCell);

        if (!settings.HasPane)
        {
            sheetView.RemoveAllChildren<Pane>();
            svcm.SetElement(XLSheetViewContents.Pane, null);
            return null;
        }

        var pane = sheetView.Elements<Pane>().FirstOrDefault();
        if (pane == null)
        {
            pane = new Pane();
            sheetView.InsertAt(pane, 0);
        }

        svcm.SetElement(XLSheetViewContents.Pane, pane);

        pane.HorizontalSplit = settings.SplitColumn;
        pane.VerticalSplit = settings.SplitRow;
        pane.TopLeftCell = settings.TopLeftCell;
        pane.ActivePane = ToOpenXml(settings.ActivePane);
        pane.State = ToOpenXml(settings.State);

        return pane;
    }
```

Note the reordering: the original creates a `Pane`, fills it, and *then* removes it again when both
splits are zero (`:139-144`). Asking the resolver first makes that dead work go away. Behaviour is
unchanged — the removal branch already discarded everything it had written.

`PaneTopLeftCellAddress` is an `XLAddress` today, not an `XLAddress?`. Pass
`addr is { IsValid: true } ? addr : null` at the call site rather than changing its type; that keeps
this spec's footprint to the two writers.

- [x] **Step 3: Put the streaming path onto it**

```csharp
    private void WriteSheetViews(XmlWriter xml)
    {
        xml.WriteStartElement("sheetViews", Main2006SsNs);
        xml.WriteStartElement("sheetView", Main2006SsNs);
        xml.WriteAttribute("workbookViewId", 0u);

        // The streaming API exposes neither a pane scroll position nor an active cell, so both
        // resolver inputs are null here. Everything else is decided in exactly one place.
        var settings = XLPaneSettings.Resolve(_freezeColumns, _freezeRows, null, null);

        if (settings.HasPane)
        {
            xml.WriteStartElement("pane", Main2006SsNs);

            if (settings.SplitColumn is { } xSplit)
                xml.WriteAttribute("xSplit", xSplit);

            if (settings.SplitRow is { } ySplit)
                xml.WriteAttribute("ySplit", ySplit);

            xml.WriteAttributeString("topLeftCell", settings.TopLeftCell);
            xml.WriteAttributeString("activePane", ToAttribute(settings.ActivePane));
            xml.WriteAttributeString("state", ToAttribute(settings.State));

            xml.WriteEndElement(); // pane
        }

        xml.WriteEndElement(); // sheetView
        xml.WriteEndElement(); // sheetViews
    }
```

Delete `ResolveActivePane` (`:511-517`). `XLHelper.GetColumnLetterFromNumber` now produces the
`topLeftCell` on both sides, replacing the `Point` round trip — verify with task 1 that the string is
unchanged.

- [x] **Step 4: Confirm the decision lives in one place**

Run: `grep -rn 'PaneStateValues\|"frozenSplit"\|"frozen"\|bottomRight\|bottomLeft\|topRight' XLibur --include=*.cs`

Expected: `XLPaneSettings.cs` (the enums and the resolver), `SheetViewWriter.cs` and
`XLStreamingWorksheet.cs` (one mapping switch each), and `WorksheetElementReader.cs:76-77` (the
reader, unchanged and deliberately still accepting both). Nothing else.

- [x] **Step 5: Build and run**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS, including the full `Both_write_paths_agree_on_the_pane` matrix.

- [x] **Step 6: Commit**

```bash
git add XLibur/Excel/IO/XLPaneSettings.cs XLibur/Excel/IO/SheetViewWriter.cs XLibur/Excel/Streaming/XLStreamingWorksheet.cs
git commit -m 'refactor(io): resolve the pane once, emit it twice (spec 29 task 3)'
```

---

### Task 4 — `XLColumnSettings`

**Landed:** `67193e8e`. The column test passed cold, so this is a refactor with no defect behind
it. Double-rounding solved by passing the raw width — see Results.

**Files:**
- Create: `XLibur/Excel/IO/XLColumnSettings.cs`
- Modify: `XLibur/Excel/IO/ColumnWriter.cs:127-167`
- Modify: `XLibur/Excel/Streaming/XLStreamingWorksheet.cs:519-556`

**Interfaces:**
- Produces: `XLColumnSettings`,
  `XLColumnSettings.Resolve(uint, uint, uint?, double?, bool, bool, int) → XLColumnSettings`.

- [x] **Step 1: Write the resolver**

```csharp
    internal static XLColumnSettings Resolve(
        uint min, uint max, uint? styleId, double? rawWidth,
        bool hidden, bool collapsed, int outlineLevel)
        => new()
        {
            Min = min,
            Max = max,
            StyleId = styleId,
            Width = rawWidth is { } w ? ColumnWriter.GetColumnWidth(w).SaveRound() : null,
            Hidden = hidden,
            Collapsed = collapsed,
            OutlineLevel = outlineLevel > 0 ? (byte)Math.Min(outlineLevel, byte.MaxValue) : (byte)0,
        };
```

`ColumnWriter.GetColumnWidth` stays where it is. It was already the shared piece
(`XLStreamingWorksheet.cs:536`); the resolver now carries it plus the five decisions that were not
shared.

- [x] **Step 2: Put `BuildColumnElement` onto it**

`ColumnWriter.BuildColumnElement` (`:127-167`) computes `styleId`, `columnWidth`, `isHidden`,
`collapsed` and `outlineLevel` and then builds a `Column`. Split it: keep the lookup, hand the raw
values to `Resolve`, and build the `Column` from the result.

```csharp
        var settings = XLColumnSettings.Resolve(
            (uint)columnNumber, (uint)columnNumber, styleId, rawWidth,
            isHidden, collapsed, outlineLevel);

        var column = new Column
        {
            Min = settings.Min,
            Max = settings.Max,
            Style = settings.StyleId,
            Width = settings.Width,
            CustomWidth = settings.CustomWidth ? true : null,
        };

        if (settings.Hidden) column.Hidden = true;
        if (settings.Collapsed) column.Collapsed = true;
        if (settings.OutlineLevel > 0) column.OutlineLevel = settings.OutlineLevel;
```

**One behaviour to preserve exactly.** The DOM path always supplies a width, so `CustomWidth` is
always `true` there today (`:156`). The `else` arm at `:144-148` supplies
`ctx.WorksheetColumnWidth`, which is already through `GetColumnWidth().SaveRound()` at
`SheetViewWriter.cs:264` — so pass it to `Resolve` **pre-resolved is wrong**; either pass the raw
worksheet width or add a `Resolve` overload that takes an already-resolved width. Whichever you
choose, prove it with task 1's column test and with:

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*Column*/*"`

Double-rounding a width is the failure mode to watch for.

`WritePreColumns` (`:80-101`) and `WritePostColumns` (`:169-183`) also build `Column` elements
directly. Route them through the resolver too, or leave them and say why — they write the worksheet
default rather than a column's own settings.

- [x] **Step 3: Put the streaming writer onto it**

```csharp
        foreach (var column in ordered)
        {
            var settings = XLColumnSettings.Resolve(
                (uint)column.FirstColumn, (uint)column.LastColumn,
                column.Style is null ? null : _workbook.Styles.GetOrAdd(column.Style),
                column.Width, column.Hidden, column.Collapsed, column.OutlineLevel);

            xml.WriteStartElement("col", Main2006SsNs);
            xml.WriteAttribute("min", settings.Min);
            xml.WriteAttribute("max", settings.Max);

            if (settings.Width is { } width)
            {
                xml.WriteAttribute("width", width);
                xml.WriteAttributeString("customWidth", TrueValue);
            }

            if (settings.StyleId is { } styleId)
                xml.WriteAttribute("style", styleId);

            if (settings.Hidden)
                xml.WriteAttributeString("hidden", TrueValue);

            if (settings.OutlineLevel > 0)
                xml.WriteAttribute("outlineLevel", settings.OutlineLevel);

            if (settings.Collapsed)
                xml.WriteAttributeString("collapsed", TrueValue);

            xml.WriteEndElement(); // col
        }
```

Note that `GetOrAdd` is now called only when a style exists, where before it was inside the same
`if` — confirm the call is still made in the same order relative to the width, since `GetOrAdd`
mutates the style table.

- [x] **Step 4: Confirm the width rule has one caller each**

Run: `grep -rn 'GetColumnWidth' XLibur --include=*.cs`
Expected: its definition in `ColumnWriter.cs:185`, the call in `XLColumnSettings.Resolve`, and the
`SheetFormatProperties` default-width call at `SheetViewWriter.cs:264`. **Not**
`XLStreamingWorksheet.cs:536` — that call moves into the resolver.

- [x] **Step 5: Build and run**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS.

- [x] **Step 6: Commit**

```bash
git add XLibur/Excel/IO/XLColumnSettings.cs XLibur/Excel/IO/ColumnWriter.cs XLibur/Excel/Streaming/XLStreamingWorksheet.cs
git commit -m 'refactor(io): resolve a column once, emit it twice (spec 29 task 4)'
```

---

### Task 5 — Narrow the fabricated `SaveContext`

**Landed:** `8596b08f`. Both greps return nothing.

`XLStreamingWorkbook.cs:225` builds a ten-member `SaveContext` to satisfy a signature, fills three of
its dictionaries, and throws all of it away. Give the shared code a narrower parameter instead.

**Files:**
- Modify: `XLibur/Excel/IO/WorkbookStylesPartWriter.cs:100-180`, `:1042-1077`
- Modify: `XLibur/Excel/Streaming/XLStreamingWorkbook.cs:222-231`
- Modify: `XLibur/Excel/IO/ColumnWriter.cs:20-61`, `:103-167`

- [x] **Step 1: Narrow `ResolveFonts`**

It is the only shared helper still taking a whole `SaveContext`, and it reads and rewrites
`context.SharedFonts` and nothing else.

```csharp
    private static void ResolveFonts(Stylesheet stylesheet, Dictionary<XLFontValue, FontInfo> sharedFonts)
```

`ResolveNumberFormats`, `ResolveFills` and `ResolveBorders` already take explicit dictionaries.
Update `GenerateContent` to pass `context.SharedFonts`.

- [x] **Step 2: Drop the parameter from `GenerateStreamingContent`**

Its three dictionaries are write-only from the caller's point of view, so it can own them:

```csharp
    internal static void GenerateStreamingContent(Stylesheet stylesheet,
        IReadOnlyList<XLStyleValue> orderedStyles)
    {
        var sharedFonts = new Dictionary<XLFontValue, FontInfo>();
        var sharedNumberFormats = new Dictionary<XLNumberFormatValue, NumberFormatInfo>();
        var sharedStyles = new Dictionary<XLStyleValue, StyleInfo>();
        // ... body, with context.X replaced by the locals
    }
```

Keep the existing `<remarks>` at `:90-99` — it is the reason this overload exists and is worth more
than the code it documents. Add one line:

```csharp
    /// <para>
    /// Takes no <c>SaveContext</c>: it needs three of that type's ten members and the caller
    /// discards all of them, so the bag was a signature to satisfy rather than state to carry.
    /// </para>
```

- [x] **Step 3: Delete the fabricated instance**

```csharp
        WorkbookStylesPartWriter.GenerateStreamingContent(stylesheet, Styles.OrderedStyles);
```

- [x] **Step 4: Narrow `ColumnWriter.WriteColumns`**

It reads `context.SharedStyles` twice (`:27`, `:138`, `:146`) and never writes it.

```csharp
    internal static void WriteColumns(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        double worksheetColumnWidth,
        IReadOnlyDictionary<XLStyleValue, StyleInfo> sharedStyles)
```

Thread it through `WriteMainColumns` and `BuildColumnElement`. Fix the call site in the save path.
If `SharedStyles` is typed `Dictionary<,>` on `SaveContext`, passing it as `IReadOnlyDictionary<,>`
is implicit — no change needed there.

- [x] **Step 5: Gate it**

Run: `grep -rn 'new SaveContext()' XLibur/Excel/Streaming/`
Expected: no output.

Run: `grep -n 'SaveContext' XLibur/Excel/IO/WorkbookStylesPartWriter.cs XLibur/Excel/IO/ColumnWriter.cs`
Expected: `WorkbookStylesPartWriter.GenerateContent` and its `AddDifferentialFormats` chain only —
those genuinely use the differential-format and colour-filter members. `ColumnWriter.cs` returns
nothing.

- [x] **Step 6: Build and run**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS.

- [x] **Step 7: Commit**

```bash
git add XLibur/Excel/IO/WorkbookStylesPartWriter.cs XLibur/Excel/IO/ColumnWriter.cs XLibur/Excel/Streaming/XLStreamingWorkbook.cs
git commit -m 'refactor(io): give the shared style and column writers a narrower parameter (spec 29 task 5)'
```

---

### Task 6 — Assess `<sheets>` and styles.xml

**Decision: no resolver for either.** Recorded with line references and measured evidence in
Results below. The spec's "seven `GenerateContent` calls" is eight.

Two of the four duplicated elements are named in this spec but not yet folded in. Decide, and record
the decision either way. **A recorded "no" is the deliverable if that is what the code supports.**

**Files:**
- Modify: `docs/specs/29-write-path-resolvers.md` — a Results section

- [x] **Step 1: Assess `<sheets>`**

Read `WorkbookPartWriter.GenerateContent` (`:31-48`) and its three helpers: `UpdateExistingSheets`
(`:112`), `AppendNewSheets` (`:126`), `ReorderSheets` (`:157`). Compare against
`XLStreamingWorkbook.WriteWorkbookPart` (`:248-258`).

The claim to test: the DOM `<sheets>` writer is a **patcher over an existing element**, not a
generator. It removes deleted sheets, updates the ones already present, appends only the new ones,
and reorders around `xlWorkbook.UnsupportedSheets` so chartsheets and macro sheets keep their
position — the mechanism `docs/round-trip-fidelity.md` documents. The streaming path has no existing
element, no unsupported sheets and no visibility model, so the only decision the two share is
"name, sheetId, r:id". That is three field copies with no defaulting in either path.

If that holds, the answer is **no resolver** — an `XLSheetSettings` carrying three verbatim fields
would be indirection without a decision. Record it with the line references that show it.

If it does not hold — if the streaming path is silently defaulting something the DOM path decides
(`sheetId` is `(uint)worksheet.Index` in streaming versus `xlSheet.SheetId` in the DOM path, and
those are not obviously the same number) — that is a finding. Write a `Both_write_paths_agree_on_the_sheets_entry`
test into task 1's harness and treat it the way task 1 treated the pane.

- [x] **Step 2: Assess styles.xml**

`GenerateContent` (68 lines) and `GenerateStreamingContent` (81 lines) already share five private
helpers. What is left is id assignment, and the remarks at `:90-99` explain why it cannot be shared:
the streaming writer hands a style id to a cell before the styles part exists, so ids are its input,
not its output.

Confirm that by listing what each calls. `GenerateContent`: `ResolveDefaultFormatId`,
`CollectWorkbookStyles`, `CollectCustomNumberFormats`, `ResolveNumberFormats`, `ResolveFonts`,
`ResolveFills`, `ResolveBorders`, `BuildSharedStyleMappings`, `ResolveCellStyleFormats`,
`ResolveRest`, `RemapStyleIds`, `AddDifferentialFormats`. `GenerateStreamingContent`:
`ResolveNumberFormats`, `ResolveFonts`, `ResolveFills`, `ResolveBorders`, `BuildCellFormat`.

The seven `GenerateContent` calls with no streaming counterpart are all deduplication and remapping —
exactly what the streaming path cannot do. If that is what you find, record **no further sharing**
and say which seven.

- [x] **Step 3: Write the Results section**

Add a `## Results (<date>)` section to this file recording, for each element: the decision, the
evidence, and the line references. Include the numbers task 1 produced — which attributes agreed on
the first run and which did not.

- [x] **Step 4: Commit**

```bash
git add docs/specs/29-write-path-resolvers.md
git commit -m 'docs(specs): record the sheets and styles.xml assessment for spec 29'
```

---

### Task 7 — Confirm streaming's bounded memory is intact

**Measured:** 107.9 MB shared strings / 14.0 MB inline strings — identical to spec 01's baseline.

The resolvers add two struct constructions per sheet — one pane, one per `<col>` — on a path whose
whole purpose is bounded memory. Show it did not move.

- [x] **Step 1: Measure**

```
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -f net10.0 -- profile streaming
```

- [x] **Step 2: Compare against spec 01's recorded baseline**

| Writer | Peak heap | Elapsed |
|---|---|---|
| `XLStreamingWorkbook`, shared strings | 107.9 MB | 10.3 s |
| `XLStreamingWorkbook`, inline strings | 14.0 MB | 9.5 s |

Expected: within noise of both. `XLPaneSettings` is resolved once per sheet;
`XLColumnSettings` once per registered column range, which is bounded by the number of ranges the
caller declared, not by rows. Neither is in the per-cell loop.

**If peak heap has risen at all**, the likely cause is a resolver being constructed inside a row or
cell loop rather than in the part header. Check that `XLColumnSettings.Resolve` is called only from
`WriteColumns`, which runs once per sheet from `EnsureStarted` (`XLStreamingWorksheet.cs:479`).

Note that spec 01 records ~40% run-to-run timing variance on this machine, so treat elapsed as
indicative and peak heap — which it describes as stable across runs — as the real gate.

- [x] **Step 3: Run the whole suite on both frameworks**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Expected: PASS on net8.0 and net10.0.

- [x] **Step 4: Commit the numbers**

```bash
git add docs/specs/29-write-path-resolvers.md
git commit -m 'docs(specs): record the streaming memory numbers for spec 29'
```

---

## Acceptance criteria

1. `grep -n 'PaneStateValues.FrozenSplit' XLibur/Excel/IO/SheetViewWriter.cs` returns nothing.
2. `grep -rn 'new SaveContext()' XLibur/Excel/Streaming/` returns nothing.
3. `grep -n 'SaveContext' XLibur/Excel/IO/ColumnWriter.cs` returns nothing.
4. `grep -rn 'GetColumnWidth' XLibur --include=*.cs` returns exactly three sites: the definition in
   `ColumnWriter.cs`, the call in `XLColumnSettings.Resolve`, and the default-width call in
   `SheetViewWriter.WriteSheetFormatProperties`. `XLStreamingWorksheet.cs` is not among them.
5. `grep -rn 'bottomRight\|bottomLeft\|topRight' XLibur --include=*.cs` shows the corner names only in
   `XLPaneSettings.cs` and in one mapping switch per writer — not in a decision.
6. `WorksheetElementReader.cs:76-77` still accepts both `Frozen` and `FrozenSplit`, unchanged.
7. `WritePathAgreementTests` passes every argument of `Both_write_paths_agree_on_the_pane` and
   `Both_write_paths_agree_on_a_column`, with no assertion weakened and no attribute removed from
   either comparison list.
8. Both `SheetViewWriter.cs` and `XLStreamingWorksheet.cs` still exist and still emit their own XML.
   Neither calls the other's serializer.
9. Full suite green on net8.0 and net10.0.
10. Streaming peak managed heap at 1M × 10 within noise of spec 01's 107.9 MB (shared strings) and
    14.0 MB (inline strings).
11. No public API change.
12. Task 6's decision on `<sheets>` and styles.xml is recorded in a Results section with line
    references, whether the answer is "fold in" or "do not".
13. Any premise this spec states and task 1 disproves — the `xSplit="0"` claim in particular — is
    corrected in the Results section rather than quietly dropped.

## Conflicts

- **Spec 31** (worksheet element writers get one interface, being written in parallel today and not
  yet on disk in `docs/specs/`) touches `SheetViewWriter.cs` and `ColumnWriter.cs`. **That is a real
  conflict.** 29 rewrites `SetupPane` and `BuildColumnElement`; 31 rewrites the interface those
  methods sit behind. The two cannot run concurrently.

  **Recommended order: 29 first.** Two reasons, and the second is the stronger one.

  First, footprint. 29's whole change to those two files is: one enum value at
  `SheetViewWriter.cs:124`, `SetupPane` reduced to a resolver call plus five assignments, two private
  helpers deleted, and `BuildColumnElement` split at its existing seam. 31 is a structural sweep
  across every worksheet element writer. Rebasing 29 onto 31 means redoing a small change; rebasing
  31 onto 29 means redoing a large one.

  Second, and this is what decides it: **29 is a correctness fix with a live defect behind it**, and
  its task 1 gate is a test 31 also wants. If 31 lands first, 31's own refactor of `SetupPane` will
  carry `PaneStateValues.FrozenSplit` forward into whatever interface it defines, and the fix then
  has to be made inside new structure that was built around the wrong value. Landing 29 first hands
  31 a `SheetViewWriter` whose pane decision is already extracted into `XLPaneSettings` — one fewer
  thing for its interface to have to carry, and a cross-path test that will notice if the sweep
  breaks either emitter.

  If 31 must start first, 29 waits in full. Do not split it and land task 2 alone: the one-line fix
  without task 1's harness is exactly the situation that let the divergence ship.

- **Spec 01** (streaming write API) is **done** and created the second write path. Its Results section
  is required reading before starting: it records why `XLStreamingWorkbook` writes its own OPC
  package over `ZipArchive` in Create mode (`System.IO.Packaging` opens read/write, which is
  `ZipArchiveMode.Update`, and that holds every part's uncompressed bytes until close), and why the
  wide `IXLSheetDataSource` seam was declined in favour of sharing the leaf serializers in
  `CellXmlWriter`. Both facts constrain this spec: the streaming path cannot reuse the DOM emitters,
  and the sharing must stay at the leaf.

- **Spec 03** (save-path allocations, in progress) touches `SheetDataWriter.cs` and `XLWorksheet.cs`.
  **29 touches neither.** No collision, no sequencing needed. 29's non-goals deliberately exclude
  `SheetDataWriter` internals and `CellXmlWriter` for exactly this reason. The two can run
  concurrently.

- **Spec 24** (worksheet element load dispatch) modifies `WorksheetElementReader.cs`. 29 reads
  `LoadSheetViewPane` as evidence and **does not modify it** — acceptance criterion 6 pins that. The
  two are disjoint in practice; if 24 is mid-flight, 29's criterion 6 grep will need its new line
  numbers.

- **Spec 23** (one implementation per style interface) works on the style facades and `XL*Key.cs`.
  29 touches `WorkbookStylesPartWriter.cs` in task 5 only, and only to change two method signatures —
  it reads `XLStyleValue` but changes nothing about how styles are built. Low collision risk; if both
  are live, land 23's key changes first since they are the wider ones.

- **Spec 18 task 5** is on the load path (`LoadWorksheetElements`). No overlap with either write path.

---

## Results (2026-08-27)

**Status:** Done. Branch `refactor/29-write-path-resolvers`, seven commits off `upstream/main`
at `806d69f7`.

| # | Task | Commit |
|---|---|---|
| 1 | Cross-path agreement harness, landed failing | `7c57efd8` |
| 2 | Pane state fixed towards `frozen` | `bce00355` |
| 3 | `XLPaneSettings`; both paths onto it | `ba6a206b` |
| — | The four pane fixtures regenerated | `4f9f92cf` |
| 4 | `XLColumnSettings`; both paths onto it | `67193e8e` |
| 5 | Fabricated `SaveContext` narrowed | `8596b08f` |
| 6–7 | This section, and the memory numbers | see below |

Full suite green on net8.0 and net10.0: 28,358 tests, 0 failed, 10 skipped (5 per framework,
all pre-existing).

### Task 1 — what the first run actually read

The harness writes the same workbook down both paths, opens both packages and compares the
named attributes on the raw tag. Every reading below is off the bytes, not off the source.

One correction to the harness as the spec wrote it: **the two paths use different namespace
spellings**, and the spec's regexes matched neither DOM tag. The OpenXML SDK prefixes every
element (`<x:pane>`); the streaming writer writes a default namespace and no prefix
(`<pane>`). Without allowing an optional prefix, `PaneTag` returned `""` for the DOM package
and all four tests failed on `IsNotEmpty` before reaching a single attribute comparison. The
helpers now accept an optional prefix. That is a serialisation difference, not a disagreement
about content, and no assertion was weakened to absorb it.

`Both_write_paths_agree_on_the_pane`, attribute by attribute, before any fix:

| Freeze | Attribute | DOM | Streaming | Agreed |
|---|---|---|---|---|
| rows 1, cols 0 | `state` | `frozenSplit` | `frozen` | **no** |
| | `activePane` | `bottomLeft` | `bottomLeft` | yes |
| | `topLeftCell` | `A2` | `A2` | yes |
| | `xSplit` | `"0"` | *absent* | **no** |
| | `ySplit` | `1` | `1` | yes |
| rows 0, cols 2 | `state` | `frozenSplit` | `frozen` | **no** |
| | `activePane` | `topRight` | `topRight` | yes |
| | `topLeftCell` | `C1` | `C1` | yes |
| | `xSplit` | `2` | `2` | yes |
| | `ySplit` | `"0"` | *absent* | **no** |
| rows 2, cols 1 | `state` | `frozenSplit` | `frozen` | **no** |
| | `activePane` | `bottomRight` | `bottomRight` | yes |
| | `topLeftCell` | `B3` | `B3` | yes |
| | `xSplit` | `1` | `1` | yes |
| | `ySplit` | `2` | `2` | yes |

The tags verbatim:

```xml
<!-- rows 1, cols 0 -->
<x:pane xSplit="0" ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozenSplit" />
    <pane        ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" />
<!-- rows 0, cols 2 -->
<x:pane xSplit="2" ySplit="0" topLeftCell="C1" activePane="topRight" state="frozenSplit" />
    <pane xSplit="2"          topLeftCell="C1" activePane="topRight" state="frozen" />
<!-- rows 2, cols 1 -->
<x:pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozenSplit" />
    <pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen" />
```

**The `xSplit="0"` premise held.** The DOM path did write `xSplit="0"` for a rows-only freeze
and `ySplit="0"` for a columns-only freeze, where the streaming path omitted the attribute
entirely. Task 3's resolver therefore keeps the split-omission rule, and both paths now omit
the unused axis, which is what Excel writes.

`activePane` and `topLeftCell` agreed on all three arguments, confirming what the spec said
about them. In particular `XLHelper.GetColumnLetterFromNumber(splitColumn + 1) + (splitRow + 1)`
and the `Point` round trip produced the same string, so replacing the latter in task 3 changed
nothing.

**`Both_write_paths_agree_on_a_column` passed on the first run.** `<col>` was already in
agreement for that case, so task 4 is a refactor with no defect behind it — kept anyway, because
the point is that the agreement stops being coincidental. The tags:

```xml
<x:col min="2" max="2" width="33.710625" style="0" customWidth="1" outlineLevel="1" />
    <col min="2" max="2" width="33.710625"           customWidth="1" outlineLevel="1" />
```

Two observations that are *not* failures, because neither attribute is in the comparison list
and both are the "which columns get written" difference the spec puts out of scope:

- `style` differs — the DOM path writes `style="0"` explicitly, the streaming path omits it when
  the column carries no style of its own.
- The DOM path also emits `<col min="1" max="1">` back-filling the worksheet default; the
  streaming path writes only the one registered range.

### Task 2 — the fixtures that pinned the defect

`grep -rn 'frozenSplit|PaneStateValues' XLibur.Tests` returned nothing, as the spec said, but
four tests still went red. All four are XLibur-authored expected-output fixtures asserting the
state indirectly through a whole-file comparison — the case the spec anticipated. None was
reverted and no assertion was weakened.

Because task 3 additionally drops `xSplit="0"` from the same panes, the four were regenerated
**once, after task 3** (`4f9f92cf`) rather than twice. Each was proved with a full-part diff of
the entire package before replacement:

| Fixture | Sheet | Delta |
|---|---|---|
| `Resource/Examples/Misc/FreezePanes.xlsx` | 1 | `state` `frozenSplit`→`frozen` |
| | 2 | `state`, and `xSplit="0"` dropped |
| | 3 | `state`, and `ySplit="0"` dropped |
| | 4 | `state` |
| `Resource/Examples/Misc/SheetViews.xlsx` | 8 | `state`, and `xSplit="0"` dropped |
| `…/PivotTableReferenceFiles/PivotSubtotalsSource/output.xlsx` | 1 | `state`, and `xSplit="0"` dropped |
| `…/PivotTableReferenceFiles/TwoPivotTablesWithSingleSource/output.xlsx` | 1 | `state`, and `xSplit="0"` dropped |

No other part changed content in any of the four. The two `Examples` packages also carry a new
`_rels/.rels` id set and a new core-properties `.psmdcp` name; every save regenerates those and
`ExcelDocsComparer` ignores them. The two pivot packages, which are saved over a loaded package,
have byte-identical part inventories.

A known-red baseline of exactly those six test names was recorded before task 3 began and
checked at every gate. Nothing outside it ever went red.

### Task 6 — `<sheets>`: no resolver

**The claim holds.** The DOM `<sheets>` writer is a patcher over an existing element, not a
generator:

- `WorkbookPartWriter.GenerateContent:33-34` removes sheets the model deleted from the
  `workbook.Sheets` that came out of the loaded package.
- `UpdateExistingSheets:112-124` patches `Name` on entries already present and reads each
  sheet's `RelId` **back out of the file** into the model (`:121`).
- `AppendNewSheets:126-155` appends only entries whose `r:id` is not already there, and
  allocates a new one from `context.RelIdGenerator` only when the sheet did not come from a
  loaded file (`:132-141`).
- `ReorderSheets:157-` reorders around `xlWorkbook.UnsupportedSheets` (`:168`, `:171`, `:184`,
  `:204`, `:216`) so chartsheets and macro sheets keep their position — the mechanism
  `docs/round-trip-fidelity.md` documents.

`XLStreamingWorkbook.WriteWorkbookPart:248-258` has no existing element, no deletions, no
unsupported sheets and no visibility model. It writes three attributes per sheet from a list it
owns.

On the `sheetId` question the spec flagged as "not obviously the same number" — measured, not
reasoned. Three sheets added in order, then the same three with the middle one deleted:

```xml
<!-- DOM -->        <x:sheet name="One" sheetId="1" r:id="rId2" />
                    <x:sheet name="Two" sheetId="2" r:id="rId3" />
                    <x:sheet name="Three" sheetId="3" r:id="rId4" />
<!-- streaming -->    <sheet name="One" sheetId="1" r:id="rId1" />
                      <sheet name="Two" sheetId="2" r:id="rId2" />
                      <sheet name="Three" sheetId="3" r:id="rId3" />
<!-- DOM, "Two" deleted -->
                    <x:sheet name="One" sheetId="1" r:id="rId2" />
                    <x:sheet name="Three" sheetId="3" r:id="rId3" />
```

`name` and `sheetId` agree for every workbook both paths can express. They agree *by
coincidence of provenance*, not by a shared rule, and the deletion case shows it: DOM `SheetId`
is a workbook-lifetime identity counter (`XLWorksheets.cs:22`, `:264`, and `XLWorkbook_Load.cs:67`
on load) that leaves a gap when a sheet is removed, while streaming's is a part ordinal
(`XLStreamingWorksheet.Index`) that cannot gap because streaming has no delete API. `r:id`
differs numerically and legitimately — the DOM generator is shared with the theme, styles and
sharedStrings parts, streaming numbers worksheet relationships from 1 inside its own rels part,
and a relationship id is only meaningful relative to its own `.rels`.

So there is nothing being silently defaulted, and no `Both_write_paths_agree_on_the_sheets_entry`
test was written. **Decision: no resolver.** An `XLSheetSettings` carrying three verbatim fields
would be indirection without a decision, and it would assert a shared rule for two numbers that
only happen to coincide.

### Task 6 — styles.xml: no further sharing

**The claim holds**, with one arithmetic correction to the spec.

`GenerateContent` (`WorkbookStylesPartWriter.cs:18`) calls twelve helpers;
`GenerateStreamingContent` (`:104`) calls five. Four are shared —
`ResolveNumberFormats`, `ResolveFonts`, `ResolveFills`, `ResolveBorders` — and `BuildCellFormat`
is called directly by the streaming writer and reached from `GenerateContent` through
`ResolveRest`, so five helpers are shared in total, as the spec said.

The spec says "the seven `GenerateContent` calls with no streaming counterpart". **There are
eight, not seven:** `ResolveDefaultFormatId`, `CollectWorkbookStyles`, `CollectCustomNumberFormats`,
`BuildSharedStyleMappings`, `ResolveCellStyleFormats`, `ResolveRest`, `RemapStyleIds`,
`AddDifferentialFormats`. Twelve minus the four shared is eight; the spec's own two lists say so
and the count in the prose is off by one.

All eight are deduplication, remapping and differential-format work — exactly what the streaming
path cannot do, for the reason its own `<remarks>` at `:87-102` gives: it hands a style id to a
cell the moment that cell is written, long before the styles part exists, so ids are its input
rather than its output. **Decision: no further sharing.**

What task 5 *did* remove was the wrong parameter rather than duplicated logic: `ResolveFonts`
now takes `Dictionary<XLFontValue, FontInfo>` (`:1051`) instead of the whole ten-member bag, and
`GenerateStreamingContent` owns its three dictionaries as locals instead of filling a
`SaveContext` the caller threw away.

### Task 7 — streaming's bounded memory

`dotnet run -c Release --project XLibur.Benchmarks -f net10.0 -- profile streaming`, 1,000,000
rows × 10 columns:

| Writer | Peak heap | Spec 01 baseline | Elapsed | Spec 01 baseline |
|---|---|---|---|---|
| `XLStreamingWorkbook`, shared strings | **107.9 MB** | 107.9 MB | 10.19 s | 10.3 s |
| `XLStreamingWorkbook`, inline strings | **14.0 MB** | 14.0 MB | 8.64 s | 9.5 s |

Peak heap is identical to the baseline to the reported precision on both configurations. Elapsed
is within the ~40% run-to-run variance spec 01 records for this machine and is indicative only;
peak heap is the gate and it did not move.

That is what the design predicts: `XLPaneSettings` is resolved once per sheet and
`XLColumnSettings` once per registered column range, both from `WriteSheetViews` / `WriteColumns`
under `EnsureStarted` (`XLStreamingWorksheet.cs:478-479`), which runs once per sheet. Neither is
in the row or cell loop, and both are `readonly struct`s.

### Acceptance criteria

| # | Criterion | Result |
|---|---|---|
| 1 | no `PaneStateValues.FrozenSplit` in `SheetViewWriter.cs` | **see below** |
| 2 | no `new SaveContext()` under `Excel/Streaming/` | met — grep returns nothing |
| 3 | no `SaveContext` in `ColumnWriter.cs` | met — grep returns nothing |
| 4 | `GetColumnWidth` at exactly three sites | met — `ColumnWriter.cs:188` (definition), `XLColumnSettings.cs:66`, `SheetViewWriter.cs:250`. `XLStreamingWorksheet.cs` is not among them |
| 5 | corner names only in mapping switches | met — `XLStreamingWorksheet.cs:516-521` only; the DOM side maps to SDK enum members. The one other hit, `PivotTableDefinitionPartWriter2.cs:667`, is a pivot-area type and unrelated |
| 6 | reader still accepts both spellings | met, unchanged — now at `WorksheetElementReader.cs:192-193`, not `:76-77`; spec 24/28 moved it |
| 7 | both agreement tests pass, nothing weakened | met — 4/4 arguments green, both comparison lists intact |
| 8 | both serializers still exist and emit their own XML | met — neither calls the other's serializer |
| 9 | full suite green on net8.0 and net10.0 | met — 28,358 tests, 0 failed |
| 10 | streaming peak heap within noise of baseline | met — identical, not merely within noise |
| 11 | no public API change | met — `PublicAPI.Unshipped.txt` untouched |
| 12 | task 6 decision recorded with line references | met — above |
| 13 | disproved premises corrected here | met — the `xSplit="0"` premise was **confirmed**, not disproved; the namespace-prefix and helper-count corrections are recorded above |

**Criterion 1 is not met as literally written, and it cannot be.** It contradicts this spec's own
design section: the snippet at lines 296–302 mandates a mapping switch in `SheetViewWriter`
containing `XLPaneState.FrozenSplit => PaneStateValues.FrozenSplit`, which is the exact string
criterion 1 greps for. The implementation follows the design section. What criterion 1 is
protecting is met in substance — `SheetViewWriter` no longer *decides* the pane state; the token
survives only as one arm of a total translation from XLibur's enum to the SDK's, at
`SheetViewWriter.cs:158-163`, reachable only if the resolver ever returned `FrozenSplit`, which
it never does. Dropping the arm to satisfy the grep would make the translation silently coerce
`FrozenSplit` to `Frozen`, which is worse code. Criterion 5, written later, describes the right
test: the spelling appears in a mapping switch, not in a decision.

### What was deliberately not done

- **`WritePreColumns` / `WritePostColumns` were routed through the resolver**, not left alone.
  The spec allowed either. They build a `<col>` carrying the worksheet default width and style
  over a filler range, which is the same per-`<col>` attribute decision, and routing them removed
  the last hand-written `CustomWidth = true`. *Which* columns they write is untouched.
- **The double-rounding trap was solved by passing the raw width**, not by adding a
  pre-resolved-width overload. `ColumnWriteContext` now carries `RawWorksheetColumnWidth`
  alongside the resolved `WorksheetColumnWidth`, and `Resolve` receives the raw one so
  `GetColumnWidth(raw).SaveRound()` applies exactly once and yields the identical value. Chosen
  because it leaves the width rule with one entry point rather than two. Proved by task 1's
  column test and by the 30 `*Column*` tests.
- **`CellXmlWriter`, `SheetDataWriter` internals and `WorksheetElementReader.LoadSheetViewPane`
  were not touched**, per the non-goals.
- **The two write paths were not merged.** Both serializers emit their own XML over one resolved
  value.
- **No `<sheets>` or styles.xml resolver**, for the reasons recorded above.

### What spec 31 inherits

- `SetupPane` (`SheetViewWriter.cs:113-148`) is now a resolver call plus five assignments and two
  small `ToOpenXml` mappings (`:150-163`). `GetActivePaneValue` and `GetActivePaneForActiveCell`
  are gone. The pane decision is already extracted into `XLPaneSettings`, so 31's interface does
  not have to carry it.
- `BuildColumnElement` (`ColumnWriter.cs:134-147`) is split at its existing seam: a lookup, a
  `Resolve` call, and `ToColumnElement` (`:158-176`), which is the only place a `Column` element
  is now built. `WorksheetDefaultColumn` (`:153-156`) covers the three filler cases.
- Both writers take narrowed parameters — `IReadOnlyDictionary<XLStyleValue, StyleInfo>` rather
  than `SaveContext` — so 31's interface does not need to thread the ten-member bag.
- `WritePathAgreementTests` will notice if 31's sweep breaks either emitter. It reads the bytes,
  so unlike a load-and-compare test it can see a divergence the reader would normalise away.
