# Spec 24 — The worksheet element load gets one interface

**Area:** Architecture · Refactor
**Effort:** S–M (~3–4 days)
**Dependencies:** None hard. **Conflicts with spec 18 task 5** — both work inside
`LoadWorksheetElements`. See Conflicts.
**Status:** Proposed.

## Goal

Move the worksheet element-name dispatch inside the reader that owns the element bodies, so
`WorksheetElementReader`'s interface goes from **14 entry points to 1** and adding a worksheet
element touches one module instead of two.

## Why this spec exists

Reading a worksheet element is split across a seam for no reason the code gives:

- **`XLWorkbook_Load.cs`** holds the element-name → handler dispatch, as an `if`/`else if` chain on
  `reader.ElementType`, split across two methods (`LoadStructureElement`,
  `LoadPrintAndExtensionElement`) purely to stay under a cognitive-complexity limit.
- **`WorksheetElementReader.cs`** holds 14 element bodies, one `internal static LoadX` each, and
  knows nothing about which element name reaches it.

Adding a worksheet element means editing both. Neither module is independently comprehensible: the
loader names elements it cannot read, and the reader reads elements it cannot name.

Applying the deletion test to `WorksheetElementReader`: deleting it would move 624 lines into
`XLWorkbook_Load` and concentrate nothing. Its interface is 14 wide and each implementation is
small — a shallow module, and the only one of its kind in the IO layer. Every other reader and
writer under `XLibur/Excel/IO/` exposes one or two entry points.

The dispatch is also spread wider than the two named modules. The 17 element types are handled by
**four** different owners:

| Owner | Elements |
|---|---|
| `WorksheetElementReader` | `SheetViews`, `AutoFilter`, `SheetProtection`, `DataValidations`, `Hyperlinks`, `PrintOptions`, `PageMargins`, `PageSetup`, `HeaderFooter`, `RowBreaks`, `ColumnBreaks`, `SheetProperties` |
| `WorksheetSheetDataReader` | `Columns` |
| `ConditionalFormatReader` | `ConditionalFormatting`, `WorksheetExtensionList` |
| `XLWorkbook_Load` itself | `SheetFormatProperties`, `MergeCells`, `LegacyDrawing` |

## Non-goals

- **No change to the two-pass load.** `<sheetData>` is deliberately skipped in pass 1 and read in
  pass 2 with a raw `XmlReader` (`XLWorkbook_Load.cs:343`, `:369`). The remarks on `LoadSheetDataRaw`
  record that rescanning the part is 0.05–0.65 ms and that collapsing the passes is not worth the
  rewrite. This spec does not revisit that.
- **No behaviour change.** Same elements read, same order, same results.
- **No performance work.** Spec 18 task 5 owns the per-sheet cost. This spec must not make it worse
  and does not try to make it better — see task 4.
- **No public API change.**

## Current state

Verified against the tree at `d05b0753` (2026-08-23).

- `LoadWorksheetElements` — `XLWorkbook_Load.cs:327-370`, the two-pass driver
- `LoadWorksheetElement` — `:404-419`, calls the two dispatch halves
- `LoadStructureElement` — `:421-450`, 8 element types, returns `bool` handled
- `LoadPrintAndExtensionElement` — `:452-487`, 9 element types
- `LoadSheetFormatProperties` — `:490`, `private void`, uses `this` only as an `XLWorkbook` argument
- `LoadMergeCellsStreaming` — `:510`, already `private static`
- `CalculateColumnWidth` — `:945`, `private static (double, IXLFont, XLWorkbook)`
- `WorksheetElementReader.cs` — 624 lines, 14 `internal static` entry points

`pageSetupProperties` is threaded through the dispatch: `SheetProperties` produces it in the main
loop (`:361`), `PageSetup` consumes it (`:479`). That ordering dependency is real and must survive.

## The design

Mirror the pattern `WorksheetSheetDataReader` already uses — `in <Context>, ref <State>` — so the two
readers in the load path read the same way.

```csharp
/// <summary>Everything a worksheet element handler needs and cannot change.</summary>
internal readonly struct WorksheetElementContext
{
    internal required WorksheetPart Part { get; init; }
    internal required XLWorksheet Worksheet { get; init; }
    internal required StylesheetData Styles { get; init; }
    internal required LoadContext Load { get; init; }
    internal required XLWorkbook Workbook { get; init; }
}

/// <summary>
/// What one element hands to a later one. <c>c:sheetPr</c> produces the page-setup properties that
/// <c>c:pageSetup</c> consumes, so the two are ordered and the carrier is mutable.
/// </summary>
internal struct WorksheetElementState
{
    internal PageSetupProperties? PageSetupProperties;
}
```

`WorksheetElementReader` gains one entry point and loses fourteen:

```csharp
/// <summary>
/// Reads the element the reader is currently positioned on into the worksheet model.
/// </summary>
/// <returns>
/// <c>true</c> if the element was recognised and consumed; <c>false</c> if the caller should
/// leave it alone. A <c>false</c> return has not touched the reader.
/// </returns>
internal static bool TryLoad(
    OpenXmlPartReader reader,
    in WorksheetElementContext context,
    ref WorksheetElementState state);
```

The 14 `LoadX` methods become `private static`. The dispatch chain moves in with them.

`SheetProperties` stops being special-cased in the main loop: it is dispatched like every other
element, and writes its output into `state.PageSetupProperties`.

## File structure

```
XLibur/Excel/IO/WorksheetElementContext.cs   new — the context and state structs
XLibur/Excel/IO/WorksheetElementReader.cs    modified — gains TryLoad, 14 methods go private
XLibur/Excel/XLWorkbook_Load.cs              modified — dispatch deleted, ~70 lines lighter
```

## Global constraints

- Warnings are errors; nullable enabled.
- Branch per task; never commit to main. Commit prefix `refactor:`.
- No compound shell commands (`&&`, `;`) in agent tool calls.
- Build: `dotnet build XLibur/XLibur.csproj -c Release -v q`
- Tests: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
- Use `--treenode-filter`, never `--filter`. Never filter at solution level.
- `required` members need C# 11+; the repo targets net8.0/net9.0/net10.0, so this is available. If
  the language version is pinned lower, drop `required` and validate in `TryLoad` instead.

## Work plan

| # | Task | Size | Gate |
|---|---|---|---|
| 1 | Characterization test: every element survives a round trip | S | New test green on unmodified code |
| 2 | Context and state structs; `TryLoad` with the dispatch moved in | M | Suite green |
| 3 | Move the three orphan handlers off `XLWorkbook_Load` | S | Suite green |
| 4 | Confirm the per-sheet load cost is unchanged | S | Within noise of baseline |

---

### Task 1 — Characterization test

The refactor is behaviour-preserving, so it needs a test that would notice if an element stopped
being read. No single existing test covers all 17 at once.

**Files:**
- Create: `XLibur.Tests/Excel/IO/WorksheetElementRoundTripTests.cs`

**Interfaces:**
- Produces: `Every_worksheet_element_survives_a_round_trip`, the gate for tasks 2 and 3.

- [ ] **Step 1: Write the test**

```csharp
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// One workbook carrying every worksheet element the loader dispatches on, round-tripped and read
/// back. Spec 24 moves that dispatch from XLWorkbook_Load into WorksheetElementReader; this test is
/// what proves no element was dropped on the way.
/// </summary>
public class WorksheetElementRoundTripTests
{
    private static MemoryStream BuildWorkbookWithEveryElement()
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");

            // SheetFormatProperties
            ws.RowHeight = 22;
            ws.ColumnWidth = 14;

            // Columns
            ws.Column(2).Width = 33;

            // SheetData + MergeCells
            ws.Cell("A1").Value = "Header";
            ws.Cell("B1").Value = 42;
            ws.Range("D1:E1").Merge();

            // SheetViews
            ws.SheetView.FreezeRows(1);

            // AutoFilter
            ws.Range("A1:B1").SetAutoFilter();

            // SheetProtection
            ws.Protect("pw");

            // DataValidations
            ws.Range("A5:A6").CreateDataValidation().WholeNumber.Between(1, 10);

            // Hyperlinks
            ws.Cell("A8").SetValue("link").SetHyperlink(
                new XLHyperlink("https://example.invalid/"));

            // ConditionalFormatting
            ws.Range("B5:B6").AddConditionalFormat().WhenGreaterThan(5).Fill
                .SetBackgroundColor(XLColor.Red);

            // PrintOptions, PageMargins, PageSetup, HeaderFooter, breaks
            ws.PageSetup.CenterHorizontally = true;
            ws.PageSetup.Margins.Top = 1.25;
            ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
            ws.PageSetup.Header.Left.AddText("hdr");
            ws.PageSetup.AddHorizontalPageBreak(3);
            ws.PageSetup.AddVerticalPageBreak(3);

            // SheetProperties -> tab colour
            ws.TabColor = XLColor.Blue;

            wb.SaveAs(ms);
        }

        ms.Position = 0;
        return ms;
    }

    [Test]
    public async Task Every_worksheet_element_survives_a_round_trip()
    {
        using var ms = BuildWorkbookWithEveryElement();
        using var wb = new XLWorkbook(ms);
        var ws = wb.Worksheet("Sheet1");

        await Assert.That(ws.RowHeight).IsEqualTo(22d);                       // SheetFormatProperties
        await Assert.That(ws.Column(2).Width).IsEqualTo(33d).Within(0.01);    // Columns
        await Assert.That(ws.Cell("A1").GetString()).IsEqualTo("Header");     // SheetData
        await Assert.That(ws.MergedRanges.Count).IsEqualTo(1);                // MergeCells
        await Assert.That(ws.SheetView.SplitRow).IsEqualTo(1);                // SheetViews
        await Assert.That(ws.AutoFilter.IsEnabled).IsTrue();                  // AutoFilter
        await Assert.That(ws.Protection.IsProtected).IsTrue();                // SheetProtection
        await Assert.That(ws.DataValidations.Count()).IsEqualTo(1);           // DataValidations
        await Assert.That(ws.Cell("A8").HasHyperlink).IsTrue();               // Hyperlinks
        await Assert.That(ws.ConditionalFormats.Count()).IsEqualTo(1);        // ConditionalFormatting
        await Assert.That(ws.PageSetup.CenterHorizontally).IsTrue();          // PrintOptions
        await Assert.That(ws.PageSetup.Margins.Top).IsEqualTo(1.25).Within(0.01); // PageMargins
        await Assert.That(ws.PageSetup.PaperSize).IsEqualTo(XLPaperSize.A4Paper); // PageSetup
        await Assert.That(ws.PageSetup.Header.Left.GetText()).IsEqualTo("hdr");   // HeaderFooter
        await Assert.That(ws.PageSetup.RowBreaks.Count).IsEqualTo(1);         // RowBreaks
        await Assert.That(ws.PageSetup.ColumnBreaks.Count).IsEqualTo(1);      // ColumnBreaks
        await Assert.That(ws.TabColor).IsEqualTo(XLColor.Blue);               // SheetProperties
    }
}
```

- [ ] **Step 2: Run it and fix any API mismatches**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WorksheetElementRoundTripTests/*"`
Expected: PASS.

Some of the builder calls above are written from the public surface as documented; if one does not
compile, find the equivalent in `XLibur.Tests/Excel/` — every element here is already exercised
somewhere in the suite — and use that form. **Do not weaken an assertion to make it pass.** If an
element genuinely does not round-trip today, that is a pre-existing defect: record it, replace the
assertion with the current behaviour plus a comment naming the gap, and report it.

- [ ] **Step 3: Verify the gate bites**

In `XLWorkbook_Load.LoadPrintAndExtensionElement`, temporarily comment out the `RowBreaks` branch.
Re-run.
Expected: FAIL on the `RowBreaks` assertion. Restore the branch.

- [ ] **Step 4: Commit**

```bash
git add XLibur.Tests/Excel/IO/WorksheetElementRoundTripTests.cs
git commit -m 'test(io): round-trip every worksheet element in one workbook (spec 24 task 1)'
```

---

### Task 2 — Context, state, and `TryLoad`

**Files:**
- Create: `XLibur/Excel/IO/WorksheetElementContext.cs`
- Modify: `XLibur/Excel/IO/WorksheetElementReader.cs`
- Modify: `XLibur/Excel/XLWorkbook_Load.cs:327-487`

**Interfaces:**
- Produces: `WorksheetElementContext`, `WorksheetElementState`,
  `WorksheetElementReader.TryLoad(OpenXmlPartReader, in WorksheetElementContext, ref WorksheetElementState) → bool`.

- [ ] **Step 1: Create the context and state structs**

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Everything a worksheet element handler needs and cannot change. Passed by <c>in</c> so the
/// struct is not copied per element.
/// </summary>
internal readonly struct WorksheetElementContext
{
    internal required WorksheetPart Part { get; init; }
    internal required XLWorksheet Worksheet { get; init; }
    internal required StylesheetData Styles { get; init; }
    internal required LoadContext Load { get; init; }
    internal required XLWorkbook Workbook { get; init; }
}

/// <summary>
/// What one worksheet element hands to a later one. <c>sheetPr</c> produces the page-setup
/// properties that <c>pageSetup</c> consumes, so the two are ordered and this carrier is mutable.
/// </summary>
internal struct WorksheetElementState
{
    internal PageSetupProperties? PageSetupProperties;
}
```

- [ ] **Step 2: Add `TryLoad` to `WorksheetElementReader` with the dispatch moved in**

```csharp
    /// <summary>
    /// Reads the element the reader is positioned on into the worksheet model.
    /// </summary>
    /// <returns>
    /// <c>true</c> if the element was recognised and consumed; <c>false</c> if it was not, in which
    /// case the reader has not been touched.
    /// </returns>
    /// <remarks>
    /// Only recognised types are materialised. Calling <c>LoadCurrentElement()</c> on a wrapper such
    /// as <c>sheetData</c> would consume all of its children and starve the caller's loop, which is
    /// why every branch below names a concrete type.
    /// </remarks>
    internal static bool TryLoad(
        OpenXmlPartReader reader,
        in WorksheetElementContext context,
        ref WorksheetElementState state)
    {
        var elementType = reader.ElementType;
        var ws = context.Worksheet;

        if (elementType == typeof(SheetProperties))
        {
            LoadSheetProperties((SheetProperties)reader.LoadCurrentElement()!, ws,
                out var pageSetupProperties);
            state.PageSetupProperties = pageSetupProperties;
        }
        else if (elementType == typeof(SheetFormatProperties))
            LoadSheetFormatProperties((SheetFormatProperties)reader.LoadCurrentElement()!, ws,
                context.Workbook);
        else if (elementType == typeof(MergeCells))
            LoadMergeCellsStreaming(reader, ws);
        else if (elementType == typeof(SheetViews))
            LoadSheetViews((SheetViews)reader.LoadCurrentElement()!, ws);
        else if (elementType == typeof(Columns))
            WorksheetSheetDataReader.LoadColumns(context.Styles, ws,
                (Columns)reader.LoadCurrentElement()!);
        else if (elementType == typeof(AutoFilter))
            LoadAutoFilter((AutoFilter)reader.LoadCurrentElement()!, ws,
                context.Styles.DifferentialFormats);
        else if (elementType == typeof(SheetProtection))
            LoadSheetProtection((SheetProtection)reader.LoadCurrentElement()!, ws);
        else if (elementType == typeof(DataValidations))
            LoadDataValidations((DataValidations)reader.LoadCurrentElement()!, ws);
        else if (elementType == typeof(LegacyDrawing))
            ws.LegacyDrawingId = ((LegacyDrawing)reader.LoadCurrentElement()!).Id?.Value;
        else if (elementType == typeof(ConditionalFormatting))
            ConditionalFormatReader.LoadConditionalFormatting(
                (ConditionalFormatting)reader.LoadCurrentElement()!, ws,
                context.Styles.DifferentialFormats, context.Load);
        else if (elementType == typeof(Hyperlinks))
            LoadHyperlinks((Hyperlinks)reader.LoadCurrentElement()!, context.Part, ws);
        else if (elementType == typeof(PrintOptions))
            LoadPrintOptions((PrintOptions)reader.LoadCurrentElement()!, ws);
        else if (elementType == typeof(PageMargins))
            LoadPageMargins((PageMargins)reader.LoadCurrentElement()!, ws);
        else if (elementType == typeof(PageSetup))
            LoadPageSetup((PageSetup)reader.LoadCurrentElement()!, ws, state.PageSetupProperties);
        else if (elementType == typeof(HeaderFooter))
            LoadHeaderFooter((HeaderFooter)reader.LoadCurrentElement()!, ws);
        else if (elementType == typeof(RowBreaks))
            LoadRowBreaks((RowBreaks)reader.LoadCurrentElement()!, ws);
        else if (elementType == typeof(ColumnBreaks))
            LoadColumnBreaks((ColumnBreaks)reader.LoadCurrentElement()!, ws);
        else if (elementType == typeof(WorksheetExtensionList))
            ConditionalFormatReader.LoadExtensions(
                (WorksheetExtensionList)reader.LoadCurrentElement()!, ws, context.Workbook);
        else
            return false;

        return true;
    }
```

Sonar's cognitive-complexity rule is what split this chain in two in `XLWorkbook_Load`. A flat
dispatch chain is the clearest form it can take, so suppress the rule at the method with a reason,
matching how the repo handles `S4136` in `XLCellFormulaShifter`:

```csharp
    // S3776: a flat dispatch over element types is the clearest form this can take; splitting it
    // by an arbitrary cut is what put the element names in a different module from the element
    // bodies in the first place.
#pragma warning disable S3776
```

`LoadSheetFormatProperties` and `LoadMergeCellsStreaming` are referenced above but still live on
`XLWorkbook_Load` — task 3 moves them. Until then, make them `internal static` on `XLWorkbook` and
call them qualified, or do task 3 first. **Doing task 3 first is cleaner if the build complains.**

- [ ] **Step 3: Make the 14 existing entry points private**

Every `internal static void LoadX` in `WorksheetElementReader.cs` becomes `private static void`,
except `LoadAutoFilterColumns` and `LoadAutoFilterSort`, which are called from elsewhere. Confirm
which with:

Run: `grep -rn 'WorksheetElementReader\.' XLibur --include=*.cs`

Anything still called from outside stays `internal`; everything else goes `private`.

- [ ] **Step 4: Collapse the loader**

Replace `LoadWorksheetElements`' pass-1 loop, and delete `LoadWorksheetElement`,
`LoadStructureElement` and `LoadPrintAndExtensionElement` entirely:

```csharp
        var elementContext = new WorksheetElementContext
        {
            Part = worksheetPart,
            Worksheet = ws,
            Styles = styles,
            Load = context,
            Workbook = this,
        };
        var elementState = default(WorksheetElementState);

        // Pass 1: structural elements via the OpenXML SDK reader (the proven DOM path). The
        // <sheetData> hot path is skipped here — it is read in pass 2 with a raw XmlReader, which
        // is ~4x faster and allocates ~5x less than materializing every cell through the SDK
        // reader's object model. Structural elements such as <cols> are parsed here (before pass 2
        // runs), so column styles are already available when cells resolve their inherited style.
        using (var reader = new OpenXmlPartReader(worksheetPart))
        {
            while (reader.Read())
            {
                // Skipped wholesale, without descending:
                //  - CustomSheetViews carries its own auto filter data and more, ignored for now.
                //  - SheetData is read in pass 2 by the raw reader.
                // ReadNextSibling leaves the reader *on* the next sibling rather than needing
                // another Read, which is why this is a leading loop rather than a `continue`.
                while (reader.ElementType == typeof(CustomSheetViews)
                       || reader.ElementType == typeof(SheetData))
                    reader.ReadNextSibling();

                WorksheetElementReader.TryLoad(reader, in elementContext, ref elementState);
            }
        }

        // Pass 2: read <sheetData> rows/cells directly from a raw XmlReader.
        LoadSheetDataRaw(worksheetPart, in sheetDataContext, ref sheetDataState);
```

Note what disappears: the `SheetProperties` special case at `:359-365` is gone — it is now one branch
among the rest, and `pageSetupProperties` lives in `elementState` rather than as a local threaded
through four signatures.

- [ ] **Step 5: Build and run the full suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS, including task 1's round-trip test.

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel/IO/WorksheetElementContext.cs XLibur/Excel/IO/WorksheetElementReader.cs XLibur/Excel/XLWorkbook_Load.cs
git commit -m 'refactor(io): move the worksheet element dispatch into its reader (spec 24 task 2)'
```

---

### Task 3 — Move the three orphan handlers

`SheetFormatProperties`, `MergeCells` and `LegacyDrawing` are read by `XLWorkbook_Load` itself. With
the dispatch gone they are the only reason it still knows about worksheet elements.

**Files:**
- Modify: `XLibur/Excel/XLWorkbook_Load.cs` — delete `LoadSheetFormatProperties` (`:490`) and
  `LoadMergeCellsStreaming` (`:510`); widen `CalculateColumnWidth` (`:945`) to `internal static`
- Modify: `XLibur/Excel/IO/WorksheetElementReader.cs` — receive both

- [ ] **Step 1: Move `LoadMergeCellsStreaming` verbatim**

It is already `private static (OpenXmlPartReader, XLWorksheet)` and has no dependency on the
workbook. Cut it into `WorksheetElementReader` as `private static`, unchanged.

- [ ] **Step 2: Move `LoadSheetFormatProperties`, taking the workbook explicitly**

It is `private void` today and uses `this` only as an argument. Move it as:

```csharp
    private static void LoadSheetFormatProperties(SheetFormatProperties sfp, XLWorksheet ws,
        XLWorkbook workbook)
    {
        if (sfp.DefaultRowHeight is not null)
            ws.RowHeight = sfp.DefaultRowHeight;

        ws.RowHeightChanged = sfp.CustomHeight is not null && sfp.CustomHeight.Value;

        if (sfp.DefaultColumnWidth is not null)
            ws.ColumnWidth = XLHelper.ConvertWidthToNoC(sfp.DefaultColumnWidth.Value,
                ws.Style.Font, workbook);
        else if (sfp.BaseColumnWidth is not null)
            ws.ColumnWidth = XLWorkbook.CalculateColumnWidth(sfp.BaseColumnWidth.Value,
                ws.Style.Font, workbook);

        // ... remainder moved verbatim
    }
```

Widen `CalculateColumnWidth` from `private static` to `internal static` on `XLWorkbook`. It already
takes the workbook as a parameter, so nothing else changes.

- [ ] **Step 3: Confirm `XLWorkbook_Load` no longer names a worksheet element**

Run: `grep -nE 'typeof\((SheetFormatProperties|MergeCells|LegacyDrawing|SheetViews|PageSetup|HeaderFooter|RowBreaks|ColumnBreaks|Hyperlinks|PrintOptions|PageMargins|AutoFilter|SheetProtection|DataValidations|ConditionalFormatting|WorksheetExtensionList|SheetProperties|Columns)\)' XLibur/Excel/XLWorkbook_Load.cs`

Expected: no output. `CustomSheetViews` and `SheetData` may still appear — they are the two the
loader deliberately skips, which is sequencing, not element knowledge.

- [ ] **Step 4: Build and run the full suite on both frameworks**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Expected: PASS on net8.0 and net10.0.

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/IO/WorksheetElementReader.cs XLibur/Excel/XLWorkbook_Load.cs
git commit -m 'refactor(io): move the last three element handlers off the loader (spec 24 task 3)'
```

---

### Task 4 — Confirm the per-sheet load cost is unchanged

`LoadWorksheetElements` is 23.7% of the template round-trip trace (spec 18 task 5). This spec
restructures it, so the cost must be shown not to have moved.

- [ ] **Step 1: Measure the merge-base**

```
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- profile template
```

Record the header-only sheet sweep (1 / 10 / 40 sheets) from spec 18 task 5's table.

- [ ] **Step 2: Measure the branch**

Same command, same fixture.

- [ ] **Step 3: Compare the per-sheet slope**

Expected: ~1.0 ms and ~0.19 MB per header-only sheet, unchanged.

The context struct is passed by `in` and the state struct by `ref`, so neither is copied per
element and neither allocates. If allocation per sheet has risen, the likely cause is the context
struct being copied — check that every call site uses `in`, not a bare argument.

**Decision rule.** A per-sheet regression above the ~1.2 ms probe noise floor that spec 18 records
must be explained before this spec lands, not after. Record the numbers in a Results section either
way.

- [ ] **Step 4: Commit the Results section**

```bash
git add docs/specs/24-worksheet-element-dispatch.md
git commit -m 'docs(specs): record the per-sheet load numbers for spec 24'
```

---

## Acceptance criteria

1. `WorksheetElementReader` exposes **one** dispatch entry point, `TryLoad`, plus only those members
   proven to have external callers in task 2 step 3.
2. `XLWorkbook_Load.cs` names no worksheet element type except `CustomSheetViews` and `SheetData`.
3. `LoadWorksheetElement`, `LoadStructureElement` and `LoadPrintAndExtensionElement` no longer exist.
4. `pageSetupProperties` is no longer a local threaded through four signatures.
5. Task 1's round-trip test passes, with no assertion weakened.
6. Full suite green on net8.0 and net10.0.
7. Per-sheet structural load cost within spec 18's noise floor of its pre-spec value.
8. No public API change.

## Conflicts

- **Spec 18 task 5** is the live conflict. It owns the per-sheet structural cost and attributes
  23.7% of the template trace to `LoadWorksheetElements` — the exact method this spec restructures.
  The two cannot run concurrently.

  **Recommended order: this spec first.** It is behaviour-preserving and mechanical, and it leaves
  spec 18 task 5 a single method to optimise instead of four spread across two modules. Spec 18
  task 5 is still only "attributed", not designed, so it has nothing to rebase yet.

  If spec 18 task 5 starts first, this spec waits — do not rebase a performance change onto a
  structural one.
- **Spec 02** (load-path allocations) is done and touches `WorksheetSheetDataReader`, not this
  dispatch. `WorksheetSheetDataReader.LoadColumns` is called from `TryLoad` but not modified.
- No other spec touches `WorksheetElementReader.cs`.
