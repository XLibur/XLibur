# Spec 31 — The worksheet element writers get one interface

**Area:** Architecture · Refactor
**Effort:** M–L (~1–1.5 weeks)
**Dependencies:** **Spec 29 must land first** — it touches `SheetViewWriter.cs` and `ColumnWriter.cs`,
which this spec rewrites wholesale. See Conflicts. **Spec 29 landed as #413 (`8d2acfc7`) on
2026-08-27, so this is unblocked.**
**Status:** Proposed — unblocked.

> ## ⚠️ Read this before touching `SheetViewWriter`
>
> **`SplitRow` and `SplitColumn` carry two different units, and the unit lives in a sibling boolean
> rather than in the type.** After D18 (PR #416, `37c986bb`) a pane is a freeze or an unfrozen split:
> for a freeze the two numbers count frozen lines, for a split they are the split bar's position in
> twentieths of a point. `FreezePanes` is what says which. `int SplitRow` carries no unit and its
> public setter validates nothing.
>
> **This spec is the next set of hands on exactly that code**, and the agent that fixed D18 named it
> as the most likely place the distinction gets broken next. Its guidance, carried here verbatim
> because it will not be obvious from the code:
>
> - **Never read `SplitRow`/`SplitColumn` without `FreezePanes` in the same expression.** When
>   auditing, grep for the three together.
> - **Any new copy or clone path that takes the two numbers without the flag flips the units
>   silently.** `XLSheetView`'s copy constructor is correct only because it happens to copy all three.
> - **Any expression deriving a cell from `split + 1` is a landmine.** That is a cell address for a
>   freeze and meaningless for a split; an unfrozen split anchors its `topLeftCell` at `A1`.
>   `ScrollIntoView` was one instance of a pattern, not a one-off.
> - **No round-trip test can catch a unit error here**, because the reader normalises both spellings
>   into the same ints. Only byte-level assertions can — the same blindness spec 29 was written about.
>
> **The durable fix, deliberately not done by D18:** make the numbers and the mode inseparable — one
> struct holding both values and the unit, so the values cannot be obtained without the unit. If this
> spec is rewriting these writers anyway, that is the moment to do it.

## Goal

Give every worksheet element the save path emits exactly one owner, declared by the schema slot it
writes, so `GetWorksheetDom` stops naming 21 writer entry points across 10 classes in 6 signature
shapes and the required element order stops being stated in two places at once.

This is the save-side mirror of spec 24. 24 moved the *read* dispatch into the reader. 31 moves the
*write* order into data.

## Why this spec exists

### `GetWorksheetDom` is a 91-line call sequence that encodes the ECMA-376 element order

`XLibur/Excel/IO/WorksheetPartWriter.cs:133-224`. Twenty-one calls into ten other classes, in an
order that is load-bearing and written down nowhere:

| Line | Call | Slot |
|---|---|---|
| `:166` | `SheetViewWriter.WriteSheetProperties` | 1 SheetProperties |
| `:167` | `SheetViewWriter.WriteSheetDimension` | 2 SheetDimension |
| `:168` | `SheetViewWriter.WriteSheetViews` | 3 SheetViews |
| `:178` | `SheetViewWriter.WriteSheetFormatProperties` | 4 SheetFormatProperties |
| `:181` | `ColumnWriter.WriteColumns` | 5 Columns |
| `:185-192` | *inline* SheetData placeholder | 6 SheetData |
| `:200` | `SheetProtectionWriter.WriteSheetProtection` | 8 SheetProtection |
| `:201` | `AutoFilterWriter.WriteAutoFilter` | 11 AutoFilter |
| `:203` | `WorksheetPartWriter.WriteMergeCells` — *private, same file* | 15 MergeCells |
| `:205` | `ConditionalFormattingWriter.WriteConditionalFormatting` | 17 + **40** |
| `:206` | `ConditionalFormattingWriter.WriteSparklines` | **40** |
| `:207` | `DataValidationWriter.WriteDataValidations` | 18 + **40** |
| `:208` | `PageSetupWriter.WriteHyperlinks` | 19 Hyperlinks |
| `:209` | `PageSetupWriter.WritePrintOptions` | 20 PrintOptions |
| `:210` | `PageSetupWriter.WritePageMargins` | 21 PageMargins |
| `:211` | `PageSetupWriter.WritePageSetup` | 22 PageSetup |
| `:212` | `PageSetupWriter.WriteHeaderFooter` | 23 HeaderFooter |
| `:213` | `PageSetupWriter.WriteRowBreaks` | 24 RowBreaks |
| `:214` | `PageSetupWriter.WriteColumnBreaks` | 25 ColumnBreaks |
| `:216` | `WorksheetPartWriter.PopulateTablePartReferences` — *private, same file* | 39 TableParts |
| `:218` | `PictureWriter.WriteDrawings` | **30** Drawing |
| `:219` | `ChartWriter.WriteCharts` | **30** Drawing |
| `:220` | `PictureWriter.WriteLegacyDrawing` | 31 LegacyDrawing |
| `:221` | `HeaderFooterImageWriter.WriteHeaderFooterImages` | 32 LegacyDrawingHeaderFooter |

Twenty-one of those are calls into other classes; two (`:203`, `:216`) are private methods of
`WorksheetPartWriter` itself, and one (`:185-192`) is not a method at all — the SheetData slot
ceremony is written out inline. Note the two bold columns: **slot 40 has three writers and slot 30
has two.** More on both below.

### Six signature shapes for one job

| Shape | Count | Lines |
|---|---:|---|
| `(Worksheet, XLWorksheetContentManager, XLWorksheet)` | 12 | `:166 :167 :168 :200 :206 :209 :210 :211 :212 :213 :214 :220` |
| `… + SaveContext` | 2 | `:201 :205` |
| `… + SaveOptions` | 1 | `:207` |
| `… + WorksheetPart, SaveContext` | 4 | `:208 :218 :219 :221` |
| `… + double, SaveContext` | 1 | `:181` |
| `… + int, int, out double` | 1 | `:178` |

The last two are one dependency spelled as two signatures.
`SheetViewWriter.WriteSheetFormatProperties` (`SheetViewWriter.cs:244-251`) produces
`worksheetColumnWidth` through an `out` parameter and `ColumnWriter.WriteColumns`
(`ColumnWriter.cs:20-25`) consumes it as a `double` — a producer/consumer pair between slots 4 and 5,
carried by a local in the driver. This is the exact shape spec 24 found on the load side, where
`sheetPr` produces the `PageSetupProperties` that `pageSetup` consumes.

### The ceremony, twenty times

```csharp
if (!worksheet.Elements<SheetProtection>().Any())
{
    var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.SheetProtection);
    worksheet.InsertAfter(new SheetProtection(), previousElement);
}

var sheetProtection = worksheet.Elements<SheetProtection>().First();
cm.SetElement(XLWorksheetContents.SheetProtection, sheetProtection);
```

`SheetProtectionWriter.cs:17-24`. The same six lines, type substituted, appear at
`AutoFilterWriter.cs:24`, `ColumnWriter.cs:38`, `ConditionalFormattingWriter.cs:112` and `:210`,
`DataValidationWriter.cs:99` and `:182`, `PageSetupWriter.cs:32 :72 :92 :113 :192 :233 :287`,
`PictureWriter.cs:159`, `WorksheetPartWriter.cs:187 :232 :268`.

`grep -rn "GetPreviousElementFor" XLibur --include=*.cs` returns **24 lines**. Three of those are in
`XLibur/Excel/ContentManagers/XLBaseContentManager.cs` — the declaration at `:53` and two doc-comment
references at `:14` and `:69`. That leaves **21 call sites across 10 files**, of which
`SheetViewWriter.cs:220` addresses a different content manager (`XLSheetViewContents.Selection`,
inside a `<sheetView>`). So: **20 call sites against `XLWorksheetContents`, across 9 files.**

### The order is written down twice

`XLibur/Excel/ContentManagers/XLWorksheetContentManager.cs:6-48` declares
`XLWorksheetContents` — 40 numbered slots, `SheetProperties = 1` at `:8` through
`WorksheetExtensionList = 40` at `:47`. That enum *is* the ECMA-376 child order for
`<worksheet>`; `GetPreviousElementFor` (`XLBaseContentManager.cs:53-63`) walks it backwards to find
an insertion anchor.

The call sequence in `GetWorksheetDom` states the same order a second time. Nothing checks that the
two agree. A writer inserted at the wrong point in the sequence produces a worksheet whose children
are in schema-invalid order, and the only thing that would notice is Excel refusing to open the file.

Of the 40 slots, **23 are ever written** — confirmed by
`grep -rhoE "XLWorksheetContents\.[A-Za-z]+" XLibur/Excel/IO/*.cs | sort -u`. The other 17
(`SheetCalculationProperties`, `ProtectedRanges`, `Scenarios`, `SortState`, `DataConsolidate`,
`CustomSheetViews`, `PhoneticProperties`, `CustomProperties`, `CellWatches`, `IgnoredErrors`,
`SmartTags`, `DrawingHeaderFooter`, `Picture`, `OleObjects`, `Controls`, `AlternateContent`,
`WebPublishItems`) are pass-through: they exist in the loaded DOM, XLibur has no model for them, and
they survive because nothing rewrites them. `docs/round-trip-fidelity.md` and
`XLibur.Tests/Excel/RoundTripFidelityTests.cs` pin that behaviour. **They still have to end up in
the right place relative to the 23 that are rewritten**, which is what makes the anchor walk load-bearing.

### Slot 30 has two owners, and the call order is what keeps it from throwing

`PictureWriter.cs:40`:

```csharp
        var tableParts = worksheet.Elements<TableParts>().First();
```

Unconditional `.First()`. It throws `InvalidOperationException` on an empty sequence. It does not
throw today only because `PopulateTablePartReferences` runs two lines earlier at
`WorksheetPartWriter.cs:216` and guarantees a `<tableParts>` exists (it creates one if absent —
`:262-273`). `PictureWriter` then positions the `<drawing>` with
`worksheet.InsertBefore(worksheetDrawing, tableParts)` at `:45` — bypassing the content manager for
placement while still calling `cm.SetElement(XLWorksheetContents.Drawing, …)` at `:46`.

`ChartWriter.EnsureDrawingElement` (`ChartWriter.cs:1118-1133`) writes the *same* slot 30 element and
does the same lookup defensively:

```csharp
            var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
            …
            if (tableParts != null)
                worksheet.InsertBefore(drawingRef, tableParts);
            else
                worksheet.AppendChild(drawingRef);
```

Two implementations of one slot, disagreeing on whether `<tableParts>` is guaranteed. One of them is
right about the current call order and one of them is defensive; neither states the assumption.
Reordering `GetWorksheetDom` by two lines turns `PictureWriter.cs:40` into a crash. **This is a
defect in waiting, not a live defect** — the ordering invariant holds at `1b41cadd`. Task 1 pins it.

### Slot 40 has three owners

`WorksheetExtensionList` is created-or-found *and* deleted-if-empty by three independent modules:

| Module | Creates extLst | Deletes extLst |
|---|---|---|
| `ConditionalFormattingWriter.WriteExtensionDataBars` (`:100-163`) | `:110-117` | **never** |
| `ConditionalFormattingWriter.WriteSparklineGroups` (`:205`) | `:208-215` | — |
| `ConditionalFormattingWriter.RemoveSparklineExtension` (`:182-203`) | — | `:198-202` |
| `DataValidationWriter.WriteExtensionDataValidationElements` (`:174`) | `:180-187` | — |
| `DataValidationWriter.RemoveExtensionDataValidations` (`:151-172`) | — | `:167-171` |

Three modules, five entry points into one slot, none of which is the slot's owner.

**Does the collision fire? No — it is latent, and the honest finding is that the safety does not
come from the call order.** Working the cases:

- `WriteExtensionDataBars` (`:205`, first) only ever *adds*. When it adds, the extLst always ends up
  with at least one `<ext>` child (`:123-129`), so the two later removals — both guarded by
  `if (!worksheetExtensionList.HasChildren)` — cannot take it away.
- `RemoveSparklineExtension` (`:206`) removes the `<ext>` only when it carries the sparkline URI and
  is childless after `RemoveAllChildren<X14.SparklineGroups>()` (`:190-196`), then removes the extLst
  only when *that* is childless.
- `RemoveExtensionDataValidations` (`:207`) is the identical shape.

So what protects slot 40 is a **shared invariant — "remove only when childless" — implemented three
times by hand**, not the sequence. That is the same defect shape the README's round-2 table records
five instances of. Two pieces of evidence that the hand-maintained copies have already started to drift:

1. The same URI comparison is spelled two ways. `ConditionalFormattingWriter.cs:188` uses
   `StringComparison.InvariantCultureIgnoreCase`; `DataValidationWriter.cs:157` uses
   `StringComparison.OrdinalIgnoreCase`.
2. **`WriteExtensionDataBars` has no removal branch at all.** With `exlst.Length == 0` (`:108`) it
   returns having done nothing, so a stale `x14:conditionalFormattings` inherited from a loaded file
   survives after the data-bar format is deleted from the model. The other two extension writers both
   have a `Remove*` path; this one does not. That asymmetry is worth confirming under task 1 — if it
   reproduces, it is a live round-trip defect that this spec's slot owner fixes for free.

Order also decides something no schema constrains: the `<ext>` children come out in call order
(data bars, sparklines, data validations) purely because `:205 :206 :207` run in that order.

### A correction to spec 24

Spec 24, lines 30–31, states:

> Every other reader and writer under `XLibur/Excel/IO/` exposes one or two entry points.

**On the save side that is false.** Counted at `1b41cadd`:

| Class | `internal static` entry points |
|---|---:|
| `CellXmlWriter` | 12 (`:67 :73 :84 :107 :126 :142 :152 :162 :173 :183 :199 :227`) |
| `PageSetupWriter` | 7 (`:14 :65 :85 :106 :181 :223 :277`) |
| `SheetViewWriter` | 4 (`:15 :42 :62 :244`) |
| `ChartWriter` | 3 (`:32 :95 :1203`) |

Spec 24 was looking at the load path, where the claim holds. It did not look at the save path, and
the save path is where the pattern it named repeats — wider, in more classes. **This spec is the
mirror 24 did not look at.** The series treats a disproved premise as a real result; this is one,
and it is 24's premise, not this spec's.

### Baseline counts

Verified with grep, not assumed:

- `XLibur/Excel/IO/` holds **30** files matching `*Writer*.cs`. All 30 are `internal static class`.
- They expose **61** `internal static` entry points under **51** distinct method names.
- `grep -rE "interface I.*Writer|: I.*Writer|abstract class.*Writer" XLibur --include=*.cs` returns
  **nothing**. There is no writer abstraction anywhere in the library.
- `XLibur/Excel/XLWorkbook_Save.cs` contains **20** textual `*Writer.Member(` matches. One
  (`:536`, `new RichDataWriter.RichValueEntry(…)`) constructs a nested type rather than calling a
  method, so it is **19 calls across 17 classes**. *(An earlier draft of this spec said 18 across 14;
  recounted and corrected.)*
- `GetWorksheetDom` names **21 calls across 10 external classes**, plus 2 private methods on
  `WorksheetPartWriter` itself. *(An earlier draft said 11 classes; the eleventh is the host file.)*

### The deletion test

Delete `SheetProtectionWriter` and its 93 lines move into `WorksheetPartWriter`, concentrating
nothing — the module's interface is one method, its implementation is one method. Delete
`PageSetupWriter` and 328 lines move, but as seven unrelated methods writing seven unrelated slots,
so it concentrates nothing either: it is a namespace with a `.cs` extension. The grouping of the 23
slot writers into 10 classes has no principle behind it. `SheetViewWriter` owns
`<sheetPr>`, `<dimension>`, `<sheetViews>` and `<sheetFormatPr>`; `PageSetupWriter` owns
`<hyperlinks>` as well as the six print elements. The classes are alphabetically plausible and
semantically arbitrary.

## Non-goals

- **No change to the emitted XML.** Byte-identical output is the gate on every task. This is not a
  place to fix schema-order bugs or tidy attribute emission; anything found goes in Results and gets
  its own change.
- **Not touching the readers.** Spec 24 owns the load-side dispatch.
- **Not touching `CellXmlWriter` or `SheetDataWriter`.** Specs 01 and 03 own the cell/row leaves.
  `CellXmlWriter`'s 12 entry points are cited above as evidence against spec 24's claim, not as work.
- **Not touching the drawing family** — `PictureWriter`, `ChartWriter`, `HeaderFooterImageWriter`,
  slots 30/31/32. Specs 15, 16 and 17 are all in `PictureWriter.cs`, and the README already records
  31↔15/16/17 as hard. Those four calls (`:218-221`) stay explicit at the tail of the driver, with
  the reason written down in the code.
- **Not a performance spec.** Task 10 exists to show the per-sheet save cost did not move, not to
  improve it. Spec 03 owns save-path cost.
- **No public API change.** Everything here is `internal`.

## Current state

Verified against the tree at `1b41cadd` (2026-08-24).

- `WorksheetPartWriter.GetWorksheetDom` — `XLibur/Excel/IO/WorksheetPartWriter.cs:133-224`
- `WorksheetPartWriter.WriteMergeCells` — `:226-252`, `private static`
- `WorksheetPartWriter.PopulateTablePartReferences` — `:254-283`, `private static`
- Inline SheetData slot ceremony — `:185-192`
- `XLWorksheetContents` — `XLibur/Excel/ContentManagers/XLWorksheetContentManager.cs:6-48`, 40 slots
- `XLBaseContentManager.GetPreviousElementFor` — `:53-63`; `SetElement` — `:65`
- `SheetViewWriter.cs` — 280 lines, entry points `:15 :42 :62 :244`
- `PageSetupWriter.cs` — 328 lines, entry points `:14 :65 :85 :106 :181 :223 :277`
- `ColumnWriter.cs` — entry points `:20`, `:185`; ceremony `:38`
- `SheetProtectionWriter.cs` — 103 lines, one entry point `:10`, ceremony `:17-24`
- `AutoFilterWriter.cs` — entry points `:15`, `:38`; ceremony `:24`
- `ConditionalFormattingWriter.cs` — 328 lines; `:18` `WriteConditionalFormatting`,
  `:100` `WriteExtensionDataBars`, `:165` `WriteSparklines`, `:182` `RemoveSparklineExtension`,
  `:205` `WriteSparklineGroups`
- `DataValidationWriter.cs` — 249 lines; `:16` `WriteDataValidations`,
  `:84` `WriteStandardDataValidations`, `:135` `WriteExtensionDataValidations`,
  `:151` `RemoveExtensionDataValidations`, `:174` `WriteExtensionDataValidationElements`
- `PictureWriter.cs:40` — the unguarded `.First()`; `:46 :146` slot 30; `:159-168` slot 31
- `ChartWriter.cs:1118-1133` — `EnsureDrawingElement`, the second slot-30 owner
- `HeaderFooterImageWriter.cs:24` — slot 32; ceremony `:70-74`
- `SaveContext` — `XLibur/Excel/XLWorkbook_Save.NestedTypes.cs:14`, `internal sealed class`
- `SaveOptions` — `XLibur/Excel/SaveOptions.cs:5`, `public class`

**No golden-XML harness exists in the tree.** Spec 22's task 0 prescribes
`ChartGoldenCorpus.CaptureChartPartXml`, and spec 22 reports itself done on branch `task/22`, but
`grep -rn "CaptureChartPartXml" XLibur.Tests` returns nothing at `1b41cadd`. Task 0 below builds its
own, for worksheet parts rather than chart parts.

## File structure

```
XLibur/Excel/IO/WorksheetElements/IXLWorksheetElementWriter.cs   new — the interface
XLibur/Excel/IO/WorksheetElements/WorksheetWriteContext.cs       new — context struct + slot helpers
XLibur/Excel/IO/WorksheetElements/WorksheetWriteState.cs         new — the slot-4 -> slot-5 carrier
XLibur/Excel/IO/WorksheetElements/SheetPropertiesWriter.cs       new — slot 1   (from SheetViewWriter:15)
XLibur/Excel/IO/WorksheetElements/SheetDimensionWriter.cs        new — slot 2   (from SheetViewWriter:42)
XLibur/Excel/IO/WorksheetElements/SheetViewsWriter.cs            new — slot 3   (from SheetViewWriter:62)
XLibur/Excel/IO/WorksheetElements/SheetFormatPropertiesWriter.cs new — slot 4   (from SheetViewWriter:244)
XLibur/Excel/IO/WorksheetElements/ColumnsWriter.cs               new — slot 5   (from ColumnWriter:20)
XLibur/Excel/IO/WorksheetElements/SheetDataPlaceholderWriter.cs  new — slot 6   (from WorksheetPartWriter:185)
XLibur/Excel/IO/WorksheetElements/SheetProtectionElementWriter.cs new — slot 8  (from SheetProtectionWriter:10)
XLibur/Excel/IO/WorksheetElements/AutoFilterElementWriter.cs     new — slot 11  (from AutoFilterWriter:15)
XLibur/Excel/IO/WorksheetElements/MergeCellsWriter.cs            new — slot 15  (from WorksheetPartWriter:226)
XLibur/Excel/IO/WorksheetElements/ConditionalFormattingElementWriter.cs new — slot 17
XLibur/Excel/IO/WorksheetElements/DataValidationsElementWriter.cs new — slot 18
XLibur/Excel/IO/WorksheetElements/HyperlinksWriter.cs            new — slot 19  (from PageSetupWriter:14)
XLibur/Excel/IO/WorksheetElements/PrintOptionsWriter.cs          new — slot 20  (from PageSetupWriter:65)
XLibur/Excel/IO/WorksheetElements/PageMarginsWriter.cs           new — slot 21  (from PageSetupWriter:85)
XLibur/Excel/IO/WorksheetElements/PageSetupElementWriter.cs      new — slot 22  (from PageSetupWriter:106)
XLibur/Excel/IO/WorksheetElements/HeaderFooterWriter.cs          new — slot 23  (from PageSetupWriter:181)
XLibur/Excel/IO/WorksheetElements/RowBreaksWriter.cs             new — slot 24  (from PageSetupWriter:223)
XLibur/Excel/IO/WorksheetElements/ColumnBreaksWriter.cs          new — slot 25  (from PageSetupWriter:277)
XLibur/Excel/IO/WorksheetElements/TablePartsWriter.cs            new — slot 39  (from WorksheetPartWriter:254)
XLibur/Excel/IO/WorksheetElements/WorksheetExtensionListWriter.cs new — slot 40, the three-way merge

XLibur/Excel/IO/WorksheetPartWriter.cs        modified — driver becomes a loop; 2 privates deleted
XLibur/Excel/IO/SheetViewWriter.cs            deleted  — 4 entry points become 4 slot writers
XLibur/Excel/IO/ColumnWriter.cs               modified — WriteColumns moves; GetColumnWidth:185 stays
XLibur/Excel/IO/SheetProtectionWriter.cs      deleted
XLibur/Excel/IO/AutoFilterWriter.cs           modified — WriteAutoFilter moves; PopulateAutoFilter:38 stays
XLibur/Excel/IO/ConditionalFormattingWriter.cs modified — extLst halves move to the slot-40 owner
XLibur/Excel/IO/DataValidationWriter.cs       modified — extLst halves move to the slot-40 owner
XLibur/Excel/IO/PageSetupWriter.cs            deleted  — 7 entry points become 7 slot writers

XLibur.Tests/Excel/IO/WorksheetGoldenCorpus.cs        new — byte-identity capture helper
XLibur.Tests/Excel/IO/WorksheetGoldenCorpusTests.cs   new — the gate for every task
XLibur.Tests/Excel/IO/WorksheetSlotOrderTests.cs      new — one owner per slot, ascending
XLibur.Tests/Resource/Other/Worksheets/Golden/        new — committed .xml fixtures
```

`AutoFilterWriter.PopulateAutoFilter` (`:38`) and `ColumnWriter.GetColumnWidth` (`:185`) have callers
outside the worksheet element path. Confirm with grep before deleting either file; if a file still
has an external caller, keep the file and move only the slot method out of it.

## The design

### One interface

```csharp
using XLibur.Excel.ContentManagers;

namespace XLibur.Excel.IO.WorksheetElements;

/// <summary>
/// Writes one child of <c>&lt;worksheet&gt;</c>. The slot is the element's position in the
/// ECMA-376 child order, and the driver runs the writers in slot order — so the order is data on
/// each writer rather than a fact restated by the sequence of calls in
/// <see cref="WorksheetPartWriter"/>.
/// </summary>
internal interface IXLWorksheetElementWriter
{
    /// <summary>The one slot this writer owns. No two writers may declare the same slot.</summary>
    XLWorksheetContents Slot { get; }

    /// <summary>
    /// Emits, updates or removes the element in <see cref="Slot"/>. Implementations that add an
    /// element must go through <see cref="WorksheetWriteContext.EnsureElement{T}"/>; implementations
    /// that remove one must go through <see cref="WorksheetWriteContext.RemoveElement{T}"/>, so the
    /// content manager can never fall out of step with the DOM.
    /// </summary>
    void Write(in WorksheetWriteContext ctx, ref WorksheetWriteState state);
}
```

### The interface the prompt for this spec proposed does not fit, and here is why

The obvious shape is `OpenXmlElement? Write(in ctx)` — the writer returns a new element, the driver
inserts it and calls `SetElement`. That reads well and **does not survive contact with the code.**
Four of the 23 slot owners never create a detached element at all:

```csharp
        worksheet.SheetProperties ??= new SheetProperties();     // SheetViewWriter.cs:20
        worksheet.SheetDimension  ??= new SheetDimension { … };  // SheetViewWriter.cs:56
        worksheet.SheetViews      ??= new SheetViews();          // SheetViewWriter.cs:66
        worksheet.SheetFormatProperties ??= new SheetFormatProperties(); // SheetViewWriter.cs:252
```

Those are the SDK's strongly-typed child properties. The SDK places the element in schema order
itself; there is nothing to hand back and nothing for a driver to insert. That is also why
`SheetViewWriter` has only one `GetPreviousElementFor` call in 280 lines, and why it is for a
`<selection>` inside `<sheetView>` rather than for any of its four slots.

Three more owners do not fit either: the extLst writers mutate a shared element in place, and
`PopulateTablePartReferences` finds-or-creates *and* then rewrites the children of what it found.

So `Write` returns `void`, and the ceremony is factored into a **helper the writers call** rather
than a step the driver performs. That still collapses 20 hand-written copies into one, which is the
whole benefit; it just does not also let the driver own insertion for the seven owners where
insertion is not a thing that happens.

**Record this as a design premise that was tested and failed**, not as a shortcut. If a later reader
wants to revisit it, the disqualifier is `SheetViewWriter.cs:20/56/66/252`.

### The context

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using XLibur.Excel.ContentManagers;

namespace XLibur.Excel.IO.WorksheetElements;

/// <summary>
/// Everything a worksheet element writer needs and cannot replace. Passed by <c>in</c> so it is not
/// copied per writer. Every field is a reference, so the struct is one word per member and copying
/// it would be cheap anyway — <c>in</c> is for intent, not for cost.
/// </summary>
internal readonly struct WorksheetWriteContext
{
    internal required Worksheet Worksheet { get; init; }
    internal required XLWorksheetContentManager Cm { get; init; }
    internal required XLWorksheet XlWorksheet { get; init; }
    internal required WorksheetPart Part { get; init; }
    internal required SaveContext Save { get; init; }
    internal required SaveOptions Options { get; init; }

    /// <summary>
    /// Finds the element in <paramref name="slot"/>, creating and anchoring it if absent, and
    /// records it with the content manager.
    /// </summary>
    /// <remarks>
    /// This is the ceremony that was written out by hand at 20 sites before spec 31. The anchor is
    /// <c>GetPreviousElementFor</c>, which walks <see cref="XLWorksheetContents"/> backwards for the
    /// highest occupied slot below this one; <c>InsertAfter</c> accepts a null anchor and inserts
    /// first, which is what makes slot 1 work without a special case.
    /// </remarks>
    internal T EnsureElement<T>(XLWorksheetContents slot) where T : OpenXmlElement, new()
    {
        if (!Worksheet.Elements<T>().Any())
            Worksheet.InsertAfter(new T(), Cm.GetPreviousElementFor(slot));

        var element = Worksheet.Elements<T>().First();
        Cm.SetElement(slot, element);
        return element;
    }

    /// <summary>Removes every element of the slot's type and clears the slot.</summary>
    internal void RemoveElement<T>(XLWorksheetContents slot) where T : OpenXmlElement
    {
        Worksheet.RemoveAllChildren<T>();
        Cm.SetElement(slot, null);
    }
}
```

`required` needs C# 11. The repo targets net8.0/net9.0/net10.0, so it is available; spec 24's context
struct uses it too. If the language version turns out to be pinned lower, drop `required` and
validate in the driver instead.

### The state

```csharp
namespace XLibur.Excel.IO.WorksheetElements;

/// <summary>
/// What one worksheet element writer hands to a later one.
/// </summary>
/// <remarks>
/// <c>&lt;sheetFormatPr&gt;</c> (slot 4) computes the default column width that <c>&lt;cols&gt;</c>
/// (slot 5) needs, which before spec 31 was an <c>out double</c> on one signature and a plain
/// <c>double</c> parameter on the next, threaded through a local in <c>GetWorksheetDom</c>. This is
/// the same shape spec 24 found on the load side, where <c>sheetPr</c> produces the
/// <c>PageSetupProperties</c> that <c>pageSetup</c> consumes, and it is carried the same way.
/// </remarks>
internal struct WorksheetWriteState
{
    /// <summary>Set by slot 4, read by slot 5. Zero until slot 4 has run.</summary>
    internal double WorksheetColumnWidth;
}
```

`maxOutlineColumn` and `maxOutlineRow` (`WorksheetPartWriter.cs:170-176`) do **not** go in the state.
They are pure functions of `xlWorksheet` computed by the driver for one consumer; the slot-4 writer
computes them itself and the driver loses six lines.

### The driver

```csharp
    /// <summary>
    /// The children of <c>&lt;worksheet&gt;</c> XLibur writes, in ECMA-376 order. The order of this
    /// list is the schema order and nothing else; it is asserted ascending and duplicate-free by
    /// <c>WorksheetSlotOrderTests</c>. Adding an element means adding one entry here and one class,
    /// not editing a 90-line call sequence.
    /// </summary>
    private static readonly IXLWorksheetElementWriter[] ElementWriters =
    [
        new SheetPropertiesWriter(),          // 1
        new SheetDimensionWriter(),           // 2
        new SheetViewsWriter(),               // 3
        new SheetFormatPropertiesWriter(),    // 4  -> state.WorksheetColumnWidth
        new ColumnsWriter(),                  // 5  <- state.WorksheetColumnWidth
        new SheetDataPlaceholderWriter(),     // 6
        new SheetProtectionElementWriter(),   // 8
        new AutoFilterElementWriter(),        // 11
        new MergeCellsWriter(),               // 15
        new ConditionalFormattingElementWriter(), // 17
        new DataValidationsElementWriter(),   // 18
        new HyperlinksWriter(),               // 19
        new PrintOptionsWriter(),             // 20
        new PageMarginsWriter(),              // 21
        new PageSetupElementWriter(),         // 22
        new HeaderFooterWriter(),             // 23
        new RowBreaksWriter(),                // 24
        new ColumnBreaksWriter(),             // 25
        new TablePartsWriter(),               // 39
        new WorksheetExtensionListWriter(),   // 40
    ];
```

and, replacing `WorksheetPartWriter.cs:164-223`:

```csharp
        var cm = new XLWorksheetContentManager(worksheet);
        var ctx = new WorksheetWriteContext
        {
            Worksheet = worksheet,
            Cm = cm,
            XlWorksheet = xlWorksheet,
            Part = worksheetPart,
            Save = context,
            Options = options,
        };
        var state = default(WorksheetWriteState);

        foreach (var writer in ElementWriters)
            writer.Write(in ctx, ref state);

        // Slots 30/31/32 are not in the list. The drawing family writes DrawingsPart content as
        // well as the <worksheet> child that references it, so it is part-level work wearing an
        // element-level signature, and specs 15/16/17 are all inside PictureWriter.cs. It stays an
        // explicit tail call until one of those lands. See spec 31, Non-goals.
        PictureWriter.WriteDrawings(worksheet, cm, xlWorksheet, worksheetPart, context);
        ChartWriter.WriteCharts(worksheet, cm, xlWorksheet, worksheetPart, context);
        PictureWriter.WriteLegacyDrawing(worksheet, cm, xlWorksheet);
        HeaderFooterImageWriter.WriteHeaderFooterImages(worksheet, cm, xlWorksheet, worksheetPart, context);

        return worksheet;
```

The four tail calls keep running after slot 39, which preserves the invariant
`PictureWriter.cs:40` depends on.

### Slot 40 gets an owner

`WorksheetExtensionListWriter` is the only module that creates or removes `<extLst>`. The three
extension producers keep their content-building code and lose their slot-management code, becoming
methods the owner calls:

```csharp
internal sealed class WorksheetExtensionListWriter : IXLWorksheetElementWriter
{
    public XLWorksheetContents Slot => XLWorksheetContents.WorksheetExtensionList;

    /// <summary>
    /// The single owner of <c>&lt;extLst&gt;</c>. Before spec 31, three modules created it and two
    /// of them deleted it, each carrying its own copy of the "remove only when childless" rule.
    /// Nothing had failed yet, but two copies had already drifted: the URI comparison was
    /// InvariantCultureIgnoreCase in one and OrdinalIgnoreCase in the other, and the data-bar
    /// producer never removed a stale extension at all.
    /// </summary>
    public void Write(in WorksheetWriteContext ctx, ref WorksheetWriteState state)
    {
        var wantsDataBars = ConditionalFormattingWriter.HasExtensionDataBars(ctx.XlWorksheet);
        var wantsSparklines = ctx.XlWorksheet.SparklineGroups.Any();
        var wantsDataValidations = DataValidationWriter.HasExtensionDataValidations(ctx.XlWorksheet, ctx.Options);

        if (!wantsDataBars && !wantsSparklines && !wantsDataValidations)
        {
            RemoveIfEmptyAfterPruning(in ctx);
            return;
        }

        var extLst = ctx.EnsureElement<WorksheetExtensionList>(Slot);

        // Child <ext> order is unconstrained by the schema, but it is observable output, so it is
        // pinned to the pre-spec-31 order: data bars, sparklines, data validations.
        ConditionalFormattingWriter.WriteDataBarExtension(extLst, in ctx, wantsDataBars);
        ConditionalFormattingWriter.WriteSparklineExtension(extLst, in ctx, wantsSparklines);
        DataValidationWriter.WriteDataValidationExtension(extLst, in ctx, wantsDataValidations);

        if (!extLst.HasChildren)
            ctx.RemoveElement<WorksheetExtensionList>(Slot);
    }
}
```

The exact split of "what the producer decides" versus "what the owner decides" is task 8's job to
settle against the byte-identity gate. The invariant that must come out of it: **`EnsureElement` and
`RemoveElement` for slot 40 appear in exactly one file.**

### Where the risk actually is

Moving the extLst work from position `:205-207` to slot-40 position changes *when* the anchor walk
runs, and therefore what `GetPreviousElementFor(40)` can see. Today the extLst is created while slots
30/31/32/39 are still unset, so it anchors on slot 17 or 18 and later insertions (`TableParts` at
`:216`, `<drawing>` via `InsertBefore(_, tableParts)` at `:218`) end up threading themselves in front
of it. Under this spec the extLst is inserted last and anchors on slot 39, which reaches the same
final child order by a different route.

Same result, different mechanism. **That is exactly the kind of change a reasoning argument gets
wrong and a byte comparison catches**, which is why task 0 comes first and why it is not enough for
it to pass — it has to be shown to fail.

## Global constraints

- Warnings are errors (`TreatWarningsAsErrors=true`); nullable is enabled. New code must be
  null-annotated.
- Branch per spec; never commit to main. Commit prefixes `refactor:` / `fix:` / `test:` / `perf:`.
- No compound shell commands (`&&`, `||`, `;`) in agent tool calls.
- **Do not use `sed -i` on tracked files.** `.gitattributes` checks source out as CRLF and Git Bash's
  `sed -i` rewrites the file as LF, turning a one-line change into a whole-file diff. Use the
  Edit/Write tools; verify with `git diff --numstat` — a file whose changed-line count approaches its
  total line count was rewritten, not edited.
- Test filtering uses `--treenode-filter`, never `--filter`. Exit 5 = invalid option;
  exit 8 = zero tests matched. Never filter at solution level — name the `.csproj`.
- Pass `-f net10.0` for iteration; run without it before opening the PR.
- Build: `dotnet build XLibur/XLibur.csproj -c Release -v q`
- Tests: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
- Tests are TUnit and assertions are awaitable: `await Assert.That(actual).IsEqualTo(expected)`.
  A missing `await` means the assertion never runs and the test passes regardless. `[Test]`,
  `[Arguments(...)]`, `[MethodDataSource(...)]`. The suite runs serially.

## Work plan

| # | Task | Size | Gate |
|---|---|---|---|
| 0 | Golden byte-identity corpus for worksheet XML — **and prove it can fail** | M | Deliberate break fails the corpus |
| 1 | Characterization: pin the slot-30 and slot-40 contention as it is today | S | New tests green on unmodified code |
| 2 | Interface, context, state, slot helpers; convert slot 8 as the pattern | M | Corpus byte-identical; suite green |
| 3 | Slots 1–5 — the SDK-typed family and `<cols>`; `out double` becomes state | M | Corpus byte-identical |
| 4 | Slots 6, 11, 15 — SheetData placeholder, AutoFilter, MergeCells | S | Corpus byte-identical |
| 5 | Slots 17, 18 — conditional formatting and data validations, extLst left behind | M | Corpus byte-identical |
| 6 | Slots 19–25 — the seven `PageSetupWriter` entry points | M | Corpus byte-identical |
| 7 | Slot 39 — TableParts; both private methods leave `WorksheetPartWriter` | S | Corpus byte-identical |
| 8 | Slot 40 — `WorksheetExtensionListWriter` takes sole ownership of `<extLst>` | M | Corpus byte-identical; task 1 tests green |
| 9 | Collapse the driver; assert one owner per slot, ascending | S | Corpus byte-identical; new order test green |
| 10 | Confirm per-sheet save cost is unchanged | S | Within noise of baseline |

Tasks 3–7 are mechanical and each one is independently revertable. Task 8 is the one that changes
behaviour-adjacent structure; task 0 exists to size it.

---

### Task 0 — Golden byte-identity baseline

Every task after this one claims to change no output. Without a byte comparison that claim is an
assertion. Spec 22's task 0 established the pattern and the rule: **a refactor gated by a test that
cannot fail is not gated.** Spec 22's harness is chart-part-specific and is not in the tree at
`1b41cadd` (`grep -rn "CaptureChartPartXml" XLibur.Tests` → nothing), so this builds a worksheet-part
one.

**Files:**
- Create: `XLibur.Tests/Excel/IO/WorksheetGoldenCorpus.cs`
- Create: `XLibur.Tests/Excel/IO/WorksheetGoldenCorpusTests.cs`
- Create: `XLibur.Tests/Resource/Other/Worksheets/Golden/` (committed `.xml` fixtures)

**Interfaces:**
- Produces: `WorksheetGoldenCorpus.CaptureSheetXml(Action<IXLWorksheet>) → string` and
  `WorksheetGoldenCorpus.CaptureSheetXmlFromFile(string, Action<IXLWorksheet>) → string`, the gates
  for tasks 2–9.

- [ ] **Step 1: Write the capture helper**

```csharp
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;

namespace XLibur.Tests.Excel.IO;

internal static class WorksheetGoldenCorpus
{
    /// <summary>
    /// Builds a workbook, saves it, and returns the raw bytes of the first worksheet part as text.
    /// Byte-identity of this string across a refactor is the gate every task in spec 31 is measured
    /// with.
    /// </summary>
    internal static string CaptureSheetXml(Action<IXLWorksheet> build)
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            build(ws);
            wb.SaveAs(ms);
        }

        return ReadFirstSheet(ms);
    }

    /// <summary>
    /// The load-then-save path, which is the one that matters for the 17 pass-through slots: they
    /// exist only in a file XLibur did not author, and the anchor walk has to place the rewritten
    /// elements correctly around them.
    /// </summary>
    internal static string CaptureSheetXmlFromFile(string resourcePath, Action<IXLWorksheet> edit)
    {
        using var source = TestHelper.GetStreamFromResource(resourcePath);
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook(source))
        {
            edit(wb.Worksheet(1));
            wb.SaveAs(ms);
        }

        return ReadFirstSheet(ms);
    }

    private static string ReadFirstSheet(MemoryStream ms)
    {
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        var part = doc.WorkbookPart!.WorksheetParts.First();
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
```

`TestHelper.GetStreamFromResource` is written from the shape the suite already uses to open resource
workbooks. Find the actual helper — `XLibur.Tests/Excel/RoundTripFidelityTests.cs` opens resources
already — and use that form rather than adding a second one.

- [ ] **Step 2: Write the corpus over every slot this spec converts**

Twenty fixtures for the twenty in-scope slots, plus at least two loaded-file fixtures for the
pass-through slots. Grouping several slots into one fixture is fine and desirable — a fixture with
`<autoFilter>`, `<mergeCells>` and `<pageSetup>` in it proves their *relative* order as well as their
content, which a one-slot-per-fixture corpus would not.

```csharp
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// Pins the exact worksheet-part XML XLibur writes. Spec 31 reorganises every module that produces
/// this XML without changing it; a diff here is a finding to investigate, never noise to
/// re-baseline without a written explanation.
/// </summary>
public class WorksheetGoldenCorpusTests
{
    private const string GoldenDir = "Resource/Other/Worksheets/Golden";

    [Test]
    [Arguments("empty")]                 // slots 1-6 on a bare sheet
    [Arguments("format-and-columns")]    // 4 -> 5, the worksheetColumnWidth handoff
    [Arguments("frozen-panes")]          // 3, incl. <pane> and <selection>
    [Arguments("protection")]            // 8
    [Arguments("autofilter")]            // 11
    [Arguments("merged")]                // 15
    [Arguments("conditional-formats")]   // 17
    [Arguments("databar")]               // 17 + 40
    [Arguments("sparklines")]            // 40
    [Arguments("data-validation")]       // 18
    [Arguments("data-validation-x14")]   // 18 + 40, cross-sheet reference
    [Arguments("extlst-all-three")]      // 40 with all three producers at once
    [Arguments("hyperlinks")]            // 19
    [Arguments("print-setup")]           // 20-25 together
    [Arguments("breaks")]                // 24, 25
    [Arguments("tables")]                // 39
    [Arguments("tables-and-picture")]    // 39 then 30 — the InsertBefore(tableParts) dependency
    [Arguments("everything")]            // all twenty in-scope slots on one sheet
    public async Task Sheet_xml_matches_the_golden_fixture(string name)
    {
        var actual = WorksheetGoldenCorpus.CaptureSheetXml(ws => BuildFixture(name, ws));
        await AssertGolden(name, actual);
    }

    [Test]
    [Arguments(@"TryToLoad\LO\xlsx\activex_checkbox.xlsx", "loaded-controls")]
    public async Task Loaded_sheet_xml_matches_the_golden_fixture(string resource, string name)
    {
        // <controls> is slot 36 and XLibur has no model for it, so it only survives because nothing
        // rewrites it. Its position relative to the elements that ARE rewritten is decided by the
        // anchor walk, which is what spec 31 changes the timing of.
        var actual = WorksheetGoldenCorpus.CaptureSheetXmlFromFile(
            resource, ws => ws.Cell("Z99").Value = "touched");
        await AssertGolden(name, actual);
    }

    private static async Task AssertGolden(string name, string actual)
    {
        var path = Path.Combine(GoldenDir, name + ".xml");
        if (!File.Exists(path))
        {
            Directory.CreateDirectory(GoldenDir);
            File.WriteAllText(path, actual);
        }

        await Assert.That(actual).IsEqualTo(File.ReadAllText(path));
    }

    private static void BuildFixture(string name, IXLWorksheet ws) { /* … */ }
}
```

`RoundTripFidelityTests.Form_control_references_survive_in_the_worksheet_xml` already asserts that
`TryToLoad\LO\xlsx\activex_checkbox.xlsx` round-trips `controls>`, `CheckBox1343` and
`legacyDrawing`, so that resource is known to carry pass-through content. Find one or two more with
`grep -rl "phoneticPr\|ignoredErrors\|customSheetViews" XLibur.Tests/Resource` and add them.

- [ ] **Step 3: Run it twice — once to write the fixtures, once to assert them**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WorksheetGoldenCorpusTests/*"`
Expected: PASS on the first run (fixtures written), PASS on the second (fixtures asserted).

Inspect the committed fixtures by eye before trusting them. A fixture that is one `<worksheet/>` tag
long is a fixture whose builder did nothing.

- [ ] **Step 4: Prove the gate bites — three separate breaks**

One break is not enough here, because the failure modes this spec risks are different in kind. Do all
three, one at a time, reverting between:

1. **Content.** In `PageSetupWriter.WritePageMargins` (`:85`), multiply the top margin by 2.
   Expected: FAIL on `print-setup` and `everything`.
2. **Order.** In `WorksheetPartWriter.GetWorksheetDom`, move the `WriteMergeCells` call from `:203`
   to immediately after `:214`. Expected: FAIL on `merged` and `everything`, on element order.
3. **Slot management.** In `DataValidationWriter.RemoveExtensionDataValidations` (`:151`), delete the
   `if (!worksheetExtensionList.HasChildren)` guard at `:167`.
   Expected: FAIL on `data-validation-x14` or `extlst-all-three`.

If break 2 does not fail, the corpus is not sensitive to order and is useless for this spec — widen
it before continuing. If break 3 does not fail, **the extLst corpus does not reach the code task 8
rewrites**, which is the highest-risk task in the plan; widen it before continuing.

Restore all three.

- [ ] **Step 5: Commit**

```bash
git add XLibur.Tests/Excel/IO/WorksheetGoldenCorpus.cs XLibur.Tests/Excel/IO/WorksheetGoldenCorpusTests.cs XLibur.Tests/Resource/Other/Worksheets/Golden
git commit -m 'test(io): pin worksheet-part XML with a golden corpus (spec 31 task 0)'
```

---

### Task 1 — Pin the two contended slots

Two claims in this spec's evidence are premises about current behaviour, and both could be wrong.
This task tests them. **A disproved premise here is a real result and changes the plan.**

**Files:**
- Create: `XLibur.Tests/Excel/IO/WorksheetSlotContentionTests.cs`

**Interfaces:**
- Produces: `The_drawing_writer_requires_tableParts_to_exist_first`,
  `The_extension_list_survives_every_combination_of_its_three_producers`.

- [ ] **Step 1: Pin the slot-30 ordering dependency**

```csharp
/// <summary>
/// PictureWriter.cs:40 reads worksheet.Elements&lt;TableParts&gt;().First() with no guard. It works
/// only because PopulateTablePartReferences runs two lines earlier in GetWorksheetDom and always
/// leaves a &lt;tableParts&gt; behind. This test does not reach that line directly — it saves a
/// sheet that has a picture and no tables, which is the case where the guarantee is doing work.
/// If spec 31 ever moves the drawing calls above slot 39, this is what fails.
/// </summary>
[Test]
public async Task A_picture_on_a_sheet_with_no_tables_saves_without_throwing()
{
    using var ms = new MemoryStream();
    using var wb = new XLWorkbook();
    var ws = wb.AddWorksheet("Sheet1");
    ws.Cell("A1").Value = "x";
    ws.AddPicture(TestHelper.OpenPngResource()).MoveTo(ws.Cell("C3"));

    wb.SaveAs(ms);

    ms.Position = 0;
    using var reopened = new XLWorkbook(ms);
    await Assert.That(reopened.Worksheet("Sheet1").Pictures.Count).IsEqualTo(1);
}
```

Find the suite's existing image-resource helper rather than adding one — `XLibur.Tests/Excel/`
already has picture tests that open a PNG.

- [ ] **Step 2: Pin every combination of the three extLst producers**

Eight combinations of {data bars, sparklines, x14 data validations}. For each: save, reopen, assert
all three features that were asked for came back, and assert the ones that were not asked for are
absent.

```csharp
/// <summary>
/// Three modules create &lt;extLst&gt; and two of them delete it, each carrying its own copy of the
/// "remove only when childless" rule. Spec 31 task 8 replaces those five entry points with one
/// owner. These eight rows are what proves the owner behaves identically.
/// </summary>
[Test]
[Arguments(false, false, false)]
[Arguments(true,  false, false)]
[Arguments(false, true,  false)]
[Arguments(false, false, true)]
[Arguments(true,  true,  false)]
[Arguments(true,  false, true)]
[Arguments(false, true,  true)]
[Arguments(true,  true,  true)]
public async Task The_extension_list_survives_every_combination(
    bool dataBars, bool sparklines, bool x14Validations)
{
    // build -> save -> reopen -> assert each requested feature present, each other absent
}
```

- [ ] **Step 3: Test the asymmetry — does a deleted data bar leave a stale x14 extension?**

`WriteExtensionDataBars` (`ConditionalFormattingWriter.cs:100-163`) has no `else` branch. The
sparkline and data-validation writers both have `Remove*` counterparts; this one does not.

```csharp
/// <summary>
/// Load a file with a data-bar conditional format, delete the format, save, and read the raw XML.
/// If x14:conditionalFormattings survives, that is a live round-trip defect, not just an asymmetry.
/// </summary>
[Test]
public async Task Deleting_a_data_bar_removes_its_extension()
{
    // …
    await Assert.That(sheetXml).DoesNotContain("conditionalFormattings");
}
```

**If this test fails, it has found a defect.** Do not fix it under task 1. Record it in this spec's
Results section, replace the assertion with the current behaviour plus a comment naming the gap, and
make removing the stale extension an explicit acceptance criterion of task 8 — the slot-40 owner is
the natural place for it, and it becomes the second thing this spec delivers beyond the refactor.

- [ ] **Step 4: Run**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WorksheetSlotContentionTests/*"`
Expected: PASS on steps 1 and 2. Step 3 is a genuine question; either answer is a result.

- [ ] **Step 5: Commit**

```bash
git add XLibur.Tests/Excel/IO/WorksheetSlotContentionTests.cs
git commit -m 'test(io): pin the drawing and extension-list slot contention (spec 31 task 1)'
```

---

### Task 2 — The interface, the context, the state, and one writer

**Files:**
- Create: `XLibur/Excel/IO/WorksheetElements/IXLWorksheetElementWriter.cs`
- Create: `XLibur/Excel/IO/WorksheetElements/WorksheetWriteContext.cs`
- Create: `XLibur/Excel/IO/WorksheetElements/WorksheetWriteState.cs`
- Create: `XLibur/Excel/IO/WorksheetElements/SheetProtectionElementWriter.cs`
- Delete: `XLibur/Excel/IO/SheetProtectionWriter.cs`
- Modify: `XLibur/Excel/IO/WorksheetPartWriter.cs:164-223`

**Interfaces:**
- Produces: `IXLWorksheetElementWriter`, `WorksheetWriteContext` (with `EnsureElement<T>` and
  `RemoveElement<T>`), `WorksheetWriteState`, and the `ElementWriters` array on
  `WorksheetPartWriter`.

Slot 8 goes first because `SheetProtectionWriter` is the cleanest instance of the ceremony
(`:17-24`), has exactly one entry point, and its `else` branch (`:99-100`) is the exact shape
`RemoveElement<T>` replaces.

- [ ] **Step 1: Create the three infrastructure files**

Verbatim from The design, above.

- [ ] **Step 2: Convert slot 8**

```csharp
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using XLibur.Utils;

namespace XLibur.Excel.IO.WorksheetElements;

internal sealed class SheetProtectionElementWriter : IXLWorksheetElementWriter
{
    public XLWorksheetContents Slot => XLWorksheetContents.SheetProtection;

    public void Write(in WorksheetWriteContext ctx, ref WorksheetWriteState state)
    {
        var protection = ctx.XlWorksheet.Protection;
        if (!protection.IsProtected)
        {
            ctx.RemoveElement<SheetProtection>(Slot);
            return;
        }

        var sheetProtection = ctx.EnsureElement<SheetProtection>(Slot);

        // Body moved verbatim from SheetProtectionWriter.cs:27-95.
        sheetProtection.Sheet = OpenXmlHelper.GetBooleanValue(protection.IsProtected, false);
        // …
    }
}
```

The body from `SheetProtectionWriter.cs:27-95` moves **unchanged**. Do not tidy it, do not reorder
the assignments — the byte-identity gate is on attribute emission order.

- [ ] **Step 3: Introduce the list in the driver, with one entry**

```csharp
    private static readonly IXLWorksheetElementWriter[] ElementWriters =
    [
        new SheetProtectionElementWriter(), // 8
    ];
```

and replace `WorksheetPartWriter.cs:200` with the loop, positioned exactly where the slot-8 call was.
Every other call at `:166-221` stays put for now. The loop runs where the list's lowest slot used to
run, and each later task moves calls into the list from both sides of it.

- [ ] **Step 4: Build, then run the corpus, then the suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WorksheetGoldenCorpusTests/*"`
Expected: PASS, byte-identical, no fixture rewritten. If a fixture file changed on disk, `git diff`
it — that is a real behaviour change and it must be explained before continuing.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/IO/WorksheetElements XLibur/Excel/IO/WorksheetPartWriter.cs XLibur/Excel/IO/SheetProtectionWriter.cs
git commit -m 'refactor(io): give worksheet element writers one interface (spec 31 task 2)'
```

---

### Task 3 — Slots 1–5

The SDK-typed family plus `<cols>`. This is the task that kills the `out double`.

**Files:**
- Create: `SheetPropertiesWriter.cs`, `SheetDimensionWriter.cs`, `SheetViewsWriter.cs`,
  `SheetFormatPropertiesWriter.cs`, `ColumnsWriter.cs` under `WorksheetElements/`
- Delete: `XLibur/Excel/IO/SheetViewWriter.cs`
- Modify: `XLibur/Excel/IO/ColumnWriter.cs` — `WriteColumns` (`:20`) moves out
- Modify: `XLibur/Excel/IO/WorksheetPartWriter.cs` — delete `:166-181`, and `:170-176`

- [ ] **Step 1: Move the four `SheetViewWriter` entry points**

`:15`, `:42`, `:62`, `:244` become four classes. The eight private helpers (`:100 :113 :149 :162
:179 :188 :225`) travel with whichever entry point calls them —
`SetBooleanViewProperties`/`SetupPane`/`GetActivePaneValue`/`GetActivePaneForActiveCell`/
`SetTopLeftCell`/`SetupSelections`/`SetZoomScales` all belong to slot 3.

None of the four uses `EnsureElement`: they use the SDK's typed properties
(`worksheet.SheetProperties ??= new SheetProperties()`), and they must keep doing so. Changing them
to `EnsureElement` would change where the SDK places the element, which changes the bytes. They still
call `ctx.Cm.SetElement` exactly as today.

- [ ] **Step 2: Fold `maxOutlineColumn` / `maxOutlineRow` into slot 4**

`WorksheetPartWriter.cs:170-176`:

```csharp
        var maxOutlineColumn = 0;
        if (xlWorksheet.ColumnCount() > 0)
            maxOutlineColumn = xlWorksheet.GetMaxColumnOutline();

        var maxOutlineRow = 0;
        if (xlWorksheet.RowCount() > 0)
            maxOutlineRow = xlWorksheet.GetMaxRowOutline();
```

moves into `SheetFormatPropertiesWriter.Write` verbatim, guards included. Both are pure functions of
`xlWorksheet` with one consumer, so they never belonged in the driver.

- [ ] **Step 3: Turn the `out double` into state**

Slot 4 ends with `state.WorksheetColumnWidth = …` in place of the `out` assignment. Slot 5 opens with
`var worksheetColumnWidth = state.WorksheetColumnWidth;` and is otherwise `ColumnWriter.cs:26-61`
unchanged.

The ordering dependency is now expressed twice — by the writers' `Slot` values and by task 9's
ascending-order test — instead of by two adjacent lines in a 91-line method.

- [ ] **Step 4: Add all five to the list, in slot order, and delete `:166-181`**

The loop now starts the driver. `ColumnWriter.cs` keeps `GetColumnWidth` (`:185`) and the private
helpers `WriteColumns` calls; check with `grep -rn "ColumnWriter\." XLibur --include=*.cs` whether
anything outside the worksheet path calls into it before deciding how much of the file survives.

- [ ] **Step 5: Build, corpus, suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WorksheetGoldenCorpusTests/*"`
Expected: PASS, byte-identical. `format-and-columns` is the fixture that proves the state handoff.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS.

- [ ] **Step 6: Verify no whole-file rewrites**

Run: `git diff --numstat`
Expected: no modified file whose changed-line count approaches its total line count. If one does, it
was rewritten with LF endings — see Global constraints.

- [ ] **Step 7: Commit**

```bash
git add XLibur/Excel/IO/WorksheetElements XLibur/Excel/IO/SheetViewWriter.cs XLibur/Excel/IO/ColumnWriter.cs XLibur/Excel/IO/WorksheetPartWriter.cs
git commit -m 'refactor(io): move worksheet slots 1-5 behind the element writer interface (spec 31 task 3)'
```

---

### Task 4 — Slots 6, 11, 15

**Files:**
- Create: `SheetDataPlaceholderWriter.cs`, `AutoFilterElementWriter.cs`, `MergeCellsWriter.cs`
- Modify: `XLibur/Excel/IO/AutoFilterWriter.cs` — `WriteAutoFilter` (`:15`) moves out
- Modify: `XLibur/Excel/IO/WorksheetPartWriter.cs` — delete `:183-203`

- [ ] **Step 1: Slot 6 — the inline ceremony becomes a writer**

`WorksheetPartWriter.cs:185-192` is the ceremony written out longhand for `SheetData`, and it is the
only slot with no module at all:

```csharp
internal sealed class SheetDataPlaceholderWriter : IXLWorksheetElementWriter
{
    public XLWorksheetContents Slot => XLWorksheetContents.SheetData;

    /// <summary>
    /// Emits an empty <c>&lt;sheetData&gt;</c> as an anchor and nothing else. Rows and cells are
    /// streamed straight to the part by <see cref="SheetDataWriter"/> in
    /// <c>WorksheetPartWriter.StreamToPart</c>, which matches on this element to know where to
    /// splice them in — building the DOM for them here is what the empty-SheetData substitution in
    /// <c>ReadWorksheetDomWithoutSheetData</c> exists to avoid.
    /// </summary>
    public void Write(in WorksheetWriteContext ctx, ref WorksheetWriteState state)
        => ctx.EnsureElement<SheetData>(Slot);
}
```

Keep the explanatory comment from `WorksheetPartWriter.cs:194-196` — move it onto the new class.

- [ ] **Step 2: Slot 11 — `AutoFilterWriter.WriteAutoFilter`**

`AutoFilterWriter.cs:15-36`. `PopulateAutoFilter` (`:38`) and the nine private criteria builders
(`:64` onward) stay in `AutoFilterWriter.cs`; confirm with
`grep -rn "AutoFilterWriter\." XLibur --include=*.cs` that `PopulateAutoFilter` has an external
caller before leaving it behind. If it does not, it moves too and the file goes.

- [ ] **Step 3: Slot 15 — `WriteMergeCells` leaves `WorksheetPartWriter`**

`WorksheetPartWriter.cs:226-252` moves verbatim, with `:232-237` collapsing to
`ctx.EnsureElement<MergeCells>(Slot)` and `:249-250` to `ctx.RemoveElement<MergeCells>(Slot)`.

- [ ] **Step 4: Build, corpus, suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WorksheetGoldenCorpusTests/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS, byte-identical.

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/IO/WorksheetElements XLibur/Excel/IO/AutoFilterWriter.cs XLibur/Excel/IO/WorksheetPartWriter.cs
git commit -m 'refactor(io): move worksheet slots 6, 11 and 15 behind the interface (spec 31 task 4)'
```

---

### Task 5 — Slots 17 and 18, without their extension halves

**Files:**
- Create: `ConditionalFormattingElementWriter.cs`, `DataValidationsElementWriter.cs`
- Modify: `XLibur/Excel/IO/ConditionalFormattingWriter.cs`
- Modify: `XLibur/Excel/IO/DataValidationWriter.cs`
- Modify: `XLibur/Excel/IO/WorksheetPartWriter.cs` — delete `:205-207`

This task splits each of the two modules along the seam between "writes my own slot" and "writes
into slot 40". The slot-40 half stays where it is for now; task 8 collects it.

- [ ] **Step 1: Slot 17**

`ConditionalFormattingWriter.WriteConditionalFormatting` (`:18-98`) becomes the slot-17 writer,
**minus its last line** — `WriteExtensionDataBars(worksheet, cm, xlWorksheet, context);` at `:97`.
The tail call is what makes one method write two slots.

Keep the fast-path comment at `:24-33` and the priority-renumbering comment at `:46-48` intact; both
record measured decisions.

The `:37-38` and `:58-59` pairs become `ctx.RemoveElement<ConditionalFormatting>(Slot)`. Note the
`else` branch at `:63-94` does *not* use `EnsureElement` — it re-anchors on each iteration
(`previousElement = conditionalFormatting;` at `:92`) because a worksheet can have many
`<conditionalFormatting>` children in one slot. **Leave that loop exactly as it is.** It is the one
slot with a one-to-many relationship, and `EnsureElement` is single-element by contract.

- [ ] **Step 2: Slot 18**

`DataValidationWriter.WriteDataValidations` (`:16-62`) becomes the slot-18 writer, minus
`WriteExtensionDataValidations(…)` at `:61`. `WriteStandardDataValidations` (`:84-133`) and
`UsesExternalSheet` (`:64-82`) travel with it.

The classification loop at `:33-58` decides which validations are standard and which need the x14
extension. That decision is needed by both slots, so it becomes an internal method the slot-40 owner
can also call — this is the `DataValidationWriter.HasExtensionDataValidations` referenced in the
design sketch. **Do not duplicate the classification.** Duplicating it is exactly the defect shape
this spec removes.

- [ ] **Step 3: Bridge until task 8**

Slots 17 and 18 join the list. Their two extension tail calls stay in `GetWorksheetDom` as two
explicit lines, running immediately after the loop, with a `// spec 31 task 8` comment. That is
temporarily *worse* than the code being replaced — three lines where there were two — and it is
correct: it makes the seam visible and it keeps every intermediate commit green.

- [ ] **Step 4: Build, corpus, suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WorksheetGoldenCorpusTests/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WorksheetSlotContentionTests/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS, byte-identical.

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/IO/WorksheetElements XLibur/Excel/IO/ConditionalFormattingWriter.cs XLibur/Excel/IO/DataValidationWriter.cs XLibur/Excel/IO/WorksheetPartWriter.cs
git commit -m 'refactor(io): split slots 17 and 18 from their extension-list halves (spec 31 task 5)'
```

---

### Task 6 — Slots 19–25

The seven `PageSetupWriter` entry points, which have nothing in common but a filename.

**Files:**
- Create: `HyperlinksWriter.cs` (19), `PrintOptionsWriter.cs` (20), `PageMarginsWriter.cs` (21),
  `PageSetupElementWriter.cs` (22), `HeaderFooterWriter.cs` (23), `RowBreaksWriter.cs` (24),
  `ColumnBreaksWriter.cs` (25)
- Delete: `XLibur/Excel/IO/PageSetupWriter.cs`
- Modify: `XLibur/Excel/IO/WorksheetPartWriter.cs` — delete `:208-214`

- [ ] **Step 1: Split the file seven ways**

| From | Lines | To | Slot |
|---|---|---|---|
| `WriteHyperlinks` | `:14-63` | `HyperlinksWriter` | 19 |
| `WritePrintOptions` | `:65-83` | `PrintOptionsWriter` | 20 |
| `WritePageMargins` | `:85-104` | `PageMarginsWriter` | 21 |
| `WritePageSetup` + `SetPageSetupBasicProperties` + `SetPageSetupDpiAndScale` | `:106-179` | `PageSetupElementWriter` | 22 |
| `WriteHeaderFooter` | `:181-221` | `HeaderFooterWriter` | 23 |
| `WriteRowBreaks` | `:223-275` | `RowBreaksWriter` | 24 |
| `WriteColumnBreaks` | `:277-326` | `ColumnBreaksWriter` | 25 |

Slot 19 is the only one of the seven that needs `ctx.Part` and `ctx.Save`; the other six need only
`ctx.Worksheet`, `ctx.Cm` and `ctx.XlWorksheet`. The context struct absorbs that difference, which is
the point — six writers stop being able to reach a `WorksheetPart` they never used, and one keeps it.

- [ ] **Step 2: Replace each ceremony with `EnsureElement` / `RemoveElement`**

Seven of the 20 ceremony sites live in this file (`:32 :72 :92 :113 :192 :233 :287`). After this task
the count is down to the three in `WorksheetPartWriter.cs`, the two in the extension halves, and
`PictureWriter.cs:159`.

- [ ] **Step 3: Build, corpus, suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WorksheetGoldenCorpusTests/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS, byte-identical. `print-setup` and `breaks` are the fixtures that matter.

- [ ] **Step 4: Commit**

```bash
git add XLibur/Excel/IO/WorksheetElements XLibur/Excel/IO/PageSetupWriter.cs XLibur/Excel/IO/WorksheetPartWriter.cs
git commit -m 'refactor(io): split PageSetupWriter into its seven slots (spec 31 task 6)'
```

---

### Task 7 — Slot 39, and `WorksheetPartWriter` stops writing elements

**Files:**
- Create: `TablePartsWriter.cs`
- Modify: `XLibur/Excel/IO/WorksheetPartWriter.cs` — delete `:216` and `:254-283`

- [ ] **Step 1: Move `PopulateTablePartReferences`**

`WorksheetPartWriter.cs:254-283` moves verbatim. `:262-273` collapses to
`ctx.EnsureElement<TableParts>(Slot)`.

The `EmptyTableException` throw at `:257-259` stays — it is a precondition on the whole save, not
element writing, and moving it changes when it fires relative to the other writers. Keep it as the
first statement of `Write`.

- [ ] **Step 2: Confirm `WorksheetPartWriter` writes no element itself**

Run: `grep -nE "GetPreviousElementFor|SetElement|InsertAfter|RemoveAllChildren" XLibur/Excel/IO/WorksheetPartWriter.cs`
Expected: no output.

- [ ] **Step 3: Build, corpus, suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WorksheetGoldenCorpusTests/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS, byte-identical. `tables` and `tables-and-picture` are the fixtures that matter — the
second one is the `InsertBefore(_, tableParts)` dependency from `PictureWriter.cs:40`.

- [ ] **Step 4: Commit**

```bash
git add XLibur/Excel/IO/WorksheetElements XLibur/Excel/IO/WorksheetPartWriter.cs
git commit -m 'refactor(io): move table parts to a slot writer (spec 31 task 7)'
```

---

### Task 8 — Slot 40 gets one owner

The task the rest of the plan exists to make safe.

**Files:**
- Create: `WorksheetExtensionListWriter.cs`
- Modify: `XLibur/Excel/IO/ConditionalFormattingWriter.cs` — `:100-163`, `:165-203`, `:205-…`
- Modify: `XLibur/Excel/IO/DataValidationWriter.cs` — `:135-234`
- Modify: `XLibur/Excel/IO/WorksheetPartWriter.cs` — delete the two bridge lines from task 5 step 3

**Interfaces:**
- Produces: `WorksheetExtensionListWriter`, sole owner of `EnsureElement`/`RemoveElement` for
  `XLWorksheetContents.WorksheetExtensionList`.
- Produces: `ConditionalFormattingWriter.HasExtensionDataBars`,
  `ConditionalFormattingWriter.WriteDataBarExtension`,
  `ConditionalFormattingWriter.WriteSparklineExtension`,
  `DataValidationWriter.HasExtensionDataValidations`,
  `DataValidationWriter.WriteDataValidationExtension` — all taking the `<extLst>` element, never
  creating or removing it.

- [ ] **Step 1: Invert the three producers**

Each of the five entry points listed in "Slot 40 has three owners" loses its slot handling:

| Was | Becomes |
|---|---|
| `WriteExtensionDataBars` `:110-117` (create) | takes `WorksheetExtensionList` as a parameter |
| `WriteSparklineGroups` `:208-215` (create) | takes it as a parameter |
| `RemoveSparklineExtension` `:198-202` (remove) | prunes its own `<ext>`; leaves the extLst alone |
| `WriteExtensionDataValidationElements` `:180-187` (create) | takes it as a parameter |
| `RemoveExtensionDataValidations` `:167-171` (remove) | prunes its own `<ext>`; leaves the extLst alone |

The two `Remove*` methods still remove their own `<ext>` when it is childless — that is their own
element and their own business. What they stop doing is deciding the fate of the container.

- [ ] **Step 2: Settle the URI comparison**

`ConditionalFormattingWriter.cs:188` uses `InvariantCultureIgnoreCase`; `DataValidationWriter.cs:157`
uses `OrdinalIgnoreCase`. Both compare a hardcoded GUID-shaped URI against an attribute read from a
file. **`OrdinalIgnoreCase` is correct** — these are identifiers, not text, and culture-sensitive
comparison of an ASCII GUID is at best a no-op and at worst a Turkish-I hazard. Change the
conditional-formatting one and say so in the commit message.

If the corpus notices, something is reading a URI that differs by case in a culture-dependent way,
which would be a finding. It should not notice.

- [ ] **Step 3: Write the owner**

Per the sketch in The design. The container is created once if any producer wants it, each producer
fills its own `<ext>` in the pre-spec-31 order (data bars, sparklines, data validations), and the
owner removes the container if it ends up childless.

If task 1 step 3 found the stale-data-bar-extension defect, this is where it gets fixed: the owner
calls `WriteDataBarExtension(extLst, in ctx, wantsDataBars: false)` unconditionally, and that method
prunes rather than returning early.

- [ ] **Step 4: Delete the bridge lines and add slot 40 to the list**

`GetWorksheetDom` now contains the loop and the four drawing tail calls, and nothing else between
`:164` and `return worksheet`.

- [ ] **Step 5: The full gate — this is the task most likely to move a byte**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WorksheetGoldenCorpusTests/*"`
Expected: PASS, byte-identical. `databar`, `sparklines`, `data-validation-x14` and `extlst-all-three`
are the four fixtures this task is aimed at.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/WorksheetSlotContentionTests/*"`
Expected: PASS on all eight combinations.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS.

**If a fixture differs, read the diff before touching anything.** The two most likely differences are
`<extLst>`'s position among its siblings (the anchor now walks from slot 39 rather than from slot
17/18 — see "Where the risk actually is") and the order of `<ext>` children. Neither is acceptable to
re-baseline silently. Either restore the previous bytes or write down, in this spec's Results, what
changed and why the new output is at least as correct.

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel/IO/WorksheetElements XLibur/Excel/IO/ConditionalFormattingWriter.cs XLibur/Excel/IO/DataValidationWriter.cs XLibur/Excel/IO/WorksheetPartWriter.cs
git commit -m 'refactor(io): give the worksheet extension list one owner (spec 31 task 8)'
```

---

### Task 9 — Collapse the driver and assert the invariants

**Files:**
- Modify: `XLibur/Excel/IO/WorksheetPartWriter.cs`
- Create: `XLibur.Tests/Excel/IO/WorksheetSlotOrderTests.cs`

**Interfaces:**
- Produces: `Every_writer_owns_a_distinct_slot`, `The_writer_list_is_in_ascending_slot_order`.

- [ ] **Step 1: Make the list the only thing `GetWorksheetDom` knows**

Final shape of `:164` onward is in The design, above. Move the drawing tail-call comment in with it.

- [ ] **Step 2: Assert one owner per slot, ascending**

```csharp
/// <summary>
/// The order of WorksheetPartWriter.ElementWriters is the ECMA-376 child order for
/// &lt;worksheet&gt;, which is also stated by the XLWorksheetContents enum. Before spec 31 those two
/// statements were the enum and a 91-line call sequence, and nothing checked that they agreed.
/// These two tests are the check.
/// </summary>
public class WorksheetSlotOrderTests
{
    [Test]
    public async Task Every_writer_owns_a_distinct_slot()
    {
        var slots = WorksheetPartWriter.ElementWritersForTest.Select(w => w.Slot).ToList();
        await Assert.That(slots.Distinct().Count()).IsEqualTo(slots.Count);
    }

    [Test]
    public async Task The_writer_list_is_in_ascending_slot_order()
    {
        var slots = WorksheetPartWriter.ElementWritersForTest.Select(w => (int)w.Slot).ToList();
        await Assert.That(slots).IsEquivalentTo(slots.OrderBy(s => s).ToList());
    }

    [Test]
    public async Task Every_slot_a_writer_touches_is_the_slot_it_declares()
    {
        // Grep-shaped, not reflection-shaped: assert that no writer file names a
        // XLWorksheetContents value other than its own Slot. Reading the sources is legitimate here
        // — the invariant is about which slot a file may mention, which no runtime check can see.
    }
}
```

`ElementWritersForTest` is an `internal static IReadOnlyList<IXLWorksheetElementWriter>` exposing the
array; `XLibur.Tests` already has an `InternalsVisibleTo` grant.

The third test is optional and worth the effort: it is what stops a fifth extLst producer appearing.
If reading source files from a test is judged too clever, replace it with the grep gate in
Acceptance criteria 6 and run it in CI instead.

- [ ] **Step 3: Confirm the driver is clean**

Run: `grep -nE "[A-Za-z]+Writer\.[A-Za-z]+\(" XLibur/Excel/IO/WorksheetPartWriter.cs`
Expected: exactly four lines — `PictureWriter.WriteDrawings`, `ChartWriter.WriteCharts`,
`PictureWriter.WriteLegacyDrawing`, `HeaderFooterImageWriter.WriteHeaderFooterImages`, plus
`SheetDataWriter.StreamSheetData` in `StreamToPart` (`:328`), which is not element writing.

Run: `grep -rn "GetPreviousElementFor(XLWorksheetContents" XLibur --include=*.cs`
Expected: exactly one line, in `WorksheetWriteContext.EnsureElement`.

- [ ] **Step 4: Build and run everything, both frameworks**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Expected: PASS on net8.0 and net10.0.

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/IO/WorksheetPartWriter.cs XLibur.Tests/Excel/IO/WorksheetSlotOrderTests.cs
git commit -m 'refactor(io): the worksheet element order becomes data (spec 31 task 9)'
```

---

### Task 10 — Confirm the per-sheet save cost is unchanged

`GetWorksheetDom` runs once per worksheet on every save. This spec adds one virtual call and one
struct copy per element, and removes twenty duplicated `Elements<T>()` traversals in favour of a
generic helper doing the same traversals. Neither direction is obvious; measure it.

The content manager's own construction cost is already on record: its remarks
(`XLWorksheetContentManager.cs:52-62`) say the constructor used to run
`Elements<T>().LastOrDefault()` thirty-nine times and that this "made building the manager a
significant share of the per-worksheet save cost". That is the scale this task is checking against.

- [ ] **Step 1: Measure the merge-base**

```
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --filter '*CreateAndSave*'
```

and the template round-trip fixture from spec 18:

```
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- profile template
```

The template sweep (1 / 10 / 40 header-only sheets) is the one that isolates per-sheet cost;
`CreateAndSave` is dominated by cells and will not move either way.

- [ ] **Step 2: Measure the branch**

Same commands, same fixtures, same machine, same session.

- [ ] **Step 3: Compare the per-sheet slope**

`docs/specs/` records ~40% run-to-run timing variance on this machine, so a single pair of runs
proves nothing about time. **Allocation is the reliable signal**: the writers are stateless
singletons in a `static readonly` array, allocated once per process, and the context and state
structs do not escape. Per-sheet allocation should be identical to the byte.

If per-sheet allocation has risen, the cause is almost certainly the context struct being copied —
check that every call site passes `in`, not a bare argument, and that no writer stores it in a field.

**Decision rule.** A per-sheet allocation increase must be explained before this spec lands, not
after. A time regression beyond the noise floor spec 18 records must be reproduced across three runs
before it is treated as real. Record the numbers in a Results section either way.

- [ ] **Step 4: Commit the Results section**

```bash
git add docs/specs/31-worksheet-element-writers.md
git commit -m 'docs(specs): record the per-sheet save numbers for spec 31'
```

---

## Acceptance criteria

1. `grep -rn "GetPreviousElementFor(XLWorksheetContents" XLibur --include=*.cs` returns **exactly one
   line**, inside `WorksheetWriteContext.EnsureElement`. (Baseline: 20 lines across 9 files.)
2. `grep -rn "SetElement(XLWorksheetContents.WorksheetExtensionList" XLibur --include=*.cs` returns
   lines from **exactly one file**, `WorksheetExtensionListWriter.cs`. (Baseline: 4 lines across 2
   files.)
3. `grep -nE "[A-Za-z]+Writer\.[A-Za-z]+\(" XLibur/Excel/IO/WorksheetPartWriter.cs` returns at most
   **five** lines: the four deferred drawing calls and `SheetDataWriter.StreamSheetData`.
   (Baseline: 22.)
4. `grep -nE "GetPreviousElementFor|SetElement|InsertAfter|RemoveAllChildren" XLibur/Excel/IO/WorksheetPartWriter.cs`
   returns nothing. `WriteMergeCells` and `PopulateTablePartReferences` no longer exist in that file.
5. `XLibur/Excel/IO/WorksheetElements/` contains **20** types implementing
   `IXLWorksheetElementWriter`, and `WorksheetSlotOrderTests` proves their slots are distinct and
   ascending.
6. No writer file names a `XLWorksheetContents` value other than its own `Slot`. Gate:
   for each file `F` in `WorksheetElements/`, `grep -o "XLWorksheetContents\.[A-Za-z]+" F | sort -u`
   yields one value.
7. `SheetViewWriter.cs`, `PageSetupWriter.cs` and `SheetProtectionWriter.cs` no longer exist;
   `ColumnWriter.cs` and `AutoFilterWriter.cs` retain only members proven in tasks 3 and 4 to have
   callers outside the worksheet element path.
8. Every golden fixture in `XLibur.Tests/Resource/Other/Worksheets/Golden/` is byte-identical to its
   task 0 content. `git diff --stat` over that directory across the whole spec is empty — or every
   difference is explained in a Results section.
9. `WorksheetSlotContentionTests` passes all eight extLst combinations and the picture-without-tables
   case, before and after task 8.
10. Full suite green on net8.0 and net10.0.
11. Per-sheet save allocation identical to the pre-spec value; per-sheet time within spec 18's noise
    floor. Numbers recorded either way.
12. No public API change. `git diff` touches no file outside `XLibur/Excel/IO/`, `XLibur.Tests/` and
    `docs/specs/`.
13. `git diff --numstat` shows no file whose changed-line count approaches its total line count,
    except the files this spec deliberately deletes or creates.

## Conflicts

| Spec | Shared files | Severity | Order |
|---|---|---|---|
| **29** — write-path resolvers | `SheetViewWriter.cs`, `ColumnWriter.cs` | **Hard** | **29 first** |
| **17** — picture styling | `PictureWriter.cs` | Hard | 31 defers slots 30/31 |
| **16** — DrawingML infrastructure | `PictureWriter.cs`, `ChartWriter.cs` | Hard | 31 defers slots 30/31/32 |
| **15** — shapes and text boxes | `PictureWriter.cs` | Hard | 31 defers slots 30/31/32 |
| **22** — chart concept modules | `ChartWriter.cs` | Soft | Either; 31 does not enter `ChartWriter.cs` |
| **03** — save-path allocations | `SheetDataWriter.cs`, `CellXmlWriter.cs` | None | Disjoint |
| **24** — worksheet element load | `WorksheetElementReader.cs`, `XLWorkbook_Load.cs` | None | Disjoint; mirror image |
| **18 task 5** | `XLWorkbook_Load.cs` | None | Disjoint from this spec |

**Spec 29 goes first, and this is not a preference.** 29 fixes a live divergence: `SheetViewWriter.cs:124`
writes `state="frozenSplit"` unconditionally while `XLStreamingWorksheet.cs:502` writes
`state="frozen"` for the same model state, so `FreezeRows(2)` emits different XML depending on which
save path the caller used — and the DOM path is the wrong one, against 27 of 30 `<pane>` tags in the
test corpus. 29 is a small correctness fix in two files. 31 deletes `SheetViewWriter.cs` outright and
moves `ColumnWriter.WriteColumns` into a new namespace.

Rebasing 29 onto 31 means re-deriving a correctness fix against rewritten code, and a fix that has to
be re-derived is a fix that can be re-derived wrong. Rebasing 31 onto 29 means moving two extra lines.
The README already records the pair as **29→31** with the reason "never rebase a correctness fix onto
a structural sweep". Follow it.

**The drawing family (15, 16, 17) is designed around, not sequenced against.** All three live in
`PictureWriter.cs` and `ChartWriter.cs`, and 15 and 17 both hard-depend on 16, so that territory is
occupied for a while. This spec therefore leaves slots 30, 31 and 32 as four explicit calls at the
tail of the driver — stated as a non-goal, commented in the code, and cheap to finish later: three
more classes and three more list entries, with the corpus already in place to gate them. The README
anticipated this ("31 waits or defers that one writer"); deferring is the choice.

**Spec 22 is soft.** It reorganises `ChartWriter.cs` internals; this spec calls
`ChartWriter.WriteCharts` and does not enter the file. The one thing to watch is
`ChartWriter.EnsureDrawingElement` (`:1118-1133`), the second slot-30 owner — if 22 moves it, the
evidence citation in this spec moves with it, but no code does.

**Spec 03 is disjoint.** It owns `SheetDataWriter` and `CellXmlWriter`, which this spec's non-goals
exclude. The one point of contact is `WorksheetPartWriter.StreamToPart` (`:296-335`), which this spec
does not touch at all.

**Spec 24 is the mirror and shares nothing.** It works in `XLWorkbook_Load.cs` and
`WorksheetElementReader.cs`. The two can run concurrently, and doing so is a good idea: the load and
save sides end up with matching shapes — `in <Context>, ref <State>` — which is the strongest
argument either spec has for its design.
