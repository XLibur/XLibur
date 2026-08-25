# Round-trip fidelity

What survives a load → save round trip for content XLibur has no object model for.

Every claim here is backed by a test in `XLibur.Tests/Excel/RoundTripFidelityTests.cs`. If someone
later changes the save path to rebuild packages from scratch, those tests fail rather than quietly
dropping a user's chartsheets.

## The mechanism

XLibur does not build a new package on save. It reopens the package it loaded and rewrites only the
parts it models:

```csharp
// XLWorkbook_Save.cs
var package = File.Exists(filePath)
    ? SpreadsheetDocument.Open(filePath, true)
    : SpreadsheetDocument.Create(filePath, spreadsheetDocumentType);
```

`SaveAs` to a new path or stream copies the original package there first and then does the same
thing. So the default for any part XLibur does not understand is **survival**, not loss — the
opposite of what you would expect from a library that reads into a model and writes back out.

Two consequences:

- Preservation is free and automatic. There is nothing to opt into.
- Preservation requires an original package. A workbook created with `new XLWorkbook()` has nothing
  to carry forward, which is the hard boundary of everything below.

## Findings

| Content | Parts survive | Sheet references survive | Modelled | Notes |
|---|---|---|---|---|
| Chartsheets | ✅ | ✅ | ❌ | Kept as `XLWorkbook.UnsupportedSheet` |
| Dialog / macro sheets | ✅ | ✅ | ❌ | Same path as chartsheets |
| Form controls | ✅ | ✅ | ❌ | `<controls>` incl. `mc:AlternateContent` and anchors |
| ActiveX | ✅ | ✅ | ❌ | Both `activeX1.xml` and `activeX1.bin` |
| Timelines | ✅ | n/a | ❌ | `timelines/` and `timelineCaches/` |
| Custom XML | ✅ | n/a | ❌ | `customXml/item*.xml` and props |
| Slicers | ✅ | ✅ | ✅ | Modelled as of PRD 5; a slicer nobody edits is still passed through untouched |
| Threaded comments | ✅ | ✅ | ✅ | Modelled as of spec 09 |

### Chartsheets are preserved, not dropped

Spec 09 recorded that chartsheets "land in `XLWorkbook.UnsupportedSheet` and are **dropped on
save**". That is not what happens. `WorkbookPartWriter.GenerateContent` reorders the sheets it does
model *around* the unsupported ones rather than rewriting the `<sheets>` list from scratch:

```csharp
var totalSheets = sheetElements.Count() + xlWorkbook.UnsupportedSheets.Count;
```

so the `<sheet name="Chart" sheetId="3" r:id="rId3"/>` entry stays, the `xl/chartsheets/sheet1.xml`
part is never touched, and the file reopens intact in Excel and in XLibur.

The real limitation is narrower and worth stating plainly: chartsheets are **preserved but not
manipulable**. They are not `IXLWorksheet`, so `wb.Worksheets` does not include them and there is no
API to read or change one. Preservation only breaks if the sheet order is changed such that the
reorder logic mismatches, which is a separate concern from dropping content.

**Do / can't-do conclusion for chartsheet pass-through: already done, no work required.** Adding a
real chartsheet model would be a feature, not a fidelity fix.

### Form controls and ActiveX keep their anchors

Worth calling out because the worksheet part *is* regenerated from the model on every save, so this
one is not free in the way an untouched part is. `WorksheetPartWriter` carries the `<controls>`
element through verbatim, including the `mc:AlternateContent` wrapper and the `<anchor>` offsets
inside `controlPr`. Surviving `activeX1.bin` would be worthless if the sheet stopped referencing it.

### Slicers survive, and there are four separate reasons they could not have

This one was an open question until an Excel-authored fixture existed to settle it
(`Resource/TryToLoad/SlicersOnPivotAndTable.xlsx`: a table with a slicer on one sheet, a pivot table
with a second slicer on another). It is worth listing what had to hold, because "the parts are
untouched" only covers the first of four:

- **The parts.** `xl/slicers/slicer*.xml` and `xl/slicerCaches/slicerCache*.xml` are never opened, so
  they come through byte for byte — including the pivot slicer's caption, its `SlicerStyleDark3`
  style and its single selected item, none of which XLibur models.
- **The worksheet references.** The worksheet part *is* rebuilt on every save, and a slicer is
  referenced from `<extLst>`. Excel uses a different extension URI for a table slicer
  (`{3A4CF648-…}`) than for a pivot slicer (`{A8765BA9-…}`); both survive, because
  `ConditionalFormattingWriter` and friends only add and remove their own URIs instead of rewriting
  the list.
- **The workbook references.** Slicer caches are registered in `xl/workbook.xml`'s `<extLst>` in two
  places — `x14:slicerCaches` for the pivot slicer, `x15:slicerCaches` for the table slicer. Both
  survive, as do the `slicerCache` relationships they point at.
- **The defined names.** Excel emits a `#N/A` defined name per slicer cache (`Slicer_Region`,
  `Slicer_Region1`). These were the most exposed piece of the arrangement, since XLibur models
  defined names and rewrites the whole block. They come through intact.

The saved package also produces zero `OpenXmlValidator` errors, the same count as Excel's original.

#### Slicers are modelled now, and the pass-through still holds

PRD 5 gave slicers a model — `IXLWorksheet.Slicers`, `IXLPivotTable.Slicers`, creation, editing and
a cascade on pivot-table deletion — which turns the first of the four mechanisms above from a
property of the code into something the code has to keep on purpose.

It is kept the same way spec 10 keeps chart XML. **A slicer is never regenerated.** The reader takes
the parts apart through an `OpenXmlPartReader`, which streams a part and hands back a *detached*
tree, so `part.RootElement` is never materialised and the SDK has nothing to write back. Reaching a
slicer through `SlicersPart.Slicers` instead would attach a DOM the SDK re-serialises on save — the
same trap `WorksheetPartWriter` records for `worksheetPart.Worksheet` — and every attribute listed
above would be replaced by the SDK's own rendering of it.

Only two things open a slicer part, and both are gated on the caller having actually assigned
something: `SlicerPatcher` for a property, and `SlicerAnchorXml.Move` for a position. Reading every
property of every slicer changes nothing.

`SlicerReadModelTests.Reading_a_slicer_leaves_its_part_byte_for_byte_identical` and
`SlicerWriteTests.Editing_one_slicer_leaves_the_other_slicers_part_untouched` are the byte-level
gates on that, and `SlicerPositionTests.Moving_a_slicer_does_not_touch_its_slicer_part` proves the
two halves are gated separately.

One thing that does change is unrelated to slicers: `<pivotCache cacheId>` is renumbered on save
(`3` → `0`). The slicer cache binds to the pivot cache through `pivotCacheId` in
`pivotCacheDefinition1.xml`, not through this attribute, so the link is not affected.

### Comment VML is pruned surgically

`DeleteExistingCommentsShapes` removes only shapes whose `type` matches the comment shapetype:

```csharp
.Where(e => e.Name.LocalName == "shape" &&
            e.Attribute("type")?.Value == "#" + XLConstants.Comment.ShapeTypeId)
```

Form-control shapes in the same `vmlDrawing1.vml` are left alone, and `RemoveEmptyVmlPart` only
deletes the part when nothing at all is left in it.

## Gaps

- **Nothing is preserved for workbooks built from scratch**, by construction.
- **Preservation is not manipulation.** Everything in the table above still marked ❌ is opaque: it
  round-trips, but there is no API to inspect or edit it, and no guarantee it stays consistent if you
  delete the worksheet or pivot table it depends on. Slicers are the exception, and deleting a pivot
  table now takes its slicers with it rather than leaving them dangling.
- **Timelines are the remaining case of exactly that hazard.** They survive as untouched parts, and
  deleting the pivot table a timeline filters still leaves it pointing at nothing. PRD 5 task 4
  covers them.

## Asking a question can change the answer

A drawing part whose only content is a slicer frame used to come back from a save as the SDK's
serialisation rather than the producer's — self-closing tags gaining a space, `encoding="UTF-8"`
lower-cased, namespace declarations hoisted to the root. Nothing was lost: every element and
attribute survived, which is why no test caught it. They all assert with `Contains`, and `Contains`
cannot see a part that was rewritten rather than passed through.

The cause was one line in `PictureWriter.RemoveEmptyDrawingPart`, which asked whether the drawing
had any children so it could delete a part that would otherwise be saved empty. It asked through
`DrawingsPart.WorksheetDrawing`, and **reading that property attaches the SDK's typed tree to the
part**, after which the SDK writes the tree back over the original bytes on save whether or not
anything changed. The comment already on that line named the hazard and the guard order did not
avoid it: every preceding condition is true precisely for a sheet whose drawing holds only a slicer
or a timeline.

`DrawingPartProbe.HasAnyChild` answers the same question by streaming the part through an
`OpenXmlPartReader`, which leaves the root unmaterialised — the technique `SlicerReader` uses. It
falls back to the attached tree when one already exists, because by then the bytes on disk are stale
and there is no fidelity left to protect: `PictureWriter` may have just deleted the drawing's last
picture out of that tree, and answering from the stream would leave an emptied part in the package.

Both directions are pinned by tests, because each is invisible to the other's.

## Threaded comments

As of spec 09 threaded comments are modelled rather than preserved, so they are the exception to
this document. See `XLibur/Excel/Comments/Threaded/`. The one piece still handled by preservation is
`<mentions>`, whose raw XML is round-tripped on unmodified comments and dropped when the comment's
text changes, because a mention's `startIndex`/`length` index into the text it was written against.
