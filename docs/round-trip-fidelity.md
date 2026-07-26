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
| Slicers | ❓ | ❓ | ❌ | **Untested** — see gaps below |
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

### Comment VML is pruned surgically

`DeleteExistingCommentsShapes` removes only shapes whose `type` matches the comment shapetype:

```csharp
.Where(e => e.Name.LocalName == "shape" &&
            e.Attribute("type")?.Value == "#" + XLConstants.Comment.ShapeTypeId)
```

Form-control shapes in the same `vmlDrawing1.vml` are left alone, and `RemoveEmptyVmlPart` only
deletes the part when nothing at all is left in it.

## Gaps

- **Slicers are unverified.** The test suite has no fixture containing `xl/slicers/` or
  `xl/slicerCaches/`. The same mechanism should carry them through, but "should" is not "does" and
  this doc does not claim otherwise. Closing this needs an Excel-authored file with a slicer added
  to `XLibur.Tests/Resource/`.
- **Nothing is preserved for workbooks built from scratch**, by construction.
- **Preservation is not manipulation.** Everything in the table above is opaque: it round-trips, but
  there is no API to inspect or edit it, and no guarantee it stays consistent if you delete the
  worksheet or pivot table it depends on.

## Threaded comments

As of spec 09 threaded comments are modelled rather than preserved, so they are the exception to
this document. See `XLibur/Excel/Comments/Threaded/`. The one piece still handled by preservation is
`<mentions>`, whose raw XML is round-tripped on unmodified comments and dropped when the comment's
text changes, because a mention's `startIndex`/`length` index into the text it was written against.
