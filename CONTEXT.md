# XLibur

A .NET library that reads, edits and writes Excel workbooks without Excel. This glossary names the spreadsheet concepts the code models, so specs, reviews and code use one word for each.

## Language

### Formulas

**Formula text**:
The characters of a formula as stored in the file or typed by a caller, without the leading `=`. It is distinct from the parsed formula that the engine evaluates.
_Avoid_: formula string, expression (for the stored form)

**Refused formula**:
Formula text that the parser cannot read, so the references it contains are unknown.
_Avoid_: invalid formula, broken formula, unparseable formula

**Future function**:
A function newer than the file format. The file stores it with an `_xlfn.` or `_xlws.` prefix, and Excel hides that prefix when it displays the formula.
_Avoid_: prefixed function

**Cached value**:
The value saved with a formula in the file. A reader that does not recalculate shows this value as-is.

**Dirty**:
Describes a formula whose cached value may no longer match its precedents, so it must be recalculated before anyone reads it.
_Avoid_: stale, pending, uncalculated

**Circular reference**:
A formula that depends on its own value, either directly or through other formulas.

### Results and failures

**Error value**:
One of the error results Excel shows in a cell, such as `#DIV/0!` or `#N/A`. An error value is a result, not a failure.
_Avoid_: error (unqualified)

**Unsupported feature**:
A valid Excel construct that XLibur does not evaluate.
_Avoid_: not implemented

**Defect**:
A failure caused by a bug in XLibur rather than by the content of the workbook.
_Avoid_: internal error

### Names and sheets

**Defined name**:
A name that refers to formula text, for example a range, a constant or a calculation. In the file, print areas and print titles are also defined names.
_Avoid_: named range (for names in general)

**Scope**:
Where a defined name is visible: the whole workbook, or one sheet. On its own sheet, a sheet-scoped name hides a workbook-scoped name that has the same name.

**3D reference**:
A reference to the same cells on every sheet from a first sheet to a last sheet in tab order, for example `Sheet1:Sheet3!A1`. The first and last sheets are its endpoints.
_Avoid_: multi-sheet reference, cross-sheet range

**Unsupported sheet**:
A sheet that XLibur keeps in the file but does not model, for example a chartsheet. It still has a position and a name in the workbook.
_Avoid_: chartsheet (for the general case)
