# Spec 41 — One codec for a pivot cache value

**Area:** Architecture · **Defect (2 shipped, both data-destroying)**
**Effort:** M (~5–6 days)
**Dependencies:** None hard. File-disjoint from spec 39, which owns the pivot *definition*; the two
can run in parallel under different owners.
**Status:** Proposed. From the 2026-08-30 architecture review (round 3).

## Problem Statement

A pivot table's cache stores the source data twice over — once as a list of distinct shared items per
field, once as the records themselves. Both are sequences of values, and a value can be a number, a
date, a string, a boolean, an error or blank. How each of those is spelled in XML is decided in four
separate places: a reader and a writer for the shared items, and a reader and a writer for the
records. Nothing checks that the four agree, and they do not.

What a user sees:

- **A pivot over data containing an error value corrupts the file.** An error in the source range is
  written into the records as a *boolean* element. Reload it and the boolean parser is handed
  `#N/A`, which it cannot read; Excel reports the cache as needing repair. The shared-items writer, in
  a different file, gets the same value type right — which is exactly why nobody noticed.
- **Grouping is destroyed on re-save.** Open a pivot grouped by months, change nothing, save. The
  grouping is gone, and worse, the definition now promises more fields than each record carries, so
  the part is structurally invalid. XLibur then refuses to load its own output. The element that
  describes a grouping is modelled nowhere, and the save path cannot pass it through untouched
  because it clears the loaded fields before writing.

Three smaller asymmetries sit in the same code: a record count that is read and never written; the
numeric range of a field that mixes dates and numbers, lost on every save; and a public setting for
how many items to retain that has three meanings on read and two on write, so one of them silently
does nothing.

## Solution

There is one description of how a pivot cache value is spelled in XML, and all four places use it.
The records part gets a reader module of its own, beside its writer, so that the two can be tested
against each other — today the records are read from inside the definition reader, so there is no
seam to test across.

The rule that a definition and its records must agree about field count becomes a property of that
module rather than an accident of two writers making compatible decisions.

## User Stories

1. As a library consumer, I want a pivot over source data containing `#N/A` to produce a file Excel
   opens without repair, so that error values in my data are not fatal.
2. As a library consumer, I want the same for `#DIV/0!`, `#REF!` and every other error value, so that
   no error type is a special case.
3. As a library consumer, I want an error value in a pivot cache to reload as the same error value, so
   that round-tripping preserves my data.
4. As a library consumer, I want to open a pivot grouped by months, save it, and still have the
   grouping, so that XLibur can be used on files that use a standard Excel feature.
5. As a library consumer, I want the same for grouping by quarters, years, and numeric ranges, so that
   no grouping kind is destroyed.
6. As a library consumer, I want a file XLibur has saved to be loadable by XLibur, so that the library
   is never the sole producer of files it rejects.
7. As a library consumer, I want a cache field mixing dates and numbers to keep its recorded minimum
   and maximum, so that Excel's own filtering behaves correctly.
8. As a library consumer, I want the record count in the file to match the records actually written,
   so that consumers relying on it are not misled.
9. As a library consumer, I want `SetItemsToRetainPerField` to take effect for every value it accepts,
   so that a setting the interface offers is not silently ignored.
10. As a library consumer, I want every value type — number, date, string, boolean, error, blank — to
    survive a cache round trip unchanged, so that no data type is lossy.
11. As a library consumer, I want a cache with a calculated field to keep working exactly as it does
    today, so that this change does not regress the case that currently works.
12. As an XLibur maintainer, I want the spelling of a cache value defined once, so that a value type
    cannot be written as one element and read as another.
13. As an XLibur maintainer, I want the records part to have a reader module, so that its writer has
    something to be tested against.
14. As an XLibur maintainer, I want the field-count agreement between the definition and the records to
    be enforced in one place, so that two writers cannot promise different shapes.
15. As an XLibur maintainer, I want a round-trip property test over the whole value union, so that
    adding a value type later is covered without remembering to write a test.
16. As an XLibur maintainer, I want non-database fields handled by the same module that decides field
    counts, so that skipping a column and declaring a column cannot disagree.
17. As a contributing agent, I want one place to look up how a cache value is encoded, so that I do not
    copy whichever of the four encodings I find first.

## Implementation Decisions

**The seam is a value-to-element codec, used in both directions by both parts.** It owns the mapping
from value type to element name and back, and it is the only thing that knows those names. Getting the
error case wrong becomes impossible because there is one case statement, not four.

**The records part gets a reader module.** It currently has none — the records are parsed from inside
the cache definition reader — which is why the records writer has never been tested against a reader.
Creating it is most of what makes the defects testable.

**Field-count agreement becomes an invariant of the module.** Two separate decisions must currently
match by hand: whether the definition declares a field, and whether the records writer emits a column
for it. They move behind one decision.

**Grouping fields get modelled, or the save path stops destroying them.** The element describing a
grouping is not modelled at all, and pass-through is unavailable because the save path clears the
loaded cache fields before writing. The spec's preference is to model it, because the surrounding
attributes are already modelled and a half-modelled part is what caused the invalidity. If modelling
proves larger than this spec, the fallback is to preserve the loaded element verbatim for fields
XLibur does not otherwise touch — but that fallback must be a deliberate, recorded decision, not a
default.

**The three smaller asymmetries are folded in**, because each is a one-line consequence of the codec
owning the element's attribute set: the record count, the numeric range on a mixed-type field, and the
retention setting's missing write case.

**Ordering.** The error-value defect lands first with its test — it is small, isolated and
data-destroying. The records reader follows. Grouping lands last, because it is the one with genuine
design content.

## Testing Decisions

**What makes a good test here.** A good test round-trips a workbook and asserts on what came back
through the public pivot interface, or — where the file's own validity is the point — asserts that the
saved file loads. For grouping, an Excel-authored fixture is required, because XLibur cannot currently
produce the input at all.

**The centrepiece is a property test over the value union.** For every value type the cache can hold,
write it, read it back, and assert equality. That single test covers the error defect and every future
value type. It is the test that could not exist while there was no records reader.

**A self-consistency test.** Save a workbook, reload it with XLibur, assert it loads. Sounds trivial;
the grouping defect is exactly a case where it fails, and the failure is a thrown exception rather
than a wrong value.

**Excel-authored fixtures for grouping.** Group by month, quarter, year and numeric range in Excel,
save the fixtures into the test resources, and assert each survives a load-and-save. XLibur-authored
fixtures cannot cover this, which is the reason the defect shipped.

**A field-count invariant test.** After any save, the number of fields the definition declares equals
the number of columns each record carries. Asserting this once catches the whole class.

**Prior art.** The pivot filter criteria reader and writer are the area's genuinely symmetric pair and
are the model to follow — they cover every attribute of every criteria child and preserve what they do
not model as raw XML. Read them before starting. The existing pivot cache tests build their fixtures
through XLibur's own writer, which is why they cannot see any of this.

**Test seam.** `IXLPivotTable` and its cache, plus save/reload and Excel-authored fixtures. The
records reader is new but is not a test seam in its own right — it is tested through the round trip.

## Out of Scope

- The pivot table *definition*'s sixty attributes — spec 39.
- Pivot conditional formatting and the priority handshake — backlog note from the same review.
- Timelines and slicers, both delivered.
- Adding grouping as an authoring feature. This spec is about not destroying grouping that already
  exists in a file; creating a grouping through the public interface is a separate feature.
- Cache refresh semantics and the relationship between a cache and its source range.

## Further Notes

The error-value defect is a good illustration of why "two implementations, agreement by hand" is worth
treating as a defect class rather than a style preference. The shared-items writer and the records
writer are two switch statements over the same enumeration, in two files, written at different times.
One of them is right. There is no mechanism by which the other would have been corrected.

The grouping defect is the more serious of the two because its output is structurally invalid rather
than merely wrong, and because the library rejects its own product — which means any round trip
through XLibur is a one-way door for those files.

Both were reproduced by reading the reader against the writer element by element; the self-rejection
was confirmed against the loader's own structural check.
