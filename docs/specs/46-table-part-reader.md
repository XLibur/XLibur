# Spec 46 — The table part gets a reader

**Area:** Architecture · **Defect (wrong colour filter, dropped attributes)**
**Effort:** M (~5–6 days)
**Dependencies:** None hard. Reads the differential format collection that spec 28 established;
no file conflict with it.
**Status:** Proposed. From the 2026-08-30 architecture review (round 3).

## Problem Statement

A table's definition — its name, its columns, its totals row, its styling and its autofilter — has a
writer module. It has no reader module. The read half is three private methods inside the workbook
load routine, so there is nowhere that states what the two halves owe each other, and no seam at which
a test could check.

The writer also replaces the whole table part on save rather than editing it, so anything it does not
model is destroyed even on an open-and-resave that changes nothing.

What a user sees:

- **A colour filter on a table filters by the wrong colour after a round trip.** The reader that
  handles autofilter columns needs the workbook's differential formats to recognise a colour filter,
  and it takes them as an *optional* argument. The table call site omits it. The filter is therefore
  never recognised as a colour filter, the styles writer skips it when rebuilding the differential
  format collection, and the function that exists specifically to keep such a reference valid cannot
  find its entry. Since the collection is rebuilt from scratch on every save, the preserved reference
  now points at a different format. Reopen the file and the filter is by some other colour.
- **A table's sort state is dropped.** The writer emits it; the table read path never reads it back, so
  it is lost on load and then not rewritten.
- **A table can be renamed by saving it.** The writer forces the display name equal to the name and
  strips characters from it. A table whose two names differ, or whose name contains a stripped
  character, comes back renamed — which breaks every structured reference to it.
- **A totals row that is configured but hidden comes back as not configured.**
- **Column-level and table-level styling is destroyed**: the differential format applied to a column's
  data, the header and totals row formats, the border formats and the named cell styles are neither
  read nor written.

## Solution

The table definition part gets a reader module beside its writer, and the attribute set is stated once
for both directions. The differential formats the reader needs become part of its interface rather
than an optional argument a caller can leave out.

What the library does not model is carried through rather than dropped, so an open-and-resave stops
being lossy.

## User Stories

1. As a library consumer, I want a table's colour filter to filter by the same colour after a save and
   reload, so that reopening my file does not silently change what it shows.
2. As a library consumer, I want a colour filter to behave identically on a table and on a worksheet
   range, so that where the filter lives does not change whether it works.
3. As a library consumer, I want a table's sort state to survive a round trip, so that a sorted table
   stays sorted.
4. As a library consumer, I want a table's name to be unchanged by saving, so that structured references
   to it keep resolving.
5. As a library consumer, I want a table whose display name differs from its name to keep both, so that
   files authored elsewhere are not rewritten.
6. As a library consumer, I want a configured but hidden totals row to stay configured, so that
   unhiding it restores what was there.
7. As a library consumer, I want a column's data formatting to survive a round trip, so that table
   styling is not lost.
8. As a library consumer, I want header row and totals row formatting to survive a round trip, for the
   same reason.
9. As a library consumer, I want a table's named cell styles to survive a round trip.
10. As a library consumer, I want a calculated column's formula to survive a round trip, so that
    computed columns keep computing.
11. As a library consumer, I want table attributes XLibur does not model to survive an open-and-resave,
    so that the library is safe to use on files richer than its object model.
12. As a library consumer, I want opening and immediately saving a table-bearing file to produce an
    equivalent file, so that a no-op round trip is a no-op.
13. As an XLibur maintainer, I want the reader and writer to be a visible pair, so that an attribute
    present in one and missing from the other is apparent.
14. As an XLibur maintainer, I want the differential formats to be a required part of the reader's
    interface, so that a caller cannot omit them.
15. As an XLibur maintainer, I want the totals row's two attributes treated as one fact, so that they
    cannot describe contradictory states.
16. As an XLibur maintainer, I want the reader testable against the writer's output, so that the two
    halves can be checked directly.
17. As an XLibur maintainer, I want Excel-authored table fixtures in the suite, so that attributes
    XLibur never writes are still covered on load.
18. As a contributing agent, I want the table part's attribute set in one place, so that I can answer
    "does XLibur preserve X" without reading a writer and a load routine.

## Implementation Decisions

**The seam is a table definition part reader, paired with the existing writer.** Creating it is most of
the value: it gives the writer something to be tested against and gives the attribute set somewhere to
live.

**The differential format collection becomes a required argument.** The optional parameter is the
mechanism of the colour-filter defect — it turned a genuine interface requirement into something a
caller could silently skip, and one caller did. Making it required means the compiler enforces what
the interface always meant.

**The totals row is one fact with two attributes.** The count and the shown flag together describe the
totals row's state; reading only the count is why a configured-but-hidden row comes back wrong. They
are read and written as a unit.

**Name and display name are distinct.** The writer's current behaviour — force them equal, strip
characters — is a normalisation that belongs at the point a user *sets* a name, not at the point a
file is written. Saving must not rename.

**Unmodelled content is carried through.** The writer replaces the whole part, so anything not modelled
is lost. Either the unmodelled attributes get modelled, or the reader keeps them and the writer emits
them back. The area's pivot filter criteria code already does the latter for its own unmodelled
children and is the pattern to follow.

**Ordering.** The colour-filter defect lands first, on its own, with its test — it is silent data
corruption with a clear mechanism and should not wait for the reader module. The reader follows, then
the attribute set, then pass-through.

## Testing Decisions

**What makes a good test here.** A good test builds or loads a table, saves, reloads, and asserts
through the public table interface. For attributes XLibur cannot currently produce, the input must be
an Excel-authored fixture — a test built from XLibur's own writer can only ever check that the reader
agrees with this writer, which is the exact blind spot that let these defects ship.

**The centrepiece is a load-save-load fixture test.** Take Excel-authored tables covering the styling
attributes, the differing names, the hidden totals row, the calculated column and the sort state; load,
save, load again, and assert the second load equals the first. This catches every write-but-never-read
and never-touched attribute at once.

**A named regression test for the colour filter.** A table with a colour filter in a workbook that also
has other differential formats — the extra formats matter, because the defect is a reference pointing
at the wrong entry in a rebuilt collection, and with only one entry it would accidentally be right.
Save, reopen, assert the filter's colour.

**A parallel test on a worksheet range** asserting the same colour filter behaviour, so the two paths
are pinned as equivalent.

**An attribute inventory test.** For the attribute set the spec defines, assert each is read and
written. This is the mechanical check that the pair stays a pair.

**Prior art.** The existing colour filter tests all apply the filter to a worksheet range and never to a
table, which is why the table path is uncovered. The round-trip fidelity tests are the model for the
pass-through half. The pivot filter criteria reader and writer are the model for the symmetric pair.

**Test seam.** `IXLTable` and `IXLAutoFilter` through a save/reload round trip, plus Excel-authored
fixtures. The new reader is not itself a test seam — it is exercised through the round trip.

## Out of Scope

- The differential format decoding itself, which spec 28 owns and has merged.
- Table features as such — adding columns, calculated column authoring, totals functions.
- The autofilter criteria reader and writer, which are already symmetric.
- Structured reference resolution in formulas.
- Worksheet-level autofilter behaviour, except as the control case for the colour filter test.

## Further Notes

The optional parameter is the part of this worth generalising. `differentialFormats = null` looks like
a convenience and is in fact a statement that the reader can work without the workbook's formats — which
is not true, and the one caller that took the offer at face value produces corrupted output. Optional
arguments that are only optional for some callers are a way of encoding "you must know something the
signature does not tell you", which is the definition of a leaky interface.

The colour-filter case also shows why "the library round-trips its own output correctly" is a weak
guarantee. Everything here is consistent from XLibur's point of view; the damage is to a reference into
a collection that Excel reads and XLibur rebuilds.

The divergences were established by diffing the writer's emitted attributes against the load path's
consumed attributes; the optional-argument omission was verified at the call site.
