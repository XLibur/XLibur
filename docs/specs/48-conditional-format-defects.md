# Spec 48 — Conditional format defects: the crash, the throw, and the lost data bar

**Area:** **Defect (crash on load)** · Compat
**Effort:** S–M (~3 days)
**Dependencies:** None. **Deliberately scoped to require no public interface change**, so that it can
ship without waiting on spec 49. Spec 49 supersedes its internal shape but not its behaviour.
**Status:** Proposed. From the 2026-08-30 architecture review (round 3). Split from the value-object
candidate at the owner's direction so the crash fix is not blocked behind an API decision.

## Problem Statement

Three failures in conditional formatting, all reachable from files Excel routinely produces, none of
them reachable from files XLibur produces — which is why the suite cannot see any of them.

- **A data bar crashes the load.** When a user sets a data bar's negative fill to match its positive
  fill, Excel omits the negative-fill element entirely; the format allows this. XLibur dereferences it
  unconditionally and throws. The whole workbook fails to open. XLibur's own writer always emits the
  element, so no XLibur-authored fixture can produce the input.
- **A threshold without a value throws on save.** A conditional format value object may omit its value.
  The colour-scale converter guards against that; the icon-set converter does not, and dereferences
  it. Load such a file and saving throws.
- **An Excel-authored data bar comes back stripped.** The border, border colour, negative border
  colour and the two "same as positive" flags are neither read nor written, and the minimum and
  maximum bar lengths are reset to fixed values on every save regardless of the file. A bordered data
  bar becomes borderless; a bar with custom lengths gets default ones.

## Solution

Optional elements are treated as optional. Missing children produce the format's defined default
rather than an exception, and the data bar's full attribute set is read and written.

This spec changes behaviour only. The four parallel collections that make this class of defect easy to
introduce are left alone; replacing them is spec 49.

## User Stories

1. As a library consumer, I want a workbook containing a data bar with "negative fill same as positive"
   to open, so that a standard Excel setting does not make my file unreadable.
2. As a library consumer, I want that data bar to keep rendering the same way after a round trip, so
   that opening the file in XLibur does not change it.
3. As a library consumer, I want a conditional format value object with no explicit value to load
   without error, so that a valid file is accepted.
4. As a library consumer, I want to be able to save a workbook containing such a value object, so that
   loading it is not a dead end.
5. As a library consumer, I want that to hold for icon sets as well as colour scales, so that the two
   conditional format types behave alike.
6. As a library consumer, I want a data bar's border to survive a round trip, so that bordered bars stay
   bordered.
7. As a library consumer, I want a data bar's border colour and negative border colour to survive a
   round trip.
8. As a library consumer, I want the "negative bar colour same as positive" and "negative border colour
   same as positive" flags to survive a round trip, so that the setting that triggers the crash is also
   the setting that is preserved.
9. As a library consumer, I want a data bar's minimum and maximum lengths to survive a round trip, so
   that saving does not resize my bars.
10. As a library consumer, I want a data bar I created through XLibur to keep behaving exactly as it does
    today, so that the fix does not disturb the working case.
11. As a library consumer, I want every conditional format type to survive an open-and-resave of an
    Excel-authored file, so that XLibur is safe to use on files it did not create.
12. As an XLibur maintainer, I want optional elements handled as optional throughout the conditional
    format readers, so that this class of crash is eliminated rather than patched once.
13. As an XLibur maintainer, I want Excel-authored conditional format fixtures in the suite, so that
    inputs XLibur cannot produce are covered.
14. As an XLibur maintainer, I want each defect pinned by a test that fails before the fix, so that the
    fixes are proofs rather than assertions.
15. As a contributing agent, I want the format's optional children documented where they are read, so
    that the next reader does not assume presence.

## Implementation Decisions

**Optionality is read from the format, not inferred from XLibur's own output.** Each child element the
readers dereference is checked against the schema for whether it is required. Where it is optional, the
reader supplies the format's default. The negative-fill element is the known instance; the audit covers
the rest of the conditional format readers in the same pass, because one unguarded dereference
strongly suggests others.

**The two converters are brought into agreement.** One guards a missing value and one does not. The
guarded behaviour is correct and becomes the behaviour of both.

**The data bar's attribute set is completed.** The unmodelled attributes are added to the model, read
and written. The fixed minimum and maximum lengths become values carried from the file rather than
constants written on every save.

**No public interface change.** The four parallel collections stay, along with their alignment
requirement. This is a deliberate constraint: it keeps the crash fix independent of the API decision in
spec 49. Where a new data bar attribute needs a home, it is added additively.

**Ordering.** The load crash lands first, alone. Then the save throw. Then the data bar attributes.
Each with the fixture that fails before it.

## Testing Decisions

**What makes a good test here.** A good test loads an Excel-authored fixture, asserts the load
succeeded and the model is right, saves, reloads, and asserts equality. Fixtures produced by XLibur's
own writer cannot express any of these defects — the writer always emits what the reader assumes — so
the fixtures are the substance of the testing work, not an incidental.

**Fixtures required.** Authored in Excel and committed to the test resources: a data bar with negative
fill matching positive; a data bar with a border and custom lengths; an icon set whose value object
omits its value; a colour scale with the same. Each fixture's provenance recorded in a comment, because
a fixture nobody can regenerate is a liability.

**A load-save-load equality test per fixture.** Load, save, load again, assert the two loads agree. This
covers both the crash and the attribute loss in one assertion.

**A negative test for the save throw.** Load the fixture with the missing value, save, assert no
exception, reload, assert the model.

**An unguarded-dereference audit as a test.** For each optional child the conditional format readers
consume, a fixture omitting it. This is the check that turns a single fix into a class fix.

**Prior art.** The existing conditional format tests all build their workbooks through the public
interface, save, reload and assert — a good shape that cannot reach these inputs. The round-trip
fidelity tests are the model for the load-save-load pattern and already carry Excel-authored fixtures.

**Test seam.** `IXLConditionalFormat` and a load-save-load round trip over Excel-authored fixtures. No
new seam.

## Out of Scope

- The four parallel collections and their alignment requirement — spec 49.
- The alignment defect itself, where a value object without a type shifts every later entry onto the
  wrong type. That is a consequence of the collection shape and is fixed in 49; it cannot be fixed here
  without the interface change.
- The two data bar converter hierarchies and the string match that joins them — spec 49.
- Differential format decoding — spec 28.
- The pivot conditional format priority handshake — backlog note from the same review.
- Adding conditional format types.

## Further Notes

All three defects share a cause that is not really about conditional formatting: XLibur's reader is
written against XLibur's writer rather than against the format. Where the writer always emits an
element, the reader assumes it; where the writer hardcodes a value, the reader never learns to carry
one. The library is self-consistent and wrong about files from anywhere else.

That is a general risk in any read-write library and the only reliable defence is fixtures the library
did not author. This spec's real deliverable may be those fixtures rather than the three fixes.

The crash mechanism was established by reading the reader against the schema's optionality; the
converter asymmetry was established by diffing the two converters.
