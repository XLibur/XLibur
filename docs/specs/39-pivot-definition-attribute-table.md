# Spec 39 — One attribute table for the pivot table definition

**Area:** Architecture · **Defect (wrong source, phantom API, documented crash)** · API (additive)
**Effort:** M–L (~6–8 days)
**Dependencies:** None hard. Shares no file with spec 35, which is complete. Fold in the pivot
write-path enum duplication noted below.
**Status:** ✅ Merged — [#420](https://github.com/XLibur/XLibur/pull/420) (**breaking**, squash `81c515c7`, 2026-08-31). From the 2026-08-30 architecture review (round 3). See Results.

## Problem Statement

A pivot table carries roughly sixty settings. They are enumerated by hand in five places — the
reader, the writer, the Excel-defaults initialiser, the copy operation, and a hand-written test of
forty assertions. Nothing checks that the five lists are the same list, and they are not.

What a user sees:

- **A pivot table re-saves with formatting the user never asked for.** The "show last column"
  emphasis is read from the *column stripes* attribute. A file with stripes on and last-column
  emphasis off re-saves with the emphasis switched on; a file with the emphasis on and stripes off
  loses it.
- **Two public properties do nothing.** A pivot table's title and description can be set through the
  public interface, are documented, and are touched by neither the reader nor the writer. Set them,
  save, reload, and they are empty.
- **Copying a pivot table silently resets twenty-three settings** that both the reader and the writer
  carry — including compact and outline form, visual totals, the grand-total caption and the data
  caption.
- **A documented crash condition is reachable.** The subsystem's own notes record that a definition
  whose row or column fields reference the values field must carry the data-position attribute. The
  loader never sets it, so a file that needs it is written without it.

## Solution

The sixty settings are described once, as data: for each setting, its attribute name, the property it
maps to, and its OOXML default. Reading, writing, default-seeding and copying are all driven from that
one description.

A setting can then no longer be present in the writer and absent from the copy, or read from the wrong
source, or given one default in one place and a different default in another — because there is only
one place.

## User Stories

1. As a library consumer, I want a pivot table's last-column emphasis to round-trip independently of
   its column stripes, so that re-saving a file does not change how it looks.
2. As a library consumer, I want a pivot table's column-stripe setting to round-trip without altering
   any other setting, so that one change does not have side effects.
3. As a library consumer, I want the title I set on a pivot table to be saved and reloaded, so that a
   public property that accepts a value keeps it.
4. As a library consumer, I want the description I set on a pivot table to be saved and reloaded, for
   the same reason.
5. As a library consumer, I want a copied pivot table to keep its compact and outline form, so that
   the copy is laid out like the original.
6. As a library consumer, I want a copied pivot table to keep its grand-total and data captions, so
   that the copy reads the same.
7. As a library consumer, I want a copied pivot table to keep its visual-totals, drop-zone and
   asterisk-totals settings, so that copying is not a partial operation.
8. As a library consumer, I want every setting the reader and writer both carry to survive a copy, so
   that I do not have to know which subset is supported.
9. As a library consumer, I want a pivot whose axes reference the values field to be written with the
   attribute that makes it valid, so that the file opens without repair.
10. As a library consumer, I want a setting left at its Excel default to be omitted from the file, so
    that XLibur's output stays close to what Excel itself writes.
11. As a library consumer, I want a setting explicitly set to a non-default to be written, so that my
    intent reaches the file.
12. As a library consumer, I want the defaults a new pivot table starts with to match the defaults the
    reader assumes, so that a created pivot and a loaded pivot behave the same.
13. As a library consumer, I want `ShowLastColumn` available on the public pivot table interface
    alongside its four siblings, so that I can set and assert it like any other display setting.
14. As an XLibur maintainer, I want adding a pivot setting to be one row of data, so that it cannot be
    added to the writer and forgotten in the copy.
15. As an XLibur maintainer, I want a setting's default stated exactly once, so that two places cannot
    disagree about what the default is.
16. As an XLibur maintainer, I want the copy operation derived from the same list as the round trip, so
    that it cannot be a subset of it.
17. As an XLibur maintainer, I want one property-based test to replace forty hand-written assertions,
    so that the roughly twenty settings the hand-written test misses become covered.
18. As an XLibur maintainer, I want the pivot write path to use the shared enum mapping rather than its
    own string table, so that nine duplicated enums stop needing to agree by hand.
19. As a contributing agent, I want the attribute list to be readable as a single table, so that I can
    answer "does XLibur support setting X" without cross-referencing four files.

## Implementation Decisions

**The seam is a declarative attribute table.** Each entry names the OOXML attribute, the property it
binds to, and the value that counts as the default. Reading, writing, defaulting and copying are four
consumers of that one table. This is the same shape as spec 38's view-state module, and the two should
read similarly.

**Derived attributes are computed, not stored.** Some attributes are functions of the pivot's other
state rather than settings in their own right — the row and column page counts, and the data position
that records where the values field sits on its axis. These are marked as derived in the table and
recomputed on write. The data-position defect exists precisely because it is treated as stored state
that only one of the two construction paths sets.

**`ShowLastColumn` becomes public.** Its four siblings are already on the public pivot table
interface; it alone is internal, which is why no test can reach the defect through the interface. This
is an additive change to `PublicAPI.Unshipped.txt`. Confirmed as the preferred option with the owner
before this spec was written.

**Title and description get resolved, not left ambiguous.** They are public, settable, and persisted
nowhere. The spec's decision is to give them their real home in the file rather than remove them from
the interface — removal would be a breaking change, and the attributes exist in the format. If they
turn out not to have a home, the finding is reported rather than worked around.

**The nine duplicated pivot enums fold in.** The pivot write path carries its own raw-string table for
nine enums that the shared enum converter also maps, with four of the converter's methods dead as a
result. They agree today — including the two spellings that differ only in case, which is exactly the
kind of agreement that fails silently. Since this spec is already rewriting how the writer produces
attributes, the enums move to the shared converter here rather than in a spec of their own.

**Ordering.** The wrong-source defect and its test land first, on their own, because it is a one-line
fix with a user-visible effect and should not wait for the table. The table follows, then the four
consumers migrate one at a time.

## Testing Decisions

**What makes a good test here.** A good test sets pivot settings through the public interface, saves,
reloads, and asserts through the same interface. Where the thing under test is whether an attribute is
omitted at its default, the test reads the emitted attributes instead — that is the one case a reload
cannot see, because the reader supplies the default again on the way back.

**The centrepiece is a property-based round trip.** Set every attribute in the table to a non-default
value, save, reload, and compare the whole table. Then do the same through a copy. Because the table
is data, the test iterates it, so a setting added later is covered automatically. This subsumes the
existing forty-assertion test and covers the roughly twenty settings it does not reach.

**Named regression tests for the four defects.** The last-column emphasis read from the wrong source;
the title and description that vanish; the twenty-three settings a copy resets; the missing
data-position attribute on a pivot whose axis references the values field.

**Default polarity gets byte-level assertions.** For each attribute, assert that it is absent from the
output when it holds its default and present when it does not. This is the check that would have
caught the two places where the defaults disagree.

**The enum fold-in gets a round-trip test per enum.** Nine enums, both directions, every member —
including the two whose spellings differ only in case, which is the pair most likely to drift.

**Prior art.** The existing pivot table options round-trip test is the right shape and the right home;
it builds a workbook, sets options, saves, reloads and asserts. Its weakness is that its assertion
list is written by hand, so it can only ever check what someone remembered. The pivot filter criteria
reader and writer are the area's good example of a symmetric pair and are worth reading before
starting.

**Test seam.** `IXLPivotTable` and a save/reload round trip, plus emitted-attribute assertions for
default polarity. No new seam.

## Out of Scope

- The pivot cache — its value encoding, its records part and its field definitions. That is spec 41.
- Pivot conditional formatting and the priority handshake that links it to the sheet's rules.
  Recorded as a backlog note from the same review.
- Pivot table area movement on a structural edit, which spec 33 covered and merged.
- Timelines, which spec 35 delivered.
- Adding pivot features. This spec makes existing settings round-trip and copy correctly.

## Further Notes

Five enumerations of one list is the clearest instance of this round's shape, and the wrong-source
defect is a good illustration of why it matters: it is a single wrong identifier, in a line that looks
exactly like the four lines above it, in a file where sixty such lines are written out by hand. No
amount of care makes that line reliably correct. Making the list data makes it impossible to write.

The `ShowLastColumn` case is worth noting for its own sake — the one property in the group that no
test could reach through the interface is the one that is wrong. That is what "the interface is the
test surface" means in practice.

The wrong-source line was verified directly while writing this spec.

## Results

**Implemented 2026-08-30 on `task/39` (worktree `xl-wt-39`), head `f78a2fe8`, 11 commits, cut from
`upstream/main` `37c986bb`. Not yet pushed or merged.** Full suite **run independently by the
orchestrator** on both TFMs: 28,526 total, 0 failed, 10 pre-existing skips.

**The agent's final report was lost** — its session was `/clear`ed from the tab before the
orchestrator read it — so this section is reconstructed from the commit bodies, the diff, and the
orchestrator's own checks. The "least sure of" answer for this spec was never captured; treat the
review findings below as the substitute.

**Order landed as briefed.** `060bf988` the wrong-source fix alone, red-first
(`ShowLastColumn_reads_from_its_own_attribute_not_column_stripes`); `9777a892` `ShowLastColumn`
public (four lines in `PublicAPI.Unshipped.txt`); `1e1468d3` the other three regressions red;
`d2db5b4c`/`72355690` prerequisites; `5e2625cb` the table + reader; `2c17cddc` writer; `a890197d` the
three fixes + copy; `e2f1b7d2` enum fold-in; `f8df87d0` the table-driven property test replacing the
40-assertion `PivotTableOptionsSaveTest`.

**Answers to the two questions the brief asked to be settled.**

- **Title/Description have a home**: the x14 `pivotTableDefinition` extension's `altText` /
  `altTextSummary` — the same extension element that already carried `EnableCellEditing` and
  `ShowValuesRow`, and the pair Excel surfaces as a pivot table's Alt Text. Read and written there;
  `Title_and_description_round_trip_through_save_and_reload`.
- **`dataPosition`**: the loader added the `data` (-2) field straight to an axis without passing
  through the one path that set `DataPosition`, so a loaded-then-resaved file lost the attribute. Now
  a **computed** property derived from where -2 sits in `RowAxis`/`ColumnAxis`, never stored;
  `DataPosition_is_written_after_a_reload_when_an_axis_references_the_values_field`. Whether Excel
  actually crashes on the omission was not re-verified in-branch — the commit body records only the
  documented condition and the fix.

**What the spec predicted that turned out wrong.** The "Excel-defaults initialiser" as a fifth
enumeration barely existed: of `SetExcelDefaults()`'s 28 assignments, 23 duplicated a field
initialiser or the CLR default; only five carried real information, now inline on the property.
`ShowRowHeaders`/`ShowColumnHeaders` default `true` for a fresh pivot but `false` as the OOXML
round-trip default — pre-existing, intentional, now stated in the table rather than silently split.

**The table.** `XLibur/Excel/IO/PivotTableAttributes.cs` (243 lines): each row is attribute name,
bound property, read/write/copy accessors and default. `dataPosition` and the location element's
page counts are deliberately **not** rows — derived, recomputed on write. The two enumerating tests
(`Pivot_table_attribute_table_round_trips_through_save_and_reload`,
`..._through_copy`) iterate it; the copy arm was shown to fail on the first attribute when the
`CopyTo` migration was temporarily reverted. The x14 flags and the `pivotTableStyleInfo` group stay
explicitly copied, now including `ShowLastColumn`.

**Enum fold-in.** The writer's nine-enum string table is gone; `EnumConverter` gained forward
`ToOpenXml` for the five it lacked; `EnumConverterPivotTests` round-trips every member of all nine
both ways and pins the case-only pair (`stdDevp`/`varp` on `ST_DataConsolidateFunction` vs
`stdDevP`/`varP` on `ST_ItemType`). Byte-identical output against the golden fixtures.

**Eight golden `.xlsx` fixtures were regenerated**, six in `060bf988` and two in `a890197d`. The
orchestrator diffed one part-by-part: the only change in `PivotTableWithoutSourceData-output.xlsx` is
`showLastColumn="1"` appearing on `pivotTableStyleInfo` — the attribute the old reader discarded.
The other seven were not individually diffed; a reviewer should spot-check the two from `a890197d`
(`PivotTables.xlsx`, `PivotSubtotalsSource/output.xlsx`), which should show the x14 alt-text and/or
`dataPosition` changes and nothing else.

**Owner's review pass, 2026-08-30 evening — committed as `5e3a0a27`, `27f54e6a`, `d5d73267`.** PR [#420](https://github.com/XLibur/XLibur/pull/420) opened 2026-08-31 at head `2088efc0` (16 commits, one further regression test; suite 28,528 per TFM at head). The PR body records that the regenerated fixture set carries package-metadata churn (`.psmdcp` rename, `_rels/.rels`) from being resaved. Four `/code-review` findings: (1)
`CopyTo` copied the table-level layout but not each field's Compact/Outline pair — now copied
per field, since fields in a loaded file can legitimately differ; test
`CopyTo_carries_the_layout_of_every_pivot_field_not_just_the_table`, shown to fail with one line
disabled. (2) `ToAttr<T>` allocated an `EnumValue<T>` per attribute; `IEnumValue.Value` already
exposes the interned string — zero allocations, byte-identical output. (3) **`ShowLastColumn` on
`IXLPivotTable` is source- and binary-breaking for external implementers** — moved in `CHANGELOG.md`
from Bug Fixes to ⚠️ Breaking Changes (the `IXLSheetView.FreezePanes` precedent). The PR title needs
the `!`. (4) `chartFormat`'s next-free id was copied without the `ChartFormats` collection it indexes;
now an explicit, documented copy exemption, and the copy test asserts the exemption plus a count so a
future silent exemption fails. Suite green on net10.0 across all four projects after the pass.

**Deliberately not done.** Pivot cache (spec 41), pivot CF handshake (backlog), area movement (33).

**What the next consumer inherits.** Spec 41 finds the definition side already table-driven and
should mirror the shape for the cache. Spec 38 built the same shape for sheet view; the two should be
read together as the round's pattern. `CHANGELOG.md` carries five `### Fixed` entries under
`## Unreleased`.

**Merged 2026-08-31** as [#420](https://github.com/XLibur/XLibur/pull/420) (squash `81c515c7`,
branch tip `2088efc0`). The merged CHANGELOG refines the copy numbers: the hand-written list carried
29 of the definition's 65 attributes and dropped 36; 35 are now copied and `chartFormat` is the one
deliberate, documented exemption.
