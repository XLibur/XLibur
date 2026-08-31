# Spec 44 — Data validation: one mapping, two adapters

**Area:** Architecture · **Correctness (divergence)**
**Effort:** M (~4–5 days)
**Dependencies:** None hard. Same shape as spec 29's write-path resolvers and should read like them.
**Status:** Proposed. From the 2026-08-30 architecture review (round 3).

## Problem Statement

A data validation rule is thirteen settings — what is allowed, the operator, two formulas, whether
blanks are ignored, whether the dropdown shows, the input and error messages and their titles and
style. Excel spells a validation rule two ways: the standard element, and an extension element used
when a rule needs something the standard form cannot express, such as a list that refers to another
sheet or a formula over 255 characters.

That thirteen-item list is written out by hand **seven times** in the library: a reader and a writer
for each of the two spellings, the copy operation, the initialiser, and the equality test used when
consolidating adjacent rules. Nothing checks that the seven agree.

Two of them already disagree:

- The standard writer always emits both formula elements, even when empty; the extension writer omits
  an empty one. Two files describing the same rule.
- The standard path applies a sheet-level "has anything changed" gate before writing; the extension
  path applies no gate at all.

A third piece of evidence sits in the consolidation equality test: it compares one property twice
under two names, because one is an alias for the other. A comparison that can never fail is what a
hand-copied list looks like after someone adds a property.

There is also no data-validation IO module at all. The extension reader lives inside the conditional
formatting reader, which the file's own summary comment apologises for.

## Solution

One module describes what a validation rule is and how it is spelled, with two adapters at its seam —
one per spelling. The choice between them, and the rule that decides when a validation needs the
extension form, sit behind that seam rather than in the writer's control flow.

The copy operation and the consolidation equality test derive from the same property list instead of
restating it.

## User Stories

1. As a library consumer, I want a validation rule to round-trip identically whether it is written in
   the standard or the extension form, so that which form was chosen is invisible to me.
2. As a library consumer, I want a list validation that refers to another worksheet to survive a save
   and reload, so that cross-sheet lists work.
3. As a library consumer, I want a validation with a formula longer than the standard form allows to
   survive a round trip, so that long formulas are usable.
4. As a library consumer, I want the ignore-blanks setting to round-trip in both forms, so that it does
   not depend on which form my rule happened to need.
5. As a library consumer, I want the in-cell dropdown setting to round-trip in both forms, for the same
   reason.
6. As a library consumer, I want input messages and their titles to round-trip in both forms.
7. As a library consumer, I want error messages, error titles and the error style to round-trip in both
   forms.
8. As a library consumer, I want the operator and the allowed-value type to round-trip in both forms.
9. As a library consumer, I want an empty formula to be written consistently, so that two files
   describing the same rule are not spelled differently.
10. As a library consumer, I want a validation I have not modified to be written the same way whichever
    form it uses, so that the dirty-tracking gate does not apply to only half of my rules.
11. As a library consumer, I want copying a validation to carry every setting, so that a copy is a copy.
12. As a library consumer, I want consolidating adjacent identical validations to merge them and
    non-identical ones to stay separate, so that consolidation is correct on every property.
13. As a library consumer, I want two validations differing only in a property the equality test
    currently ignores to remain separate, so that consolidation does not merge unlike rules.
14. As an XLibur maintainer, I want the thirteen settings listed once, so that adding a fourteenth is
    one edit rather than seven.
15. As an XLibur maintainer, I want the "does this rule need the extension form" decision stated in one
    place, so that the two writers cannot make it differently.
16. As an XLibur maintainer, I want the extension reader to live with the other validation code rather
    than inside conditional formatting, so that the file layout reflects the concepts.
17. As an XLibur maintainer, I want a table-driven test that round-trips every property through both
    adapters, so that a divergence between them fails a test.
18. As a contributing agent, I want two adapters at a named seam, so that I can tell which spellings
    exist without reading both writers.

## Implementation Decisions

**The seam is a validation-rule-to-element mapping with two adapters.** This is a genuine seam rather
than a hypothetical one: two spellings exist, both are exercised in production, and the library must
choose between them. That satisfies the two-adapters test — a seam is worth introducing because
something actually varies across it.

**The property list becomes the single source.** Reading, writing, copying, initialising and comparing
are five consumers of one list. The consolidation equality test is derived from it, which removes the
duplicated-alias comparison as a side effect rather than as a separate fix.

**The extension predicate moves behind the seam.** The decision to use the extension form currently
lives in the writer's control flow, expressed as two conditions — whether the rule references another
sheet, and whether a formula exceeds the length limit. It becomes a property of the mapping.

**Empty-formula emission is unified, and the choice is made deliberately.** The two writers disagree
today. The spec does not assume one is right: the correct behaviour is whatever Excel itself emits,
established by inspecting an Excel-authored file, and recorded in the spec's results.

**The dirty gate applies uniformly or not at all.** Today it filters the standard path at sheet level
and does not touch the extension path. Whichever behaviour is kept, it is the same for both adapters.

**The extension reader moves.** It currently sits inside the conditional formatting reader for
historical reasons the file itself notes. It moves to the validation module.

**No public API change.** `IXLDataValidation` keeps its shape; only where the mapping lives changes.

## Testing Decisions

**What makes a good test here.** A good test builds a validation through the public interface, saves,
reloads, and asserts the settings came back. Where the two spellings must agree byte for byte, the
test reads the emitted elements instead — a reload cannot see a spelling difference, because the
reader normalises both spellings back to one model. That is the same reasoning the existing write-path
agreement tests state, and the reason the frozen-pane divergence shipped unnoticed.

**The centrepiece is a property-by-adapter matrix.** For each of the thirteen properties, set it to a
non-default value, round-trip it through the standard adapter and through the extension adapter, and
assert equality in both. Today the entire extension path has one test, covering one validation type
with one long value.

**A byte-level agreement test between the adapters.** For a rule expressible in both forms, assert the
two adapters agree about what they emit for each property — in particular the empty-formula case. This
is the test that catches the class rather than the instance.

**A consolidation equality test per property.** Two validations differing in exactly one property must
not merge. Thirteen cases, one per property, driven from the same list.

**A copy fidelity test.** Copy a validation with every property set off-default and assert all thirteen
survive.

**Prior art.** The existing write-path agreement tests are the model for the byte-level half and their
header states the reasoning. The existing data validation tests are the right home for the round-trip
half. The conditional formats consolidation tests show the shape for the equality cases.

**Test seam.** `IXLDataValidation` and `IXLDataValidations` through a save/reload round trip, plus the
existing byte-level harness. No new test seam.

## Out of Scope

- Conditional formatting, beyond moving the misplaced extension reader out of it. Conditional format
  defects are specs 48 and 49.
- Adding validation types or operators.
- The sheet-level dirty-tracking mechanism itself; this spec only makes its application consistent.
- Extension-list content beyond data validation.

## Further Notes

Seven restatements of one list is the largest count this review round found, and the dead comparison in
the equality test is the most direct evidence available that the list is copied rather than derived —
it is what happens when someone adds a property to a list by copying the line above it.

The two live divergences are both benign today in the sense that no user has reported them, and both
are exactly the kind that becomes a data-loss defect the moment a third form is added or a reader
becomes stricter. The frozen-pane divergence spec 29 fixed had the same character right up until it
did not.

The divergences were established by reading the two writers against each other; the dead comparison
was verified against the property it aliases.
