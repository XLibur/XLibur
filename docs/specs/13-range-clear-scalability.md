# Spec 13 — `Clear` and `CopyTo` scalability

**Area:** Perf (edit) · Correctness | **Effort:** S | **Status:** Proposed (July 2026)

Tracking issue: [XLibur/XLibur#271](https://github.com/XLibur/XLibur/issues/271)

## Summary

`IXLRange.CopyTo` gets steadily slower as the target worksheet grows, so **copying a range in a loop is
quadratic in the number of copies**. A 1×10 `CopyTo` costs ~120 µs on a nearly-empty 30,000-row sheet and
~420 µs by the time 30,000 copies have been made — on a sheet with no data validations, no conditional
formats, no merged ranges and no sparklines.

The cause is not the copying. `XLRangeBase.Clear(XLClearOptions.All)`, which `CopyTo` calls on its target
before writing, **creates a data validation covering the target range and immediately deletes it on every
call, whether or not the worksheet has any data validations**. That create-and-delete is ~85% of
`CopyTo`'s cost at 30,000 rows and is the whole of its growth.

A two-line guard removes it: `CopyTo` goes to **13 µs** per call and stops growing — ~30× on this
workload — with the full 11,603-test suite passing unchanged. That fix has been prototyped and measured;
this spec is what remains to do properly around it.

## Why this is worth a spec rather than just the two-line fix

Three reasons, in increasing order of importance.

1. **The guard hides a second bug rather than fixing it.** Something inside
   `DataValidations.Add`/`Delete` costs time proportional to the worksheet's used rows *with an empty
   collection*. Guarding the caller means nobody trips over it today; it does not mean it is not there.
   Any other caller of that pair has the same exposure.
2. **There is no regression test that would have caught this, and none that will catch its return.** The
   suite asserts behaviour, and this is a scaling property. Spec 12's benchmark caught it only because
   that spec happened to have an acceptance criterion about super-linearity.
3. **The blast radius is wider than `CopyTo`.** Anything that clears or copies ranges in a loop over a
   large sheet inherits the curve, and nothing in the API hints that it should.

## Measurements

`XLibur.Report.Benchmarks -- phases 30000`, .NET 10, Release. Ten buckets of 3,000 operations over a
30,000-row × 10-column sheet with no validations, formats, merges or sparklines. A rising column is the
quadratic term.

| rows so far | `CopyTo` 1×10 | `Clear()` 1×10 | `Clear(DataValidation)` | `Clear(Contents)` | `Range()` new address | 1-cell `CopyFrom` |
|---|---|---|---|---|---|---|
| 3,000 | 162 µs | 43 µs | 46 µs | 3.7 µs | 0.3 µs | 5.3 µs |
| 12,000 | 169 µs | 149 µs | 139 µs | 1.1 µs | 0.4 µs | 4.5 µs |
| 21,000 | 286 µs | 254 µs | 249 µs | 2.0 µs | 0.7 µs | 2.1 µs |
| 30,000 | **420 µs** | **372 µs** | **357 µs** | 1.8 µs | 0.2 µs | 1.9 µs |
| growth | 2.6× | 8.6× | 7.7× | flat | flat | flat |

Read as a set of eliminations:

- `Clear(DataValidation)` alone reproduces the entire curve, and is ~85% of `CopyTo`'s absolute cost at
  30,000 rows. **This is the defect.**
- `Clear(Contents)` skips every `All`-only step and is flat, so clearing *cells* is not the problem.
- Creating range objects at ever-new addresses is flat, so the range repository is **not** implicated.
  This was the first hypothesis — the repository holds one weak entry per address ever seen and prunes
  itself in an O(n) sweep — and it is wrong. Recorded because it is the plausible-sounding wrong answer
  and the next person will think of it too.
- Single-cell `CopyFrom` is flat, so the per-cell copy machinery is fine.

## The code

`XLibur/Excel/Ranges/XLRangeBase.cs`, in `Clear`:

```csharp
if (clearOptions.HasFlag(XLClearOptions.DataValidation))
{
    var validation = CreateDataValidation();      // adds a DV covering this range
    Worksheet.DataValidations.Delete(validation); // and removes it again
}
```

The round trip is not pointless. Going through `XLDataValidations.Add` is what runs
`SplitExistingRanges`, which carves the cleared area out of any validation that overlaps it. A plain
`DataValidations.Delete(range)` would instead drop whole validations that merely intersect, removing
validation from cells outside the cleared range. So the round trip is load-bearing **when a validation
intersects**, and pure waste otherwise.

Call path into it: `XLRangeBase.CopyTo` → `XLCellCopyHelper.CopyFromRange` → `targetRange.Clear()`
(`XLibur/Excel/Cells/XLCellCopyHelper.cs`, the `Clear` before the copy loop).

## Design

### Task 1 — Guard the round trip (the fix)

```csharp
if (clearOptions.HasFlag(XLClearOptions.DataValidation)
    && Worksheet.DataValidations.GetAllInRange(RangeAddress).Any())
{
    var validation = CreateDataValidation();
    Worksheet.DataValidations.Delete(validation);
}
```

`GetAllInRange` is a spatial-index query, so the common case becomes a lookup that finds nothing.
Behaviour is unchanged whenever a validation actually intersects the cleared range, which is the only
case the round trip exists for.

**Prototyped and measured**: `CopyTo` 420 µs → 13 µs at 30,000 rows and flat; 11,603 core tests and 447
report tests pass unchanged. The prototype was reverted rather than committed, so Task 1 is to land it
with the tests below.

### Task 2 — Find out why the round trip is O(used rows) at all

The guard means an empty worksheet no longer pays, but a worksheet **with** validations still pays
per clear, and the cost apparently scales with the sheet rather than with the number of validations. That
is the second bug, and it is the one worth understanding.

Where to look, in order of suspicion:

- `XLDataValidations.Delete(IXLDataValidation)` ends with
  `_dataValidationIndex.RemoveAll(e => ReferenceEquals(e.DataValidation, xlDataValidation))`. `RemoveAll`
  over a predicate on a spatial index is a candidate full scan, and the index's size is not obviously
  bounded by the validation count.
- `XLDataValidation`'s construction from a live range, and the event wiring
  (`RangeAdded`/`RangeRemoved`/`CoverageReplaced`) that `Add` and `Delete` attach and detach per call.
- `SplitExistingRanges` → `GetIntersectedRanges`, which is cheap on an empty index but is the path taken
  when there *are* validations.

Deliverable: either a fix, or a recorded finding naming the mechanism and its bound. A measurement that
says "this is O(k) in the number of validations, not O(rows)" is a perfectly good outcome — it would mean
the guard is the whole fix.

### Task 3 — A regression test for the scaling property

The gap that let this ship is that nothing tests shape. Add a test that asserts the *ratio* rather than
absolute time, which is what makes a timing test tolerable in CI:

- Clear (or copy) N times on a sheet of N rows, then again on a sheet of 4N rows.
- Assert the per-operation cost of the larger run is within a generous multiple — 2× or so — of the
  smaller. A quadratic term shows as 4×; noise on a shared runner does not reach 2×.
- Skip rather than fail if the smaller run is too fast to time reliably.

Two operations are worth covering: `Clear(All)` and `CopyTo`. Put it where slow tests are tolerated, and
mark it so a maintainer reading a failure knows it is a scaling assertion and not a correctness one.

### Task 4 — Sweep the neighbours

`Clear`'s other `All`-only steps were measured on collections that were **empty**, so they are unproven
rather than exonerated:

- `RemoveConditionalFormatting` — scans `Worksheet.ConditionalFormats`
- `ClearMerged` — scans `Worksheet.Internals.MergedRanges`
- `RemoveSparklines` — `SparklineGroups.GetSparklines` walks every group

Each is a linear scan of a workbook-level collection per clear, so each is quadratic in a loop *once that
collection is non-trivial*. Measure `Clear` on a sheet carrying a few hundred conditional formats, merges
and sparklines, and report. `ExpansionPhaseProbe` takes a new case in about five lines.

## Acceptance criteria

1. `Clear(XLClearOptions.DataValidation)` on a range of a worksheet with **no** data validations is O(1) in
   the worksheet's size: per-operation cost flat within noise from 3,000 to 30,000 rows.
2. `CopyTo` of a 1×10 range on a 30,000-row sheet costs **under 30 µs** and does not grow with the number
   of copies already made.
3. Clearing data validation still splits an intersecting validation exactly as it does today, asserted by
   a test that clears the middle of a validated range and checks the validation survives on both sides.
4. The full core suite (11,603 tests) passes unchanged, as does `XLibur.Report.Tests`.
5. A scaling regression test exists for both `Clear(All)` and `CopyTo`, asserting a ratio rather than an
   absolute time, and it fails if the guard from Task 1 is reverted.
6. Task 2 ends in either a fix or a written finding naming the mechanism; Task 4 ends in measurements for
   the three neighbours, and issues for any that are quadratic.

## Risks

- **A timing assertion in CI is a flake risk.** Mitigated by asserting a ratio with a wide margin and by
  skipping when the baseline is too fast to measure. If it flakes anyway, delete it and put the check in
  the benchmarks project with a documented threshold instead — a flaky test is worse than no test.
- **Guarding rather than fixing.** Task 1 alone leaves the underlying cost in place for workbooks that do
  have validations. That is a real improvement and an incomplete one, which is why Task 2 exists and why
  it should not be dropped once the headline number looks good.
- **`GetAllInRange` semantics.** The guard assumes it returns every validation intersecting the range.
  Confirm against the split behaviour: a validation that overlaps only partially must still be found, or
  the guard would skip a split that today happens.

## References

- `XLibur/Excel/Ranges/XLRangeBase.cs` — `Clear`, `CreateDataValidation`
- `XLibur/Excel/Cells/XLCellCopyHelper.cs` — `CopyFromRange`, and the `Clear` call this is all about
- `XLibur/Excel/DataValidation/XLDataValidations.cs` — `Add`, `Delete`, `ProcessRangeAdded`,
  `SplitExistingRanges`
- `XLibur.Report.Benchmarks/ExpansionPhaseProbe.cs` — the probe that produced the table above
  (`-- phases 30000`)
- Spec 12's finding 10 — the report engine's own workaround (copy by doubling), and the reason this was
  found at all
