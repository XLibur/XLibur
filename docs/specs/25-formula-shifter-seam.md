# Spec 25 — Narrow the formula shifter's fallback, and name its seam

**Area:** Architecture · **Correctness (masking)**
**Effort:** S (~2 days)
**Dependencies:** None. Touches no file any other open spec touches.
**Status:** Proposed. Smallest of the four architecture specs (22–25) and the least urgent — see
"Why this is ranked low".

## Goal

Stop `catch (Exception)` from silently routing a **bug** down the regex fallback path, and give the
two shifter implementations a named seam so a test can select one without provoking an exception.

## Why this spec exists

Reference shifting has two implementations:

| Implementation | File | Lines |
|---|---|---:|
| Parser-based (`ClosedXML.Parser`) | `XLibur/Excel/Cells/XLCellFormulaShifter.cs` | 433 |
| Regex-based (legacy) | `XLibur/Excel/Cells/XLCellFormulaShifter.Legacy.cs` | 437 |

They produce different answers for 9 of the 2,072 corpus rows. That is known, deliberate and pinned:
`FormulaShifterCorpusTests` carries a separate `LegacyExpected` column and asserts both paths on
every run.

The problem is not the divergence. It is **how the implementation is chosen**:

```csharp
try
{
    FormulaParser<object?, object?, ShiftPlan>.CellFormulaA1(parseable, plan, ShiftCollector.Instance);
}
catch (Exception)
{
    return axis == ShiftAxis.Row
        ? ShiftFormulaRowsLegacy(formulaA1, worksheetInAction, shiftedRange, shift)
        : ShiftFormulaColumnsLegacy(formulaA1, worksheetInAction, shiftedRange, shift);
}
```

`XLCellFormulaShifter.cs:119`, and the same shape at `:65` for the batch row overload.

The intent is documented on the class: formulas the parser rejects — external workbook references
such as `'[file.xlsx]Sheet'!A1` — fall back so that "no formula that shifted before stops shifting
now." That intent is sound. `catch (Exception)` is not the way to express it.

### What the broad catch actually covers

`CellFormulaA1` calls back into `ShiftPlan.Visit` → `TryShiftReference` → `Render`
(`XLCellFormulaShifter.cs:227-421`) for every reference in the formula. Any exception thrown by
**XLibur's own shift logic** — a `NullReferenceException`, an `IndexOutOfRangeException` from
`WithExtent`, an `InvalidOperationException` from the range repository — is caught here and turned
into a silent switch to a different algorithm that returns a plausible-looking answer.

A bug in the shift plan therefore presents as "this formula shifted slightly differently", not as a
crash. That is the worst possible failure mode for a correctness-critical path, and it is the class
of defect spec 05 already found once here: a deletion removing the tail of a range dropped a
surviving row (`A2:A8` with rows 5–9 deleted gave `A2:A3`, not `A2:A4`).

The codebase already knows the right exception type. `FormulaParser.cs:35`:

```csharp
catch (ParsingException ex)
{
    throw new ExpressionParseException(ex.Message);
}
```

`ParsingException` is `ClosedXML.Parser`'s own type, and it is exactly the signal the shifter means.

### Why this is ranked low

The divergence between the two implementations is already understood, counted and pinned. Both paths
are asserted against a 2,072-row corpus on every test run, and the nine differing rows are recorded
in the data rather than only in a comment. **This spec is a clarity and masking fix, not a
correctness fix** — it does not change any output for any formula that shifts correctly today.

It is worth doing because it is small, because it converts a silent degradation into a loud failure,
and because a named seam makes the legacy path directly testable. It is not worth doing before
specs 22, 23 or 24.

## Non-goals

- **Not deleting the regex path.** It is live, it is the documented fallback for external references,
  and it is pinned by the corpus.
- **Not reconciling the 9 divergences.** They are recorded behaviour. Changing them is a separate
  decision with its own compatibility argument.
- **Not writing a `CanParse` predicate.** Deciding up front whether `ClosedXML.Parser` will accept a
  formula means reimplementing its accept/reject rules, which is a second thing to keep in sync —
  the exact defect class specs 22 and 23 remove. The parser's own verdict stays the selector; only
  the *channel* it arrives on gets narrowed.
- **No public API change.**

## Current state

Verified against the tree at `d05b0753` (2026-08-23).

- `XLCellFormulaShifter.Shift` — `XLibur/Excel/Cells/XLCellFormulaShifter.cs:89-131`, `catch` at `:119`
- Batch row overload — `:53-88`, `catch` at `:65`
- `ShiftPlan` and its visitor — `:227-421`, the code the broad catch currently shields
- `XLCellFormulaShifter.Legacy.cs` — `ShiftFormulaRowsLegacy` (`:26`),
  `ShiftFormulaColumnsLegacy` (`:237`)
- `FormulaShifterCorpusTests` — `XLibur.Tests/Excel/Cells/FormulaShifterCorpusTests.cs`,
  2,072 rows from `XLibur.Tests/Resource/Other/FormulaShifterCorpus.tsv`
- Corpus regeneration:
  `dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile shiftercorpus`

## File structure

```
XLibur/Excel/Cells/XLCellFormulaShifter.cs         modified — narrowed catch, named seam
XLibur/Excel/Cells/XLCellFormulaShifter.Legacy.cs  unchanged
XLibur.Tests/Excel/Cells/FormulaShifterCorpusTests.cs  modified — fallback coverage
```

No new files. This spec makes an existing seam explicit rather than introducing one.

## The design

Two changes, in order of value:

**1. Narrow the catch.** `catch (ParsingException)` instead of `catch (Exception)`. Anything else
propagates, as it would anywhere else in the library.

**2. Name the seam.** Extract the fallback into a single private method with a name that says what
it is, so both call sites route through one place and a test can reach it:

```csharp
/// <summary>
/// Shifts a formula the parser cannot read, using the regex implementation.
/// </summary>
/// <remarks>
/// Reached only for formulas <see cref="ClosedXML.Parser"/> rejects — external workbook references
/// such as <c>'[file.xlsx]Sheet'!A1</c>. The two implementations disagree on 9 of the 2,072 rows in
/// <c>FormulaShifterCorpus.tsv</c>, all of them the tail-deletion clamp; the corpus pins both.
/// </remarks>
internal static string ShiftUnparseable(string formulaA1, XLWorksheet worksheetInAction,
    XLRange shiftedRange, int shift, ShiftAxis axis);
```

`internal` rather than `private` so `FormulaShifterCorpusTests` can call the fallback directly
instead of through `ShiftFormulaRowsLegacy` / `ShiftFormulaColumnsLegacy`, which is what makes the
seam a real test surface rather than a notional one.

## Global constraints

- Warnings are errors; nullable enabled.
- Branch per task; never commit to main. Commit prefix `fix:` for task 2, `refactor:` for task 3,
  `test:` for tasks 1 and 4.
- No compound shell commands (`&&`, `;`) in agent tool calls.
- Build: `dotnet build XLibur/XLibur.csproj -c Release -v q`
- Corpus tests: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/FormulaShifterCorpusTests/*"`
- Use `--treenode-filter`, never `--filter`. Never filter at solution level.

## Work plan

| # | Task | Size | Gate |
|---|---|---|---|
| 1 | Prove the fallback is reachable and pin which formulas take it | S | New test green; the taking-set recorded |
| 2 | Narrow `catch (Exception)` to `catch (ParsingException)` | S | Full corpus green — 2,072 × 2 assertions |
| 3 | Name the seam; route both call sites through it | S | Suite green |
| 4 | Add external-reference rows to the corpus | S | Corpus regenerates clean |

Task 2 is the one that carries risk. Task 1 exists to size that risk before taking it.

---

### Task 1 — Prove the fallback is reachable, and pin what takes it

Before narrowing the catch, establish which formulas actually reach the fallback today. If any
formula reaches it for a reason **other** than a parse failure, narrowing the catch will change that
formula's behaviour, and this task is what finds it.

**Files:**
- Modify: `XLibur.Tests/Excel/Cells/FormulaShifterCorpusTests.cs`

**Interfaces:**
- Produces: `The_parser_rejects_exactly_the_formulas_the_fallback_is_documented_for`.

- [ ] **Step 1: Write a test that asserts the parser rejects external references, and accepts the corpus**

```csharp
/// <summary>
/// The fallback exists for formulas ClosedXML.Parser rejects. This pins both halves of that claim:
/// an external workbook reference is rejected, and every formula in the corpus is accepted — so
/// narrowing the catch in spec 25 task 2 cannot silently reroute a corpus row.
/// </summary>
[Test]
[Arguments("='[file.xlsx]Sheet'!A1", false)]
[Arguments("=SUM('[book.xlsx]Data'!A1:A5)", false)]
[Arguments("=A1+B2", true)]
[Arguments("=SUM(A1:A5)", true)]
[Arguments("=Sheet2!A1", true)]
public async Task The_parser_accepts_only_what_the_fallback_is_not_for(string formula, bool parseable)
{
    await Assert.That(TryParse(formula)).IsEqualTo(parseable);
}

/// <summary>Every corpus formula must parse, or the corpus is silently testing the regex path.</summary>
[Test]
[MethodDataSource(nameof(Corpus))]
public async Task Every_corpus_formula_is_accepted_by_the_parser(CorpusCase test)
{
    await Assert.That(TryParse(test.Formula)).IsTrue();
}

private static bool TryParse(string formula)
{
    var text = formula.Length > 0 && formula[0] == '=' ? formula[1..] : formula;
    text = FormulaTransformation.ProtectStructuredRefColons(text, out _);
    try
    {
        FormulaParser<object?, object?, object?>.CellFormulaA1(text, null, ProbeFactory.Instance);
        return true;
    }
    catch (ClosedXML.Parser.ParsingException)
    {
        return false;
    }
}
```

`ProbeFactory` is a do-nothing `IAstFactory` used only to drive the parser. If writing one is more
than a few lines, use the shifter's own `ShiftCollector.Instance` with a throwaway `ShiftPlan`
instead — the point is only whether the parse throws.

- [ ] **Step 2: Run it**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/FormulaShifterCorpusTests/*"`
Expected: PASS, including all 2,072 rows of `Every_corpus_formula_is_accepted_by_the_parser`.

**If any corpus row fails to parse**, the corpus's `Expected` column was recorded from the regex
path, not the parser path, for that row. Record which rows and stop — task 2's risk assessment
changes and this spec needs updating before proceeding.

- [ ] **Step 3: Commit**

```bash
git add XLibur.Tests/Excel/Cells/FormulaShifterCorpusTests.cs
git commit -m 'test(shift): pin which formulas the parser rejects (spec 25 task 1)'
```

---

### Task 2 — Narrow the catch

**Files:**
- Modify: `XLibur/Excel/Cells/XLCellFormulaShifter.cs:65`, `:119`

- [ ] **Step 1: Add the using**

```csharp
using ClosedXML.Parser;
```

Already present at the top of `XLCellFormulaShifter.cs` — confirm rather than duplicate.

- [ ] **Step 2: Narrow both catches**

At `:119`, replace `catch (Exception)` with:

```csharp
        // ParsingException specifically, not Exception: the fallback exists for formulas the parser
        // cannot read, and nothing else. A bug in ShiftPlan used to be caught here and answered with
        // a plausible-looking result from the other implementation instead of surfacing.
        catch (ParsingException)
```

At `:65`, the batch row overload, the same narrowing. Keep its existing comment about degrading to
per-run application and add the same rationale.

- [ ] **Step 3: Run the full corpus — 2,072 rows, both paths**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/FormulaShifterCorpusTests/*"`
Expected: PASS. Every row must produce the same output as before, because task 1 established that
every corpus formula parses, so none of them was reaching the fallback.

- [ ] **Step 4: Run the whole suite — this is where a masked bug would surface**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`

Expected: PASS.

**If a test now throws where it previously passed, you have found a real bug that the broad catch was
hiding.** Do not widen the catch back. Record the formula and the exception, fix the underlying
defect in `ShiftPlan`, and note it in this spec's Results section — that finding is worth more than
the rest of this spec.

Pay attention to the shift-adjacent suites:

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*Shift*/*"`

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/Cells/XLCellFormulaShifter.cs
git commit -m 'fix(shift): fall back only on a parse failure, not on any exception (spec 25 task 2)'
```

---

### Task 3 — Name the seam

**Files:**
- Modify: `XLibur/Excel/Cells/XLCellFormulaShifter.cs`

**Interfaces:**
- Produces: `XLCellFormulaShifter.ShiftUnparseable(string, XLWorksheet, XLRange, int, ShiftAxis) → string`.

- [ ] **Step 1: Extract the fallback**

```csharp
    /// <summary>
    /// Shifts a formula the parser cannot read, using the regex implementation in
    /// <c>XLCellFormulaShifter.Legacy.cs</c>.
    /// </summary>
    /// <remarks>
    /// This is the other side of the shifter's one seam. It is reached only for formulas
    /// <see cref="ClosedXML.Parser"/> rejects — external workbook references such as
    /// <c>'[file.xlsx]Sheet'!A1</c>. The two implementations disagree on 9 of the 2,072 rows in
    /// <c>FormulaShifterCorpus.tsv</c>, all of them the tail-deletion clamp described on
    /// <see cref="FormulaShifterCorpusTests"/>; the corpus pins both columns so neither can drift.
    /// <para>
    /// Internal rather than private so a test can exercise this adapter directly, instead of having
    /// to provoke a parse failure to reach it.
    /// </para>
    /// </remarks>
    internal static string ShiftUnparseable(string formulaA1, XLWorksheet worksheetInAction,
        XLRange shiftedRange, int shift, ShiftAxis axis)
        => axis == ShiftAxis.Row
            ? ShiftFormulaRowsLegacy(formulaA1, worksheetInAction, shiftedRange, shift)
            : ShiftFormulaColumnsLegacy(formulaA1, worksheetInAction, shiftedRange, shift);
```

`ShiftAxis` is currently a `private enum` at `XLCellFormulaShifter.cs:133`. Widen it to `internal`
so it can appear in an `internal` signature — otherwise the build fails with an inconsistent
accessibility error.

- [ ] **Step 2: Route the `Shift` catch through it**

```csharp
        catch (ParsingException)
        {
            return ShiftUnparseable(formulaA1, worksheetInAction, shiftedRange, shift, axis);
        }
```

- [ ] **Step 3: Leave the batch overload's catch as it is**

The batch row overload at `:65` does something different — it decomposes the deletion map into runs
and applies the legacy shifter once per run, because there is no batch regex shifter. That is not
the same operation and must not be folded into `ShiftUnparseable`. Add a comment saying so, so a
later reader does not "tidy" the two together:

```csharp
            // Not ShiftUnparseable: there is no batch regex shifter, so the map is decomposed into
            // runs and the single-block fallback is applied once per run.
```

- [ ] **Step 4: Build and run the suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/Cells/XLCellFormulaShifter.cs
git commit -m 'refactor(shift): name the parser/regex seam (spec 25 task 3)'
```

---

### Task 4 — Give the fallback its own corpus rows

The corpus covers 2,072 formulas the parser accepts. It covers **none** that take the fallback, so
the live regex path — the one that actually runs in production for external references — is
exercised only by being called directly with formulas that did not need it.

**Files:**
- Modify: `XLibur.Tests/Excel/Cells/FormulaShifterCorpusTests.cs`

- [ ] **Step 1: Write a test that goes through the front door**

```csharp
/// <summary>
/// External workbook references are what the fallback exists for, and until spec 25 nothing
/// exercised them through Shift itself — the corpus calls the two implementations directly, and
/// every one of its formulas parses. These go in the front door and come out the regex path.
/// </summary>
[Test]
[Arguments("='[file.xlsx]Sheet'!A1", 5, 9, -5)]
[Arguments("=SUM('[book.xlsx]Data'!A1:A20)", 3, 4, 2)]
[Arguments("='[file.xlsx]Sheet'!A1+B2", 1, 2, 3)]
public async Task An_external_reference_shifts_through_the_fallback(
    string formula, int first, int last, int shift)
{
    using var wb = new XLWorkbook();
    var shiftedSheet = (XLWorksheet)wb.AddWorksheet("Sheet1");
    var range = (XLRange)shiftedSheet.Range(first, 1, last, XLHelper.MaxColumnNumber);

    var throughShift = XLCellFormulaShifter.ShiftFormulaRows(
        formula, shiftedSheet, range, shift);

    var throughFallback = XLCellFormulaShifter.ShiftUnparseable(
        formula, shiftedSheet, range, shift, XLCellFormulaShifter.ShiftAxis.Row);

    // Reaching the same answer both ways is what proves Shift routed here rather than succeeding
    // on the parser path with a different result.
    await Assert.That(throughShift).IsEqualTo(throughFallback);
}
```

- [ ] **Step 2: Run it**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/FormulaShifterCorpusTests/*"`
Expected: PASS.

If the two differ, `Shift` is **not** routing external references to the fallback — the parser is
accepting them and shifting them itself. That would mean the fallback is dead code for its
documented purpose, which is a finding worth recording and would change this spec's conclusion.

- [ ] **Step 3: Consider regenerating the corpus with these rows included**

If the extractor can emit external-reference rows, regenerate:

```
dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile shiftercorpus
```

It reports divergences between the two columns on stderr. Any **new** divergence beyond the recorded
9 is a finding, not a re-baseline. If the extractor cannot produce these rows, the explicit test
above is sufficient — say so in the test's summary.

- [ ] **Step 4: Commit**

```bash
git add XLibur.Tests/Excel/Cells/FormulaShifterCorpusTests.cs XLibur.Tests/Resource/Other/FormulaShifterCorpus.tsv
git commit -m 'test(shift): exercise the fallback through Shift, not just directly (spec 25 task 4)'
```

---

## Acceptance criteria

1. Neither `catch (Exception)` remains in `XLCellFormulaShifter.cs`. Gate:
   `grep -n 'catch (Exception)' XLibur/Excel/Cells/XLCellFormulaShifter.cs` returns nothing.
2. The fallback is reached through one named method, `ShiftUnparseable`, from the single-block path.
3. The batch overload's per-run decomposition is documented as deliberately *not* routed through it.
4. All 2,072 corpus rows pass on both the parser and legacy columns, unchanged.
5. At least three external-reference formulas are exercised through `Shift` itself.
6. Full suite green on net8.0 and net10.0.
7. No public API change.
8. Any exception unmasked by task 2 is recorded in a Results section, whether or not it is fixed
   under this spec.

## Conflicts

None. No other spec in `docs/specs/` touches `XLibur/Excel/Cells/XLCellFormulaShifter*.cs`.

Spec 05 rewrote reference shifting onto `ClosedXML.Parser` and is **done** — this spec is a follow-on
to that work, tightening the fallback 05 left in place. Spec 05's Results section is worth reading
before starting: it records the tail-deletion bug the parser path fixed, which is the same nine rows
the corpus still shows the regex path getting wrong.
