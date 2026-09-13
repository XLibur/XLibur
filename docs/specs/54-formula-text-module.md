# Spec 54 — Formula text gets one module, and a refused formula gets one meaning

**Area:** Architecture · **Defects (4, executed)**
**Effort:** M (~4–5 days)
**Dependencies:** None hard. **Must land before spec 55**, which builds sheet rename and sheet delete
on this module. Soft file overlaps with specs 42, 56 and 04 — see *Conflicts*.
**Status:** Proposed. From the 2026-09-13 architecture review (round 4). Every design decision below
was taken by the owner in a design interview; the table in *Decisions* records them. Also recorded:
`docs/adr/0002-refused-formula-never-rewritten.md`, and the terms *formula text*, *refused formula*
and *future function* in `CONTEXT.md` at the repo root.

Line numbers are from `d68dffdc` (2026-09-13). Verify them against current code before editing.

---

## Goal

One module, `FormulaText`, is the only code in XLibur that hands formula text to `ClosedXML.Parser`.
It does everything that must happen to the text before the parser sees it and after the parser is
done, and it turns the parser's refusal into one result. A caller makes one decision only: what a
refused formula means to it.

## Why this spec exists

Since round 3, nine of the ten library fixes landed on the formula pipeline, each fixing one syntax
form on one path:

| PR | Fix | Where it had to go |
|---|---|---|
| #474 | Parser fork 3.0.0 | `FormulaParser.cs` and five visitors |
| #475 | A defined name holding a bang reference evaluates | `FormulaParser.cs`, `XLDefinedName.cs` |
| #476 | Converting a formula keeps its surrounding whitespace | `XLCellFormula.cs` |
| #477 | A refused formula throws `ExpressionParseException` from `CopyTo` and `FormulaR1C1` | `FormulaTransformation.cs` |
| #480, #481 | `@` and the space operator evaluate | `CalculationVisitor.cs` |
| #443 | A defined name keeps its formula through a shift | `XLDefinedName.cs`, `FormulaReferences.cs` |

Each fix reached one call site. There are eight places where XLibur hands text to the parser, and
each one re-implements the steps around the call:

| # | Path | Call site | Colon protection | Refusal policy |
|---|---|---|---|---|
| 1 | Evaluate | `CalcEngine/FormulaParser.cs:31-32` | **no** | wraps as `ExpressionParseException` |
| 2 | SUBTOTAL/AGGREGATE nesting check | `CalcEngine/CalcContext.cs:317` | **no** | **none — raw `ParsingException`** |
| 3 | Shifter, single block | `Cells/XLCellFormulaShifter.cs:63` | yes (`:59`) | regex fallback (spec 25) |
| 4 | Shifter, multi block | `Cells/XLCellFormulaShifter.cs:123` | yes (`:107`) | regex fallback (spec 25) |
| 5 | Rewrite: sheet rename, name copy, future-function remap | `Visitors/FormulaTransformation.cs:58` | yes | **none — raw `ParsingException`**, except when reached from `FixFutureFunctions` (`:38-48`), which swallows every exception |
| 6 | A1 ↔ R1C1 conversion: `FormulaR1C1`, copy, shared formulas on load | `Visitors/FormulaTransformation.cs:88` | yes | wraps (since #477) |
| 7 | Defined-name references | `Visitors/FormulaReferences.cs:52` | **no** | returns `false` |
| 8 | Shift extent | `Visitors/FormulaExtent.cs:46` | yes (`:47`) | reports the whole sheet |

Every row was read at `d68dffdc`; `grep -rn "FormulaParser<\|FormulaConverter\." XLibur/` returns
exactly these lines, plus the `Convert` delegate at row 6.

Four facts are implemented at each call site that needs them, and the call sites disagree:

1. **What a refusal means.** Six policies across eight call sites, and two of them have no policy at
   all.
2. **The colon inside a table column name.** `Table[Start: Date]` must reach the parser with the colon
   hidden, or the parser reads a range operator. Five call sites protect it; the evaluator, the name
   reference collector and the SUBTOTAL check do not.
3. **The future-function prefix.** Evaluation strips `_xlfn.` with a culture-sensitive, case-sensitive
   `StartsWith` (`FormulaParser.cs:313`) and `_xlws.` ordinally (`:320`). The trie that decides whether
   to add a prefix matches case-insensitively (`FormulaTransformation.cs:245-263`).
   `RenameFunctionsVisitor` adds the prefix. The SUBTOTAL visitor matches function names in any case.
4. **The leading `=`.** `GetAst` strips it (`FormulaParser.cs:25-26`). `NameNode.GetValue`
   (`AstNode.cs:429-431`) and `Lookup.cs:796-798` put it back, then pass the text to `GetAst`, which
   strips it again. The comment that justifies adding it is stale.

### The four defects, executed

All four were executed against a scratch build at `d68dffdc` before this spec was written.

**D49 — renaming a sheet throws, after part of the rename has happened.**

```
Sheet2!A1.FormulaA1 = "'[Book2.xlsx]Sheet1'!A1"     accepted
Sheet2!B1.FormulaA1 = "Sheet1!A1+1", Sheet1!A1 = 5  ->  B1 = 6
wb.Worksheet("Sheet1").Name = "Data"                ->  throws ClosedXML.Parser.ParsingException
afterwards: sheet Name is "Sheet1", B1 text is "Sheet1!A1+1", B1 value is #REF!
```

The mechanism (read in code): the `Name` setter (`XLWorksheet.cs:273`) calls `XLWorksheets.Rename`,
which removes the old dictionary key and adds the new one (`XLWorksheets.cs:247-248`), then notifies
listeners in order — the calc engine, every sheet's cells collection, every defined name
(`:256-277`). The cells collection reaches `SafeModifyA1` (`XLCellFormula.cs:460`), which has no
`catch`, so the parser's exception escapes. By then the calc engine has renamed `Sheet1` to `Data` in
its dependency tree, and `_name = value` (`XLWorksheet.cs:274`) never runs. B1 evaluates to `#REF!`
because the engine and the sheet no longer agree on the sheet's name. The exception type belongs to a
dependency and XLibur documents it nowhere. **This is D22's shape: mutate, then throw.**

The refused text needs no unusual input: `FixFutureFunctions` accepts it by design, and its comment
(`FormulaTransformation.cs:44-46`) names this exact form — an external reference typed the way the
formula bar shows it.

**D50 — an upper-case future-function prefix evaluates to `#NAME?`.** `_xlfn.CONCAT("a","b")` gives
`"ab"`. `_XLFN.CONCAT("a","b")` gives `#NAME?`.

**D51 — SUBTOTAL over a refused formula throws the parser's exception.**
`A1.FormulaA1 = "'[Book2.xlsx]Sheet1'!A1+SUBTOTAL(9,B5)"` is accepted.
`C1.FormulaA1 = "SUBTOTAL(9,A1:A2)"`. Reading `C1.Value` throws a raw `ParsingException` about A1's
text. The caller read C1.

**D52 — a table column whose name contains a colon evaluates to `#REF!`.**
`SUM(Table1[Start: Date])` over the values 3 and 4 gives `#REF!`, not 7, before and after a save and
reload. The text survives the round trip. Only evaluation fails, because evaluation is one of the three
call sites that skip the colon protection. This also settles an open question from the review: parser
3.1.0 does not handle the colon itself, so the protection is needed and must apply everywhere.

### Read, not executed

- **No test names** `FormulaReferences`, `RenameRefModVisitor`, `RenameFunctionsVisitor`,
  `SafeModifyA1` or `FixFutureFunctions`. The shifter has a 2,072-row corpus; no other path has one.
- **A conditional format's `Equals` and `GetHashCode` run the R1C1 converter**
  (`XLConditionalFormat.cs:33-40`, `:74-78`). Since #477 the converter throws on a refused formula, so a
  conditional format that holds one may throw from `GetHashCode`. Task 1 executes this.
- `DefaultFormulaVisitor` has no subclasses and no references.

## Decisions (owner, 2026-09-13)

| # | Decision |
|---|---|
| Q3 | All eight call sites move behind the module. INDIRECT's hand-written reference parser (`Lookup.cs:700-779`) stays out and is recorded as a follow-up. |
| Q4 | The module returns a refusal as a result. The caller chooses the fallback. Only the public edges — evaluation, `FormulaR1C1`, `CopyTo` — turn a refusal into `ExpressionParseException`. |
| Q5 | `FormulaA1` keeps accepting a refused formula. |
| Q6 | Sheet rename and copy leave a refused formula's text unchanged. No regex fallback outside the shifter, no validate-then-throw. Recorded as ADR 0002. |
| Q7 | The module owns the future-function prefix in both directions, case-insensitively. |
| Q18 | The module is `FormulaText`, in `Excel/CalcEngine/`. It replaces `FormulaTransformation`; it does not wrap it. |
| Q19 | The SUBTOTAL nesting check treats a refused formula as not calling SUBTOTAL, so the cell's value counts. No guessing from the text. |
| Q20 | The regression gate is a corpus: syntax form × path, in the shape of `FormulaShifterCorpus.tsv`. |
| Q36 | Cost gate: structural-edit and load benchmarks, three runs each, medians compared. More than 10% regression gives the task authority to revert. |

## Non-goals

- **INDIRECT's reference parser.** It parses a reference, not a formula. Follow-up.
- **The shifter's regex fallback.** It stays, exactly as spec 25 left it. This spec changes only how the
  shifter learns that the parser refused its text.
- **Sheet delete, the rename gate on defined names, and every other holder of sheet names.** Spec 55.
- **Classifying evaluation failures.** Spec 56. A refused formula is one of its kinds; this spec
  supplies the refusal that 56 classifies.
- **What `ExpressionParseException` means to a public caller.** Unchanged.
- **Display text.** `FormulaA1` keeps returning the stored text, `_xlfn.` included. Stripping it for
  display is a separate decision.
- **The data-table placeholder text** `{TABLE(A1,}` (`XLCellFormula.cs:34`, `docs/specs/todos.md`).

## File structure

| File | Change |
|---|---|
| `XLibur/Excel/CalcEngine/FormulaText.cs` | **New.** The module. |
| `XLibur/Excel/CalcEngine/Visitors/FormulaTransformation.cs` | **Deleted.** Its contents move into `FormulaText`. |
| `XLibur/Excel/CalcEngine/FormulaParser.cs` | Calls the module; the prefix strip at `:313-324` moves out. |
| `XLibur/Excel/CalcEngine/CalcContext.cs` | The SUBTOTAL check (`:309-324`) calls the module. |
| `XLibur/Excel/CalcEngine/Visitors/FormulaReferences.cs`, `FormulaExtent.cs` | Call the module. |
| `XLibur/Excel/Cells/XLCellFormulaShifter.cs` | Both parse calls go through the module; the refusal selects the spec 25 fallback. |
| `XLibur/Excel/Cells/XLCellFormula.cs`, `XLCell.cs`, `Ranges/XLRangeBase.cs`, `IO/WorksheetSheetDataReader.cs`, `DefinedNames/XLDefinedName.cs` | Call sites renamed from `FormulaTransformation` to `FormulaText`. `XLCellFormula.cs:460` handles the refusal. |
| `XLibur/Excel/CalcEngine/AstNode.cs`, `Functions/Lookup.cs` | Stop re-adding `=`. |
| `XLibur/Excel/CalcEngine/DefaultFormulaVisitor.cs` | **Deleted**, if still unreferenced. |
| `XLibur.Tests/Resource/Other/FormulaTextCorpus.tsv` | **New.** The corpus. |
| `XLibur.Tests/Excel/CalcEngine/FormulaTextCorpusTests.cs`, `FormulaTextTests.cs` | **New.** |

## The design

### 1. What moves into the module

- **The leading `=`.** One rule, in one place. The two call sites that add it back stop doing so.
- **The colon protection.** Applied on every path, including the three that skip it today. D52 is fixed
  by this move alone.
- **The future-function prefix, both directions, case-insensitively.** On the way in, typed text gains
  its prefixes (`acot(A5)` → `_xlfn.ACOT(A5)`). On the way to evaluation, `_xlfn.` and a nested
  `_xlws.` are stripped whatever their case. D50 is fixed by this move.
- **The parse call itself**, in A1 and R1C1, with the caller's factory or visitor.
- **Rewrite** (`FormulaConverter.ModifyA1`) and **convert** (`ToR1C1`, `ToA1`).
- **The refusal.** The parser signals a refusal by throwing `ParsingException`. The module catches it
  in exactly one place and returns it as a value. No other file in XLibur catches `ParsingException`.

### 2. The interface — a sketch

This is the shape, not the final signatures. Task 2 settles the signatures against the callers.

```csharp
internal static class FormulaText
{
    // Walks formula text with a parser callback factory.
    // False, with the refusal, when the parser refuses the text.
    internal static bool TryWalk<TScalar, TNode, TContext>(
        string text, TContext context, IAstFactory<TScalar, TNode, TContext> factory,
        FormulaNotation notation, out TNode? root, out FormulaRefusal refusal);

    internal static bool TryRewrite(
        string text, string sheetName, Point origin, FormulaModifier modifier,
        out string rewritten, out FormulaRefusal refusal);

    internal static bool TryConvert(
        string text, Point origin, FormulaNotation to,
        out string converted, out FormulaRefusal refusal);

    // A refused formula comes back unchanged: it cannot contain a future function to prefix.
    internal static string AddFuturePrefixes(string typed, string sheetName, Point origin);

    internal static bool TryStripFuturePrefix(ReadOnlySpan<char> name, out ReadOnlySpan<char> bare);
}

internal readonly record struct FormulaRefusal(string Text, string Message)
{
    // The public edges' translation, written once.
    internal ExpressionParseException ToException() => ...;
}
```

`FixFutureFunctions` becomes `AddFuturePrefixes`. Its `catch (Exception)` narrows to the refusal, the
same narrowing spec 25 made in the shifter: a defect inside the remap stops being answered with the
unchanged text.

### 3. What each caller keeps

| Caller | Meaning of a refusal |
|---|---|
| Evaluation (`FormulaParser.GetAst`) | `ExpressionParseException` — unchanged |
| `FormulaR1C1` get, `CopyTo`, shared formulas on load | `ExpressionParseException` — unchanged since #477 |
| Sheet rename, in a cell formula | **The text is unchanged**, and the rename carries on. *Behaviour change* — D49 |
| Defined-name copy and rename | The text is unchanged. Today this goes through the stored `_isFormulaUnderstood` flag; it now comes from the refusal |
| Future-function prefixing on write | The text is unchanged. Today every exception is swallowed; now only a refusal is |
| Defined-name references | "Not understood" — unchanged |
| Shift extent | The whole sheet — unchanged |
| Shifter | The spec 25 regex fallback — unchanged |
| SUBTOTAL nesting check | **Does not call SUBTOTAL.** *Behaviour change* — D51 |

### 4. The corpus

`XLibur.Tests/Resource/Other/FormulaTextCorpus.tsv`, read by `FormulaTextCorpusTests`, in the shape of
`FormulaShifterCorpus.tsv`.

**Rows** are syntax forms. At least: a plain reference; a sheet reference; a quoted sheet reference; a
3D reference; a bang reference `!A1`; a name; a sheet-qualified name `Sheet1!Local`; a bang name; a
structured reference; a column name containing a colon; an external reference `[1]Sheet1!A1`; an
external reference in formula-bar form (refused); a function call; `_xlfn.` lower case; `_XLFN.` upper
case; `_xlfn._xlws.FILTER`; `@`; the space operator; the union operator; the range operator; an error
literal; a sheet-prefixed `#REF!`; an array literal; leading and trailing whitespace; a leading `=`;
empty text; text of whitespace only; unparseable text.

**Columns** are paths: `evaluate`, `references`, `extent`, `shift` (insert two rows above row 2),
`rename` (`Sheet1` → `Data`), `to_r1c1` (at C3), `add_prefix`, `calls_subtotal`. Each cell is the
expected text or value, `REFUSED`, or `THROWS <type>`.

**The fixture** behind `evaluate` is built in code, not loaded: `Sheet1` with numbers in `A1:C3`, a
table `Table1` with columns `Name` and `Start: Date`, a workbook-scoped name and a sheet-scoped name.
Task 1 writes it once and every row uses it.

## Global constraints

- **No public API change.** `PublicAPI.Shipped.txt` and `PublicAPI.Unshipped.txt` are untouched.
  `ExpressionParseException` stays the public face of a refusal.
- **ADR 0002 holds.** No new regex over formula text. The shifter's named fallback from spec 25 stays
  the only one.
- **Accepted text pays nothing new.** On text the parser accepts, the module must not allocate beyond
  what the call site allocates today. The colon protection already rents a buffer only when a
  qualifying colon exists; keep it that way.
- The ground rules in `TASKLIST-architecture-deepening-4.md` §6 apply.

## Work plan

Tasks are sequential. One owner, one branch: `refactor/54-formula-text`.

### Task 1 — Pin today's behaviour; the four defects land red

1. Write `FormulaTextCorpus.tsv` and `FormulaTextCorpusTests` against **today's** behaviour. Every
   cell passes.
2. Write `FormulaTextTests` with one test per defect, D49–D52, asserting the **correct** behaviour.
   They fail. The commit message names all four.
3. Execute the conditional-format lead: a conditional format whose formula is refused, put in a
   `HashSet`. Record in *Results* whether it throws. If it does, add it as a fifth red test.

**Gate:** the corpus is green; the defect tests fail for the stated reason and no other.

### Task 2 — `FormulaText` with the refusal, inert

The module and `FormulaRefusal`, with unit tests calling it directly — one per corpus row, per
operation. Nothing calls it yet.

**Gate:** the new unit tests pass; the rest of the suite is unchanged.

### Task 3 — Evaluation, conversion and rewrite go through it

Route call sites 1, 5 and 6. Delete `FormulaTransformation.cs`. Move the prefix strip out of
`FormulaParser.cs`. D50 and D52 turn green. Any corpus cell that changes is named in the commit
message with the reason.

### Task 4 — Name references, extent and the shifter

Route call sites 3, 4, 7 and 8. The shifter reaches its spec 25 fallback from the refusal, not from a
`catch`. `FormulaShifterCorpus.tsv` does not change. If it does, stop: that is a finding, not
something to fix up.

### Task 5 — The SUBTOTAL check and the rename

Route call site 2; D51 turns green. The cell-formula rename (`XLCellFormula.cs:460`) handles the
refusal by leaving the text unchanged; D49 turns green. Add a test: after the rename in D49, the
sheet's `Name`, the workbook's lookup by name and the calc engine all agree.

### Task 6 — The `=` rule and the dead visitor

Remove the two call sites that re-add `=` (`AstNode.cs:429-431`, `Lookup.cs:796-798`). Delete
`DefaultFormulaVisitor.cs` if `grep` still finds no reference to it.

### Task 7 — Cost

Take three runs each, before and after, and compare medians:

- `dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- structural` — the
  structural-edit profile, which spec 05 found spends 68% of its time shifting formulas;
- `-- --filter '*XLiburReadBenchmarks*'` — load, which converts shared formulas;
- `-- --filter '*FormulaEvaluationBenchmarks*'` — the evaluation parse path.

**Revert authority:** a regression of more than 10% on any median lets this task revert the routing
that caused it, and escalate. The defect fixes can then land as local fixes at their call sites. The
benchmark machine has roughly 40% run-to-run variance, so one run proves nothing.

### Task 8 — Changelog

`fix:` entries under `## Unreleased` for D49–D52. The rename no longer throws, and it leaves a refused
formula unchanged. SUBTOTAL counts a refused formula's cell instead of throwing. Neither change breaks
a caller, so neither is marked `!`.

## Acceptance criteria

| # | Criterion |
|---|---|
| 1 | `grep -rn "FormulaParser<\|FormulaConverter\." XLibur/` matches only `FormulaText.cs` |
| 2 | `grep -rn "catch (ParsingException" XLibur/` matches exactly one line, in `FormulaText.cs`, and no `catch (Exception)` surrounds a parse |
| 3 | `FormulaTransformation.cs` is gone |
| 4 | D49–D52 are green. The corpus is green, and every cell that changed is named in a commit message |
| 5 | `FormulaShifterCorpus.tsv` is unchanged |
| 6 | No public API change |
| 7 | Every benchmark median is within 10% of its baseline |
| 8 | All four test projects are green on net8.0 and net10.0 |

## Conflicts

| Spec | Shared ground | Resolution |
|---|---|---|
| **55** | 55 rewrites through `FormulaText` | **Hard. 54 first.** |
| **56** | `XLCell.cs` — 54 edits the prefixing calls (`:605`, `:631`, `:910`), 56 the catch sites (`:315`, `:343`) | Soft. Different regions; either order. 56 maps `ExpressionParseException` as "refused formula" whether 54 has landed or not |
| **42** | `XLRangeBase.cs` — the array-formula setter at `:114` calls `FixFutureFunctions`, and 42 rewrites that setter's invalidation | Soft. Trivial rebase either way |
| **04, 08** | `CalcContext.cs` | Soft. 54 touches only the SUBTOTAL helper |
| **25** | `XLCellFormulaShifter.cs` | Done. 54 changes the two parse calls, not the fallback |
| **30, 32, 37** | — | None. `FunctionDefinition.cs` and `SignatureAdapter.cs` are untouched |
