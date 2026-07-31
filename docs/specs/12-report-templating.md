# Spec 12 — Report Templating (`XLibur.Report` package)

**Area:** Feature · Arch | **Effort:** L | **Status:** Implemented (July 2026) on
`feat/spec-12-report-templating` — all eleven tasks landed; see **Still open** under Results for the
gaps left inside them, the largest being the un-ported upstream gauge corpus.

## Summary

Add a report-templating engine to the XLibur repository as a new `XLibur.Report` package:
author a report as an ordinary `.xlsx` template (placeholder expressions, named ranges, tag
markers), bind .NET data to it, and generate the finished workbook. The architecture is a port
of ClosedXML.Report (MIT, same lineage as this fork), which already proves the core model —
defined-name-bound repeating ranges, a cell-text `<<tag>>` system, and buffer-sheet rendering —
against only the public `IXL*` API surface. On top of that proven core, this spec replaces the
expression engine with Scriban behind a pluggable abstraction, bridges XLibur's calc-engine
function registry into template expressions (`{{ SUM(items.Price) }}`), and closes the three
gaps that ClosedXML.Report never could: **charts, pivot tables, and images that survive range
expansion**, plus conditional-row tags and conditional-formatting ranges that extend instead of
duplicating.

Grounding (research, July 2026):

- ClosedXML.Report 0.2.12 (May 2025) uses **System.Linq.Dynamic.Core** for `{{ }}` expressions,
  drives repetition off Excel defined names bound to `IEnumerable` variables, renders through a
  `VeryHidden` buffer sheet then splices back, and registers `<<tags>>` via a public static
  `TagsRegister`. It touches **no ClosedXML internals** — the port is a namespace/API retarget,
  not a rewrite. Its test suite is ~111 golden-file tests (template + gauge workbook pairs,
  semantic diff).
- Upstream is effectively unmaintained: single maintainer, idle since 2025-05-22, 71 open
  issues, PRs unreviewed for 15+ months. Contributing there is not a viable path.
- The most-wanted missing capabilities, per upstream issue traffic: charts (#123, #351 — "not
  supported", series references silently go stale after expansion), pivot tables (#200 corrupt
  output since 2021; #399 static-pivot regression in 0.2.12), images/shapes not moving with
  expansion (#354, #281, #249), conditional-formatting rules duplicated per generated cell
  (#216 — "3 conditions become 9 per cell… kills generation time"), and conditional row logic
  (`@if`, which MiniExcel has). XLibur — unlike ClosedXML — has a real chart API (spec 10) and
  a pivot cache model with `Refresh()`, so it is uniquely positioned to close the first two.
- Scriban (BSD-2-Clause, zero deps on net8+, very active — 7.2.6 July 2026) evaluates a bare
  expression to a **typed .NET object** (`ScriptMode.ScriptOnly` + `Template.Evaluate`),
  supports parse-once/evaluate-many, first-class custom-function registration, and real
  sandbox limits (loop/recursion/output caps, member filters, no reflection escape). Use
  ≥ 7.x only — the 2026 DoS advisories (GHSA-wgh7-7m3c-fx25 et al.) were fixed in 6.6.0.

## Decisions

Settled with the project owner, July 2026:

1. **Pluggable expression engine, Scriban as the default and flagship syntax.** A
   ClosedXML.Report-syntax compatibility engine (System.Linq.Dynamic.Core, C# expressions)
   **is in scope** (decision revised 2026-07-29), shipped as a **separate opt-in package**
   `XLibur.Report.DynamicLinq` that plugs into the `IExpressionEngine` seam — installing or
   removing it never touches the core `XLibur.Report` package or its dependency graph.
2. **No legacy-template compatibility contract.** There is no existing ClosedXML.Report
   template corpus to run unmodified. The template model (defined names, service row,
   `<<tags>>`, `{{ }}`) is kept because it is good, not because it is frozen — breaking
   cleanups are allowed where they fix known upstream design flaws.
3. **Charts, pivots and images are first-class in v1**: the engine re-points chart series
   references, moves picture anchors, re-points static pivot sources at the grown range and
   marks caches refresh-on-open. The `<<pivot>>` generation tag ports too, on XLibur's pivot
   API.
4. **Gap-fills in scope:** conditional-formatting range extension (upstream #216) and
   conditional row/range tags. **Out of scope** (deferred, see Non-goals): worksheet-per-item
   (#93), horizontal subranges (#225).
5. **Strict project isolation** (decision 2026-07-29): reporting code lives only in
   `XLibur.Report*` projects — its own Tests, Examples and Benchmarks projects. The sole
   core-side change in the whole spec is the `InternalsVisibleTo` grant.

## Scope

In scope:

- New projects `XLibur.Report` (TFMs `net8.0;net9.0;net10.0`, nullable enabled, warnings as
  errors — matching `XLibur.csproj`), `XLibur.Report.DynamicLinq`, `XLibur.Report.Tests` and
  `XLibur.Report.Examples`, mirroring the satellite-package layout of `XLibur.Fonts.SixLabors`
  (+`.Tests`/`.Examples`).
  Reporting stays fully isolated from the core projects: `XLibur.Tests`, `XLibur.Examples`
  and `XLibur.Benchmarks` are not touched — the sole core-side change in the whole spec is
  the one-line IVT grant in `XLibur/Properties/AssemblyInfo.cs`. The `ReportGenerate`
  benchmark gets its own `XLibur.Report.Benchmarks` project for the same reason.
  `XLibur.Report` is packaged and versioned by MinVer with the rest of the repo, published
  from the existing release pipeline.
- Template language: `{{ scriban-expression }}` in cell values, comments, hyperlinks and rich
  text; `&=` prefix for generation-time formula construction; Excel defined names binding
  `IEnumerable` variables to repeating ranges (vertical and horizontal); nested named ranges
  for vertical master-detail; the `<<Tag param=value>>` marker system with public custom-tag
  registration.
- Built-in tags ported: range/layout (`Range`, `SummaryAbove`, `DisableGrandTotal`,
  `OnlyValues`, `Delete`, `AutoFilter`, `ColsFit`, `RowsFit`, `Hidden`, `PageOptions`,
  `Protected`, `Height`), sorting (`Sort`/`Asc`/`Desc`), summary functions (`SUM`, `AVG`,
  `COUNT`, `MAX`, `MIN`, `PRODUCT`, `STDEV`, `VAR`, … with `over=`), grouping with subtotals,
  outline levels, `MergeLabels`, `PageBreaks`; image insertion (`Image`); pivot generation
  (`Pivot`, `Row`/`Column`/`Page`, `Data`).
- New tags: conditional inclusion (`If`) at row and range level.
- The Excel-function bridge: XLibur's `FunctionRegistry` exposed inside `{{ }}`.
- Chart/pivot/picture reference rewriting after range expansion.
- Conditional formatting: ranges extended, not duplicated.
- The ClosedXML.Report-syntax compatibility engine as a separate plug-in package
  `XLibur.Report.DynamicLinq` (see Design), verified by the upstream gauge corpus ported
  wholesale.
- Full test coverage: TUnit golden-file infrastructure plus unit tests in
  `XLibur.Report.Tests`, coverage target in the acceptance criteria; a `ReportGenerate`
  benchmark in `XLibur.Report.Benchmarks`.

Non-goals (recorded so future specs can pick them up):

- Worksheet-per-item generation (upstream #93) and subranges in horizontal tables (#225).
- Streaming/bounded-memory report generation (compose with spec 01's `XLStreamingWorkbook`
  later if demand appears) and async generation.
- Bug-for-bug equivalence with upstream ClosedXML.Report in the compatibility engine — the
  contract is "passes the ported upstream gauge corpus", not undocumented quirk parity.
- A template *designer* or any tooling beyond the library and docs.

## Design

### Project and public API

`XLibur.Report/` mirrors the upstream layout where it survives: root public types, `Tags/`
(upstream `Options/`), `Excel/` (range/subtotal helpers and `IXL*` extensions), `Expressions/`
(new — the engine seam and Scriban implementation), `Rewriting/` (new — post-expansion
reference rewriting). Public surface:

```csharp
namespace XLibur.Report;

public interface IXLTemplate : IDisposable
{
    IXLWorkbook Workbook { get; }
    void AddVariable(string alias, object value);
    void AddVariable(object value);                  // reflects public members into variables
    XLGenerateResult Generate();
    void SaveAs(string file); / SaveAs(Stream);
}

public class XLTemplate : IXLTemplate                 // ctors: path, Stream, IXLWorkbook
public class XLGenerateResult { bool HasErrors; TemplateErrors ParsingErrors; }

public interface IExpressionEngine                    // the pluggable seam
{
    object? Evaluate(string expression, ExpressionScope scope);   // typed result
    string  Interpolate(string text, ExpressionScope scope);      // mixed text + {{ }}
    CultureInfo Culture { get; }                                  // the report's culture (#275)
    bool    SupportsFunctions { get; }                            // optional capability
    void    AddFunction(string name, Delegate function);          // throws if unsupported
}

public static class TagsRegister { void Add<T>(string name, byte priority); }
public abstract class OptionTag { … Execute(ProcessingContext ctx); }
```

Errors accumulate into `TemplateErrors` and are written into the offending cells (red),
matching upstream behaviour — one bad expression must not abort the whole report (upstream
#340 is the counter-example to avoid).

### Expression engine: Scriban

`ScribanExpressionEngine` (the default, constructed by `XLTemplate` unless one is injected):

- **Whole-cell expression** `{{ expr }}` → parse with `LexerOptions { Mode =
  ScriptMode.ScriptOnly }`, evaluate via `Template.Evaluate` → typed object → `XLCellValue`
  conversion (decimal/double/DateTime/bool/string/Blank). **Mixed text** → normal
  `Template.Parse` + `Render`. Parsed templates are cached per distinct expression string
  (upstream caches compiled lambdas the same way); evaluation per data row pushes a per-row
  `ScriptObject` (`item`, `index`, `items`, plus globals) onto the context stack.
- **Identity `MemberRenamer`** (`member => member.Name`) so `{{ item.Price }}` binds to C#
  property names verbatim (Scriban's default would rename to `item.price`).
- **Relaxed access on**: missing member / null target yields null rather than throwing —
  report templates over sparse data want this.
- **Sandbox defaults on**: Scriban's loop/recursion/output limits stay at their safe defaults;
  the `TemplateContext` exposes only what `AddVariable` put in. One `TemplateContext` per
  `Generate()` call (it is not thread-safe; a template instance is single-generation at a
  time, same as upstream).
- `DataTable` variables convert to `Rows.Cast<DataRow>()`; `IDictionary` variables explode
  into individual variables — both upstream behaviours kept.

### Compatibility engine: `XLibur.Report.DynamicLinq` (separate package)

A second `IExpressionEngine` implementation that runs **ClosedXML.Report's C# expression
syntax** (`{{ item.Name.ToUpper() }}`, lambdas, LINQ over exposed collections) so
upstream-authored templates run unmodified — template *structure* (defined names, tags,
service row, `&=`) is engine-independent, so the expression language is the only delta.

- **Own NuGet package**, own project, own dependency (`System.Linq.Dynamic.Core` ≥ 1.6.x;
  never below 1.3.0 — CVE-2023-32571 was arbitrary method invocation). `XLibur.Report`
  itself never references it: engine selection is per-template
  (`new XLTemplate(path, new DynamicLinqExpressionEngine())`), so the package plugs in or
  out of a consuming app without touching the core report package.
- **Port of upstream `FormulaEvaluator` semantics** (~300–500 LOC): regex `{{ }}` splitting,
  `ParseLambda` + per-expression lambda cache, `item`/`index`/`items` row binding, `@`-prefixed
  globals inside range scope, non-generic `IEnumerable` re-cast via compiled `Cast<T>`,
  mixed-text interpolation with `InvariantCulture` and `DateTime → ToOADate()`.
- **`SupportsFunctions => false`** — the Excel-function bridge is a Scriban-engine feature;
  upstream syntax never had it, and Dynamic LINQ's static-method extension mechanism is not
  worth bridging. `XLTemplate` skips bridge registration for engines that decline.
- **Trusted templates only** — Dynamic LINQ has no sandbox; the docs state this and point
  untrusted-template scenarios at the Scriban default.
- **Conformance contract:** the upstream MIT gauge corpus (~111 template/gauge pairs) ports
  wholesale under this engine and *is* its test suite; the shared structural golden-file
  fixtures run parameterized across both engines.

### Excel-function bridge

`XLibur.Excel.CalcEngine.FunctionRegistry` (internal) already maps `SUM`, `AVERAGE`, `ROUND`,
`TEXT`, `EOMONTH`, … (~420 functions, spec 07) to implementations with Excel coercion
semantics. The bridge:

- Add `[assembly: InternalsVisibleTo("XLibur.Report")]` to `XLibur/Properties/AssemblyInfo.cs`
  (assemblies are not strong-named; the grant is a plain name, same as the existing
  Tests/Benchmarks grants).

> **Implementation decision (Task 5): two internal members added to the core.** The registry is
> built by a `private static` method into a private field, so the IVT grant alone does not reach
> it. Task 5 adds `FunctionRegistry.Names` (the registered names) and `XLCalcEngine.Functions`
> (the instance) — both `internal`, both read-only, no behaviour change; the full 11,603-test core
> suite is unaffected. This is a deviation from decision 5's "the sole core-side change is the IVT
> grant", taken because the alternative — hard-coding a function list in the report package —
> would break the property that makes the bridge worth having: a function added to the calc engine
> appears in templates with no further work.
- An adapter registers every registry function into the Scriban context under its **uppercase
  Excel name** (`SUM`, `IF`, `AVERAGE` — Scriban keywords are lowercase, so `IF` parses as an
  ordinary identifier), marshalling .NET arguments → `ScalarValue`/array values → function →
  `XLCellValue` → .NET. Functions that require a grid context (references, `OFFSET`-style)
  are **excluded** from the bridge — at template-evaluation time there is no grid; real cell
  formulas remain the tool for that, and the engine expands those.
- The bridge is one adapter over the registry, not per-function code: functions added to the
  calc engine later appear in templates automatically.

### Range expansion core

Direct port of the upstream pipeline, modernized (nullable, `XLCellValue`-native, TUnit):

- `RangeInterpreter` binds defined names to `IEnumerable` variables (including the
  `Parent_Child` underscore convention for nested sources), evaluates non-bound cells, then
  renders each bound range.
- `RangeTemplate` parses a bound range into a cell grid; the **last row is the options/service
  row** (tags + summary cells); inner defined names become recursive subrange templates
  (vertical master-detail).
- ~~`TempSheetBuffer` renders onto a `VeryHidden` buffer sheet, then `CopyTo` splices the block
  into the target sheet~~ — **superseded during implementation; see below.** The defined name is
  re-pointed via `SetRefersTo`.
- One deliberate behavioural change, replacing upstream's per-cell copying of conditional
  formats (#216): see "Conditional formatting" below.

> **Implementation decision (Task 3): insert-and-copy, not a temp-sheet buffer.**
> The buffer sheet is upstream's workaround for ClosedXML's slow and lossy row inserts. XLibur's
> structural-edit path was rewritten under spec 05 and is neither. Eight characterization tests
> (`RangeMechanicsCharacterizationTests`) pin what expansion relies on: `CopyTo` adjusts relative
> formulas and carries styles, merges and row heights; inserting rows shifts content, defined
> names and conditional-format ranges below and around the insertion point; deleting rows shrinks
> a name that spans them. Expanding by inserting sheet rows and copying the template block into
> them therefore delegates formula adjustment, merge tracking, CF extension, row sizing and
> name shifting to the core library instead of reimplementing all of it — and keeps the report
> package free of a second, divergent implementation of shifting. The buffer approach is not
> revisited unless the Task 9 benchmark shows insert-and-copy does not scale.

### Post-expansion reference rewriting (the differentiator)

The interpreter keeps an **expansion ledger** per worksheet: for every rendered bound range,
`(sheet, template area, rendered area, row/column delta below/right of it)`. After all ranges
on a sheet have rendered, `ReferenceRewriter` walks the workbook:

- **Charts** (`IXLChart` / `IXLChartSeries`): `ValueReferences` and `CategoryReferences` are
  parsed, re-pointed and written back. A reference that lies inside a template area is stretched
  to the rendered area; a reference entirely below one is shifted by the delta. XLibur does
  **not** shift chart references on row insert (spec 10's patcher never touches loaded charts
  unless edited), which is the gap the rewriter exists to close.

> **Implementation decision (Task 6): the references had to be made writable first, in the core.**
> The spec assumed the rewriter "setting `ValueReferences` marks the chart edited and the existing
> patch-on-save path persists it". A characterization test showed it does not: the setters were
> plain auto-properties raising no flag, `ChartPatcher.HasPendingChanges` never looked at them, and
> `PatchSeriesFormat` never wrote `c:cat`/`c:val`. Since every chart in a report template is a
> *loaded* chart, re-pointing one was a silent no-op — acceptance criterion 3 was unreachable. The
> core now tracks reference assignment through the same `AssignedFormat` mechanism as the
> formatting properties (`XLChartSeriesFormat.ValueReferences`/`.CategoryReferences`), seeds rather
> than assigns them when a series is created or loaded, and patches `c:f` while dropping the stale
> `c:numCache`/`c:strCache` so Excel redraws from the new range. This is the second deviation from
> decision 5's "the sole core-side change is the IVT grant", and a larger one than Task 5's — it
> changes save-path behaviour. Taken because it is also a plain core defect: `IXLChartSeries` has
> had a public settable `ValueReferences` that silently did nothing for every chart read from a
> file. The full core suite (11,603 tests) is unaffected.

- **Pictures**: nothing to do, established by measurement rather than assumed. A picture anchor is
  held as a live 1×1 range enrolled in the worksheet's range repository, so a full-row insert
  shifts it like everything else, and the shift survives the save. The rewriter therefore contains
  no picture code at all; `PicturePlacementTests` pins the inherited behaviour end to end, because
  behaviour nobody wrote is the kind that disappears quietly.
- **Pivot caches**: a cache whose source is an **area reference** intersecting a template area
  gets its source re-pointed at the rendered area (internal `XLPivotCache.Source` via the IVT
  grant; promote to a public `IXLPivotCache` setter only if a public need emerges), then
  `Refresh()` and `SetRefreshDataOnOpen()`. Caches sourced from **tables or defined names**
  need no re-point (the name/table already grew) — refresh only. This restores and hardens the
  "static pivot over a named range" pattern that upstream 0.2.12 regressed (#399), and it is
  the *documented* happy path; the `<<Pivot>>` generation tag is ported for dynamic layouts
  but the static pattern is primary.

> **Implementation finding (Task 7): a pivot table does not move, either.** Characterization showed
> the expected half — an area source is a plain sheet-plus-rectangle value and does not follow row
> inserts, while a name or table source does but still needs the refresh — and one the spec did not
> anticipate: `IXLPivotTable`'s position is a plain rectangle too, so a pivot sitting below a bound
> range stays where the template put it while the generated rows multiply underneath it and are
> written over it. `PivotRewriter` moves it. No core change was needed for any of this;
> `TargetCell` is already settable and `Source` is reachable through the IVT grant. The grant to
> `XLibur.Report.Tests` was added so the tests can assert on a cache's source and record count,
> neither of which is on the public surface.

#### The `<<Pivot>>` generation tag

As implemented. `<<Pivot dest="Summary!A3">>` in the options row builds a pivot over the generated
block; `<<Row>>`, `<<Column>>` (or `<<Col>>`), `<<Page>>` and `<<Data>>` in the options row say what
each column is for.

- **A field is named by its heading**, the cell above the range — the name the cache reads out of the
  source anyway, so a template does not write it twice. A range starting in row 1 has no heading row
  and is reported rather than given invented names.
- **`<<Data>>` sums** unless told otherwise, as a flag (`<<Data avg>>`) or a value
  (`<<Data func=avg>>`); `title=` renames the value column.
- **`dest` is read in template coordinates**, like everything else a template writes: a cell marked
  below the range is below the *generated* range in the report. It takes a sheet-qualified reference
  or the name of a single-cell defined name.
- **Refused, not half-done**: a horizontal range (a pivot names its fields from a heading row, which
  such a range has not got) and a grouped range (its subtotal rows are inside the generated block, and
  a pivot would count them as data on top of the rows they already total).
- **Runs last of the tags** (priority 255, after `<<Delete>>`), and the pivot is *built* later still —
  after the rewriters, at the very end of a run.

> **Implementation finding (Task 7b): the tag cannot build the pivot when it runs.** Everything about
> a pivot that Task 7 had already learned applies to making one: its position and its cache source are
> plain rectangles, so a pivot created while ranges are still being generated stays behind as the sheet
> moves under it — the options row being dropped out from under it is enough. Two mechanisms fell out
> of that, and both are the same mechanism. `<<Pivot>>` resolves its destination and each field's
> heading cell to **live one-cell ranges before expansion**, and `PivotBuilder` reads them once
> everything has finished moving. That is what makes `dest` mean what the template says, what lets
> `<<Delete>>` take a column and its field away together, and what removes all the arithmetic that
> would otherwise have been needed — the live-range rule from finding 6, used deliberately rather than
> merely observed. `RangeShiftMechanicsCharacterizationTests` pins the seven behaviours it relies on.
> Building *after* the rewriters also means a generated pivot shares a cache with a template pivot
> re-pointed to the same rows, rather than the workbook carrying two identical caches.

### Grouping and subtotals

`<<Group>>` sits in the options row under the column to group by, and takes that column's template
expression as its key the same way `<<Sort>>` does. Several nest, leftmost outermost. As
implemented:

- **The engine orders the rows.** All the levels are ordered in one `OrderBy`/`ThenBy` chain, so
  the leftmost is the primary key; the ordering is stable, so a `<<Sort>>` (which runs first, at
  priority 10) still decides the order within a group. `nosort` opts out per level.
- **Ordering and labelling follow `IExpressionEngine.Culture`**, not `CultureInfo.CurrentCulture`
  ([#275](https://github.com/XLibur/XLibur/issues/275)). `SortKeyComparer` collates text through that
  culture's `CompareInfo` and `KeyText` formats a key through it, so a default template is invariant
  end to end and a template handed `new ScribanExpressionEngine(new CultureInfo("sv-SE"))` sorts and
  labels the way its readers read. Both the comparer's string arm and its mixed-type text fallback
  were affected; `string.CompareTo` collates under the current culture, so strings are taken before
  the general `IComparable` case.
- **Each group gets a subtotal row** carrying whatever summary tags the options row declares, over
  that group's rows alone, plus a `{0} Total` label in the grouped column. Below the group by
  default, `summaryAbove` above it. Where several stack at one boundary they read outwards from the
  rows they cover: innermost first below a group, outermost first above one.
- **The subtotal row takes the options row's cell styling** — the only styling a template can
  express for a row that does not exist until generation, and what makes a group total look like
  the grand total.
- **The block is outlined**: data rows at the innermost level, each subtotal row one level out from
  the rows it covers, so collapsing a level in Excel leaves its totals showing. Excel's eight-level
  limit clamps rather than fails.
- **Parameters**: `by`, `desc`, `nosort`, `totalLabel`, `merge`/`mergeLabels`, `summaryAbove`,
  `pageBreaks`, `collapse`, `disableSubtotals`. Each but `by`, `desc`, `nosort` and `totalLabel`
  also exists as a range-wide options-row tag (`<<MergeLabels>>`, `<<SummaryAbove>>`,
  `<<PageBreaks>>`, `<<Collapse>>`, `<<DisableSubtotals>>`), for a template with several levels.
  `<<DisableGrandTotal>>` has no per-level form: it leaves the options row's own summaries
  unwritten, which is how a report shows a total per group and none for the report.

Rendering inserts sheet rows for the subtotals, bottom-up so that a row number worked out from the
layout is still valid when its turn comes — the same reasoning as insert-and-copy expansion, and for
the same reason: the core library shifts content, formulas, names and conditional formats correctly
already.

### Conditional formatting

During expansion the buffer does not copy CF rules per generated cell. Instead, rules whose
range intersects the template row are recorded once, and after splicing the rule's applied
range is **extended to the rendered block** (relative references stay R1C1-correct via the
existing round-trip). Rule count in the output equals rule count in the template — upstream
produces `rows × rules` duplicates (#216).

### Conditional tags

`<<If test="expr">>`:

- **Row level** (tag in one of the range's repeated rows): rows where `test` is falsy are left out.
  Scriban truthiness applies (`null`/`false` falsy) — the docs must say so explicitly, since `0` is
  truthy.
- **Range level** (tag in the options row): a falsy `test` (evaluated against the range's
  `items`) renders the range with zero rows — headers and options-row summaries behave exactly
  as an empty collection does.

> **Implementation decision (Task 8a): the test runs as an item transform, before everything.**
> Filtering is a transform on the item list, which is where sorting and grouping already happen, so
> `IfTag` is an ordinary `OptionTag.TransformItems` at priority 1 — ahead of `<<Sort>>` (10) and
> `<<Group>>` (20). What survives the test is therefore what everything downstream sees: the
> survivors are what gets sorted, grouped and totalled, and nothing else has to know a row was
> dropped. Filtering *after* grouping would leave group keys pointing at rows that no longer exist.
>
> Two supporting changes fell out of it. `OptionTag` gained **`InRepeatedRow`**, because until now
> every tag lived in the options row and placement carried no meaning; the expander now reads tags
> from the repeated rows too, so a tag that means something at both scales can tell which was meant.
> And `ProcessingContext.IsTrue` evaluates a tag parameter as an expression — a bare one, an
> interpolated one or a literal — which is what `test=` needs and, incidentally, what
> `<<Delete keep=…>>` had documented since Task 4 without it ever working: tag parameters were
> parsed from raw cell text before evaluation, so `keep="{{ ShowWorkings }}"` compared the literal
> string `{{ ShowWorkings }}` against `"true"` and always deleted the column. Both tags now go
> through one evaluator and one truthiness rule (`ExpressionTruth`), which is also what keeps the
> two engines answering the same question the same way.

### Testing

- **A dedicated `XLibur.Report.Tests` project** (TUnit, mirroring `XLibur.Tests`
  infrastructure conventions — serial execution, en-US culture defaults — without referencing
  it): `Resource/Templates/*.xlsx` + `Resource/Gauges/*.xlsx`, and a semantic
  workbook comparer asserting cell values/types/formulas, comments, hyperlinks, rich text,
  merged ranges, styles, conditional formats (count and ranges — the #216 fix makes this
  meaningful), row/column sizes and outline levels, page setup — with the actual output saved
  to a diagnostics folder on mismatch. This is the upstream `CompareWithGauge` approach
  rebuilt on TUnit/awaitable assertions.
- Upstream's MIT-licensed template/gauge pairs port **wholesale as the compatibility
  engine's conformance suite** (their syntax matches it exactly); for the Scriban engine
  they port selectively (simple `{{ item.X }}` templates unchanged; ones using C# method
  calls re-authored). Structural golden-file fixtures are parameterized to run under both
  engines. `XLibur.Report.Tests` hosts all of it — it references the DynamicLinq package;
  the shipped `XLibur.Report` package never does.
- Chart/pivot/image assertions reload the generated workbook through XLibur and assert
  series references / cache sources / anchors; per repo ground rules, file-format-affecting
  tasks (5–7) also record a manual "opens clean in Excel" check.
- Every built-in tag gets at least one golden-file test; the expression engine, function
  bridge, ledger/rewriter get direct unit tests.

### Packaging, CI, docs

MinVer versioning is inherited from `Directory.Build.props` automatically. Work items: add the
five projects to `XLibur.slnx` under an `/XLibur.Report/` solution folder (pattern: the
`XLibur.Fonts.SixLabors` folder), `GeneratePackageOnBuild` in Release for `XLibur.Report`
and `XLibur.Report.DynamicLinq` (pattern: `XLibur.Bundle`), NuGet readmes, verify
`release.yml` picks up the new `.nupkg`s and CI runs the new test project, a docs-website
section (template-language reference, tag reference, function-bridge list, chart/pivot
patterns, Scriban↔C#-syntax migration page), and a `XLibur.Report.Examples` project.

## Work plan

PR-sized tasks; each lands green (build + tests) on its own branch per repo ground rules.

**Status (July 2026, branch `feat/spec-12-report-templating`):** All eleven tasks are done, with one
carve-out: Task 10's upstream gauge corpus is not obtainable from this machine and was not ported — see
**Still open** under Results, which also lists the gaps left inside the finished tasks. 447 report tests
green on net8.0 and net10.0, the solution builds clean in Release, and the core suite (11,603 tests)
passes with the chart-reference fix Task 6 needed.

1. **Scaffold + expression engine.** `XLibur.Report` + `XLibur.Report.Tests` projects, slnx
   entries, CI test wiring, IVT grants (`XLibur` → `XLibur.Report`; `XLibur.Report` →
   `XLibur.Report.Tests`), `IExpressionEngine` + `ScribanExpressionEngine` (typed eval,
   interpolation, caching, relaxed access, identity renamer, sandbox defaults),
   `TemplateErrors` model. Unit tests for every evaluation shape (typed results incl.
   decimal/DateTime/Blank, nulls, mixed text, error capture).
2. **Golden-file test infrastructure** in `XLibur.Report.Tests`. Semantic comparer, resource
   layout, TUnit fixture helpers, diagnostics-on-mismatch. Tested against hand-made
   equal/unequal workbook pairs.
3. **Vertical range expansion.** `RangeInterpreter`, `RangeTemplate`, `TempSheetBuffer`,
   defined-name binding (incl. underscore paths), service row, R1C1 formula handling, merged
   ranges, styles/heights, comments/hyperlinks/rich text, nested vertical subranges, `&=`
   formulas. First golden-file suite.
4. **Tag framework + core tags.** `OptionTag`/`TagsRegister`/parser, sorting, summary
   functions with `over=`, layout tags, `Image`, grouping + subtotal engine (outline levels,
   `MergeLabels`, `PageBreaks`, `SummaryAbove`, grand totals). Largest port task; may split
   grouping into its own PR at implementation time.
5. **Excel-function bridge.** Registry adapter, uppercase registration, grid-context function
   exclusion list, value marshalling, docs table of exposed functions.
6. **Expansion ledger + chart/picture rewriting.** **Done** — ledger, `ReferenceRewriter` for
   chart series, picture behaviour characterized as needing nothing, tests asserting re-pointed
   references, manual Excel check done 2026-07-30.
7. **Pivot support.** **Done** — static-pivot re-point + refresh (area/table/name sources),
   `<<Pivot>>` with the `<<Row>>`/`<<Column>>`/`<<Page>>`/`<<Data>>` field tags on XLibur's pivot
   API, and the manual Excel check for both paths (upstream #200's corrupt-output history makes the
   validator *and* an Excel open non-negotiable here, 2026-07-30).
8. **Horizontal tables + conditional tags + CF extension.** **Done** — horizontal rendering
   parity (no subranges, no grouping), `<<If>>` at row and range level, and CF
   extend-not-duplicate with rule-count assertions in both directions (the CF half landed in
   Task 3, see finding 2).
9. **Packaging, docs, benchmark.** **Done** — `XLibur.Report` added to `release.yml`'s pack list and
   its package verified (three TFMs, XML docs, readme/icon/NOTICE, and exactly two dependencies:
   `XLibur` and `Scriban`); `docs/report-templating.md` written as the template-language and tag
   reference and linked from the README; `XLibur.Report.Benchmarks` added with the `ReportGenerate`
   benchmark plus a `scaling` probe and a `phases` probe. Numbers in Results.

   **Correction:** this task was first recorded as done with the note that there is no docs site in this
   repo. That was wrong and unverified — `docs-website/` is a Docusaurus site right here, built to
   `xlibur.github.io/XLibur` by `.github/workflows/docs.yml` from `main`. The reference was written as
   `docs/report-templating.md` and linked from the README, which is useful but is not the docs-website
   section this task asks for. **Outstanding:** add `docs-website/docs/report-templating.md` (adapting the
   in-repo reference to the site's conventions — see the sibling pages, e.g. `pivot-tables.md`) and add it
   to `sidebars.ts`.
10. **Compatibility engine (`XLibur.Report.DynamicLinq`).** **Done except the corpus.** The package,
    the engine (`DynamicLinqExpressionEngine` on `IExpressionEngine`), 37 tests including
    cross-engine agreement, packaging, and the docs (trusted-templates-only caveat, engine
    selection) are all in. **The upstream gauge corpus could not be ported**: it lives in
    ClosedXML.Report's GitHub test project, not in its NuGet package, so it is not obtainable from
    this machine — the cached `closedxml.report` packages hold only compiled assemblies. See the
    note under Results.
11. **`XLibur.Report.Examples`.** **Done** — ten examples, a menu in `Program`, a `README.md`, and
    five smoke tests in `XLibur.Report.Tests` that run every one of them. A runnable console project
    of worked examples, each writing
    both the template it authored and the report it generated so a reader can open the pair and
    see what the template language did. Mirrors the `XLibur.Examples` layout (one class per
    example, a menu in `Program`), references `XLibur.Report`, and is **not** packaged.

    The flagship is an **annual sales report**: a title band bound to workbook variables
    (company, year, run date), a heading row, a range bound to a collection of line items
    repeating one row per sale, per-row formulas (`&=`), grouping by region with subtotals and
    an outline, a sort inside each group, `<<AutoFilter>>` + `<<ColsFit>>`, number and date
    formats carried from the template, conditional formatting that colours the margin column
    red below target and green above it — a template rule count of one, whatever the row count,
    which is the point — and a grand-total row. Smaller examples cover one thing each: the
    minimum viable template; the Excel-function bridge (`{{ SUM(items.Total) }}`,
    `{{ IF(...) }}`, `{{ ROUND(...) }}`); custom-tag registration through `TagsRegister`; error
    handling, showing a deliberately bad expression producing a cell-level error and
    `HasErrors` rather than an exception; and — once Tasks 6 and 7 land — a chart and a pivot
    over a bound range, growing with the data.

    Each example is exercised by a smoke test in `XLibur.Report.Tests` that runs it and asserts
    it generated without errors, so an example cannot rot into a snippet that no longer
    compiles or no longer works. Examples double as the docs' worked code: the docs-website
    section in Task 9 links them rather than repeating them.

    One of the examples should be the **Excel-verification workbook**: the coverage the 2026-07-30
    manual check used (grouping, merged labels, one CF rule over a generated block, a picture below
    the range, a re-pointed chart series, an area-sourced pivot re-pointed and a table-sourced one
    refreshed), writing template and report side by side. It was a throwaway generator that time;
    committing it makes the manual check repeatable rather than reconstructed from scratch every
    time the file format is touched.

    **As built**, two examples were added beyond the list above and one was changed by what it turned
    out to be true about. `HorizontalReport` covers Task 8's `<<Horizontal>>`, which post-dates this
    plan; `ConditionalRows` covers `<<If>>` at both scales. `EverythingAtOnce` is the verification
    workbook, and it grew a `<<Pivot>>`-generated sheet to check against the drawn one. The
    error-handling example gained a second half — see finding 9 — because three of the four mistakes
    it was written to show turned out not to be mistakes.

Sequencing: 1 → 2 → 3 → {4, 5, 6, 10 in parallel} → 7 → 8 → {9, 11}. Task 5 only needs Task 1;
Tasks 6–7 need Task 3's ledger; Task 8 needs Task 4's tag framework; Task 10 needs Tasks
2–3 (fixtures + expansion core) and is independent of everything after. Task 11 wants the
feature set settled, so it goes last with Task 9 and can share its PR; its chart and pivot
examples need Tasks 6–7. Conflict map: 4↔8 (tag framework), 6↔7 (rewriter), 10↔4 (both touch
shared test fixtures — coordinate or sequence). Everything in `XLibur.Report*/` is disjoint from
the core library except the one-line IVT grant (Task 1) — no conflicts with open specs 03/04/08.

## Acceptance criteria

1. A template with a bound vertical range, nested subrange, grouping with subtotals, sorting,
   summary row, merged cells, row-relative formulas, comments and hyperlinks generates
   correctly (golden-file verified) — parity with the upstream feature set minus non-goals.
2. `{{ SUM(items.Price) }}`, `{{ IF(item.Qty > 10, "bulk", "unit") }}`, `{{ ROUND(item.Price,
   2) }}` evaluate through the calc-engine bridge with Excel semantics and **typed** results
   (a decimal sum lands as a number cell, not text).
3. A template chart whose series reference a template range shows **all** generated rows after
   expansion (reloaded references assert the rendered area), and a picture below the range
   sits below the generated rows. Manual Excel check recorded.
4. A static pivot over a bound named range reflects the full generated data after
   refresh-on-open; `<<Pivot>>`-generated output opens in Excel with **no repair dialog**
   (regression tests for upstream #399 and #200 scenarios).
5. Conditional formatting: output rule count equals template rule count with ranges covering
   the rendered block (upstream #216 scenario as a regression test).
6. `<<If test=…>>` omits rows/ranges when falsy, golden-file verified.
7. Coverage: ≥ 90% line coverage on `XLibur.Report` (MTP `--coverage`); every built-in tag has
   at least one golden-file test; one bad expression yields a cell-level error plus
   `HasErrors`, never an exception aborting generation.
8. `ReportGenerate` benchmark (100K rows × 10 cols, one group level): completes without the
   temp-buffer approach degrading super-linearly; numbers recorded in this spec's Results as
   the baseline for future perf work.
9. Packages publish from the existing tag-driven release with MinVer-derived versions;
   Scriban ≥ 7.x is the only new runtime dependency of `XLibur.Report`, and
   System.Linq.Dynamic.Core is a dependency of `XLibur.Report.DynamicLinq` only.
10. An unmodified upstream-syntax template (C# expressions, `@` globals, `&=` formulas)
    generates correctly under `DynamicLinqExpressionEngine`; the ported upstream gauge
    corpus passes; adding/removing the DynamicLinq package requires no change to code using
    the default engine.
11. `XLibur.Report.Examples` runs end to end and writes, for every example, both the template
    and the generated report. The annual sales report shows repeated rows, a bound title,
    grouping with subtotals, per-row formulas and conditional colouring, and its generated
    workbook holds **one** conditional-formatting rule however many rows it produced. Every
    example has a smoke test asserting it generates without errors.

## Results

Twelve commits on `feat/spec-12-report-templating`, July 2026.

**What landed.** `XLibur.Report` (Scriban engine, template model, range expansion in both
directions, tag framework including grouping, subtotals and conditional inclusion, Excel-function
bridge, chart reference rewriting, static pivot re-point and refresh, and pivot generation from
`<<Pivot>>`), `XLibur.Report.DynamicLinq` (the ClosedXML.Report-syntax compatibility engine),
`XLibur.Report.Tests` (447 tests), `XLibur.Report.Examples` (ten worked examples, each writing its
template and its report) and `XLibur.Report.Benchmarks` (the `ReportGenerate` benchmark and two probes).
All five are in `XLibur.slnx`; CI builds them and runs the suite with coverage, and `release.yml` packs
and publishes both report packages with the rest.

**Dependencies, as criterion 9 requires them.** Verified by unpacking the `.nupkg`s: `XLibur.Report`
depends on `XLibur` and `Scriban` and nothing else; `System.Linq.Dynamic.Core` appears only in
`XLibur.Report.DynamicLinq`. Engine choice is per template and the core report package never references
the compatibility one, so installing or removing it changes nothing for code using the default.

**Ten findings that changed the design.**

1. **The temp-sheet buffer was not needed** (Task 3). It is upstream's workaround for ClosedXML's
   slow and lossy row inserts, which spec 05 rewrote. Eight characterization tests established
   that `CopyTo` adjusts relative formulas and carries styles, merges and heights, and that row
   inserts shift content, defined names and CF ranges. Expansion is therefore insert-and-copy, and
   the report package contains no second implementation of shifting. This also produced the #216
   fix nearly for free — with one caveat, below.
2. **Copying a block copies its conditional formats**, so insert-and-copy alone still produced
   `rows × rules`. The fix is explicit: capture the template's rules, drop the copies, and widen
   the originals through the internal `XLConditionalFormat.SetAreas`. `IXLConditionalFormat.Ranges`
   is a fresh projection of the rule's area list and mutating it does nothing — the source
   documents this. Output rule count now equals template rule count.
3. **The function registry is not reachable through the IVT grant alone** (Task 5) — it is built by
   a `private static` method into a private field. Two read-only `internal` members were added to
   the core (`FunctionRegistry.Names`, `XLCalcEngine.Functions`). Recorded above as a deviation
   from decision 5.
4. **Grouping cannot be one tag acting alone** (Task 4b). Each `<<Group>>` is a *declaration*; the
   ordering, the runs, the subtotal rows and the outline are worked out together in
   `GroupRenderer`, because nesting is a property of the levels as a set. Chaining one stable sort
   per level in tag order would have made the *last* level the primary key — the opposite of how
   the levels nest — so the levels are ordered in a single `OrderBy`/`ThenBy` chain instead. The
   same reasoning applies to the subtotal rows: which of several stack first at a group boundary is
   a question about all the levels at once, not about any one of them.
5. **Charts needed a core fix; pictures needed no code at all** (Task 6) — the reverse of what the
   spec expected on both counts. Setting a loaded chart's series reference was a silent no-op, so
   the core now tracks reference assignment and patches it (recorded above as a deviation from
   decision 5). Picture anchors, which the spec listed as an unverified risk, turn out to move on
   their own: an anchor is a live range enrolled in the worksheet's range repository, so a full-row
   insert shifts it and the shift survives the save. Both were settled by characterization tests
   before a line of the rewriter was written, which is the only reason the wrong one was not built.
6. **A pivot table does not move either** (Task 7). The spec's list of what to re-point covered
   cache sources and stopped there. But a pivot table's position is a plain rectangle, like a cache
   area and unlike everything the core shifts, so a pivot below a bound range sat still while the
   generated rows multiplied underneath it and wrote over it. `PivotRewriter` moves it. The pattern
   across Tasks 6 and 7 is worth naming for whoever does Task 8: **anything the core holds as a live
   range moves for free; anything it holds as a value does not, and belongs in the rewriter.**
   Charting the two apart with characterization tests before designing has now paid three times.
7. **The ledger under-reported by one whenever the options slot survived** (Task 8). Turning the
   expander's row arithmetic into axis arithmetic forced every number to be justified, and this one
   could not be: the delta was `rendered slots − template slots`, which silently ignores that an
   options row a total was written into *stays*, moving everything below the range one further. Every
   Task 6 chart test had used a template whose options row was empty, so all of them agreed with the
   wrong answer. The delta is now what the range actually ends up occupying, and
   `ASeriesBelowARangeWhoseOptionsRowSurvivesMovesOneFurther` is the test that would have caught it.
   The general lesson: a derived quantity that happens to be right for the cases you tested is not
   the same as a correct one, and rewriting code against an abstraction is a good way to find out
   which you have.
8. **A tag cannot build a pivot at the moment it runs** (Task 7b). The live-range rule from finding 6
   turned out to be the *design* and not just an observation: `<<Pivot>>` resolves its destination and
   its fields' heading cells to live one-cell ranges before the range expands, and the pivot is built
   at the end of the run, once everything has finished moving. Doing it in the tag needed arithmetic
   for the options row being dropped, could not survive `<<Delete>>` taking a column out from under a
   field, and read `dest` in whatever coordinates the sheet happened to be in mid-generation. Holding
   the two addresses live instead removed all three problems and the arithmetic with them. The moral
   for the rest of the package: when a value has to survive the sheet moving, ask the core to hold it
   rather than working out where it went.
9. **Three of the four "mistakes" in the error-handling example are not mistakes** (Task 11). Writing
   an example that shows what a template error looks like found that an unbound variable, a missing
   property and an unregistered tag *outside* a bound range are all silent: the first two because
   relaxed member access reads them as blank, which is exactly what makes a template survive an
   optional field, and the third because tags are only read inside a bound range. Only two of the four
   cases the example was drafted with reported anything. The example now has a second half about what
   is *not* an error, which is the more useful half — a typo in a name is silent, so an empty column is
   the first place to look for one. Worth generalising: writing the documentation is a test of the
   library, and the cheapest one available.
10. **Generation was super-linear, and the cause was one call per item** (Task 9). The first measurement
    failed criterion 8 outright: the cost per row roughly doubled between 10,000 rows and 50,000, so a
    50,000-row report took 20 seconds where linear scaling predicted 3. The `phases` probe localised it
    with no report code involved at all — `IXLRange.CopyTo` costs more the larger the target worksheet
    is, independently of how much is being copied, tripling per call over 30,000 rows, while the
    per-block `CellsUsed` that evaluation walks stayed flat at about a microsecond. Expansion was making
    one `CopyTo` per item, so the engine inherited the core's curve and squared it. The fix is four
    lines: fill the inserted slots by repeatedly **doubling** what has already been written rather than
    copying the template once per item. Every copy is of the *unevaluated* template, so after any number
    of rounds the filled region is n identical blocks, and doubling is therefore indistinguishable from
    copying n times — ⌈log₂ n⌉ calls instead of n, sixteen rather than fifty thousand. All 389 existing
    tests passed unchanged, which is what says the two are equivalent; `BlockCopyingTests` then pinned
    the boundaries where a doubling round is truncated. The result is flat per-row cost out to 100,000
    rows and **6.8× faster** at 50,000. Two lessons: the acceptance criterion earned its place, because
    nothing else in the suite would have caught this; and a per-item call into an API whose cost depends
    on the total size is worth suspecting before it is worth measuring.

**Benchmark baseline (criterion 8), 2026-07-30.** Ten columns, `XLibur.Report.Benchmarks`, net10.0,
Release, one run per size:

```
dotnet run -c Release --project XLibur.Report.Benchmarks -f net10.0 -- scaling
```

| Template shape | 25K | 50K | 100K | µs/row at 100K | per-row vs 25K |
|---|---|---|---|---|---|
| Plain — ten expressions, no tags | 1.3 s | 2.7 s | 5.4 s | 54 | 0.98× |
| Totalled — three `SUBTOTAL` columns | 1.3 s | 2.6 s | 5.2 s | 52 | 1.03× |
| **Grouped — one group level, five groups, two totals** | **2.1 s** | **4.1 s** | **8.3 s** | **83** | **0.98×** |
| GroupedAndSorted — plus a sort inside each group | 2.7 s | 5.6 s | 11.9 s | 119 | 1.09× |

The bolded row is the workload criterion 8 names: **100,000 rows × 10 columns with one group level, in
8.3 seconds, with per-row cost flat**. Before finding 10's fix the same shape ran at 2.4× the per-row
cost at 50,000 rows that it did at 10,000 — the criterion was not met and the numbers here are the
second measurement, not the first.

`ReportGenerateBenchmarks` is the BenchmarkDotNet suite over the same four shapes, with a `HandWritten`
baseline that writes the same grid by an ordinary loop, for when two implementations need comparing at
one size. A full run takes the better part of an hour, so `scaling` is the instrument to reach for; a
`phases` mode splits expansion into its copy and evaluate halves, which is what localised finding 10.

**Releasing.** The report packages are **not** in core's lockstep set. They version off their own tags:

```
git tag -a report-v0.1.0 -m "XLibur.Report 0.1.0"
git push upstream report-v0.1.0        # fires .github/workflows/release-report.yml
```

`XLibur.Report.props`, imported by both packable report projects, is where the stream is declared:
`MinVerTagPrefix` is `report-v`, so MinVer reads `report-v*` and ignores core's `v*` — a core release no
longer renumbers Report, and a Report release needs no core release. `release-report.yml` builds, runs the
suite, packs the two packages, checks their versions against the tag, and publishes; a
`workflow_dispatch` run is a dry run by default.

The same file declares `XLiburDependencyVersion`, the core version the packages depend on, because a
`ProjectReference` would otherwise write whatever core happened to be at pack time — usually a
pre-release, which would drag every consumer of a stable Report onto a pre-release core. It is an
**exact** range (`[0.107.0]`) rather than a floor for as long as Report reads core internals through the
IVT grant: internals carry no compatibility contract, so a floor would be a promise core never made, and
NuGet would unify core upward and break Report at runtime. Spec 13 replaces the grant with a public
surface; when its task 3 lands, this becomes an open floor and the exactness goes.

Until core 0.107.0 is on NuGet, **a Report release is refused rather than published broken**: the workflow
reads the dependency out of the packed nuspec and fails if that version is not on NuGet, since an exact
dependency on an unpublished version is unresolvable. Verified — the gate blocks today, core's latest
stable being 0.106.0. Core's own `release.yml` also had its previous-tag resolution narrowed to `v[0-9]*`
releases so a `report-v*` release can never become the tag core's notes diff against.

**Coverage (criterion 7), 2026-07-30.** Measured by running the suite with the MTP coverage extension:

```
XLibur.Report.Tests/bin/Release/net10.0/XLibur.Report.Tests.exe --coverage --coverage-output-format cobertura
```

| Assembly | Line coverage |
|---|---|
| **`XLibur.Report`** | **90.1%** |
| `XLibur.Report.DynamicLinq` | 95.3% |
| `XLibur.Report.Examples` | 70.4% (documentation; excluded from the gate) |

Criterion 7 asks for ≥ 90% on `XLibur.Report`, so it is met — but only measured at the point the PR's
quality gate failed, which is worth recording: **SonarCloud had never seen this package's coverage at
all.** The scanner's `sonar.cs.vscoveragexml.reportsPaths` was `**/coverage.xml` while the report
suite's CI step writes `report-coverage.xml`, so the gate read 0.6% on new code for a package at 90%.
The glob is now `**/*coverage.xml`, and `XLibur.Report.Examples`/`.Benchmarks` are added to
`sonar.coverage.exclusions` alongside their core counterparts. A CI step that produces coverage nobody
reads is worse than no step, because it reports success.

**Confirmed, not assumed.** Every Scriban property the spec relied on holds: `ScriptMode.ScriptOnly`
returns typed objects (a `decimal` stays a `decimal`), an identity `MemberRenamer` keeps C# member
names, relaxed access turns sparse data into blanks, and an uppercase `IF` parses as a function
call despite `if` being a keyword. Scriban also honours a `params` delegate, which is what lets the
bridge register variadic Excel functions through the one-method engine seam.

**Still open.** All eleven tasks have landed; what remains are gaps inside them, listed here so none of
it has to be rediscovered.

- **The upstream gauge corpus was not ported** (Task 10's conformance contract). It is in
  ClosedXML.Report's GitHub test project, and the NuGet package — the only copy reachable from this
  machine — ships compiled assemblies and no test resources. What stands in for it: 24 engine unit tests
  over the syntax surface, and 13 tests that build the *same* report twice, once per syntax, and assert
  the two generated workbooks agree cell for cell. That is a real conformance statement about the engine
  seam, and it is not the same statement as "111 upstream template/gauge pairs pass". **To finish:**
  clone `ClosedXML/ClosedXML.Report`, copy `tests/ClosedXML.Report.Tests/Resources` into
  `XLibur.Report.Tests`, and run each template/gauge pair through `WorkbookComparer` under
  `DynamicLinqExpressionEngine`. Expect real failures — upstream quirk parity is explicitly *not* the
  contract (see Non-goals), so each will need a judgement about whether it is a bug or a quirk.
- **Expressions are evaluated in cell values only.** Comments, hyperlinks and rich text are in the Scope
  list and are not implemented; writing the docs turned that up.
- **Nested vertical subranges.** `RangeBinder` resolves property paths from workbook variables, but a
  child range inside a parent's rows is not yet expanded per parent item.
- **Six tags from the Scope list** have no implementation: `Image`, `PageOptions`, `Protected`,
  `Height`, `OnlyValues` and the `Range` marker.
- **The docs-website page** (Task 9). `docs/report-templating.md` is the reference, but the published site
  is `docs-website/` in this repo and has no report page. Adding one is a page plus a `sidebars.ts` entry.
- **Report releases on its own version stream, and cannot release yet.** Task 9 originally added both
  report packages to `release.yml`'s core pack loop. That was dropped when this branch merged `main`:
  #269 scoped the release tag glob to the core version stream (`v[0-9]*`) as groundwork for versioning an
  addon independently, and spec 13 states the intent plainly — Report "moves on its own cadence, is
  pre-1.0, and should neither inherit core's version nor drag core along". The independent stream is now
  built rather than deferred; see **Releasing** below. What it waits on is **core 0.107.0 being published**
  with spec 13's public surface, which the pipeline enforces rather than trusts.
- **The two `internal` core members this branch added are superseded** (finding 3). Spec 13 replaces the
  `InternalsVisibleTo` grant with a public surface — `XLFunctionLibrary.TryInvoke`/`Names`, and
  `IXLPivotCache.SourceRange`/`SourceWorksheet`/`SetSourceRange` — and keeps `FunctionRegistry.Names` and
  `XLCalcEngine.Functions` internal, serving that class instead of serving Report. Moving
  `ExcelFunctionBridge` and `PivotRewriter` onto it is spec 13's task 3, and is a prerequisite for Report
  depending on a *published* core package rather than the project. Until then Report builds only in-repo.
- **The `CopyTo` fix belongs in the core.** Finding 10's workaround makes the report engine linear; the
  underlying defect is spec 14 / [#271](https://github.com/XLibur/XLibur/issues/271), with the two-line
  fix prototyped and measured there.
- Grouping is deliberately vertical-only, and `<<Pivot>>` deliberately neither horizontal nor grouped;
  each says so rather than half-doing it.

Acceptance criteria 1 (except nested subranges), 2, 3, 4, 5, 6, 7, 8, 9 and 11 are met, including both
manual Excel checks, the benchmark baseline and criterion 7's coverage figure (90.1%, above).
Criterion 10 is met except its gauge-corpus clause, as above.

**Manual Excel check: done, 2026-07-30.** The project owner opened both the template and the
generated report and confirmed they open without a repair dialog and render correctly. That covers
the two things reasoning could not settle: that Excel rebuilds a chart's `c:numCache` from the
formula after Task 6 drops it when re-pointing a series, and that Task 7's re-pointed pivot cache
source is one Excel accepts — upstream #200 being a four-year history of pivot output it refuses,
the validator passing was necessary and not sufficient.

The workbook pair was produced by a throwaway generator the first time. It is now the
`EverythingAtOnce` example (Task 11), so the check is repeatable rather than reconstructed from scratch
next time the write path is touched:

```
dotnet run --project XLibur.Report.Examples -- EverythingAtOnce --out c:\temp\check
```

It exercises grouping with subtotals and an outline, merged group labels, a per-row formula, one
conditional-formatting rule over a generated block, a picture below the range, a chart series
re-pointed from one row to twelve, an area-sourced pivot re-pointed, a table-sourced pivot refreshed
without being re-pointed, and a `<<Pivot>>`-generated pivot whose regional totals can be read against
the drawn one's. It prints what the generated file says when read back, then a thirteen-point checklist
for Excel — the file being schema-valid is necessary and, per upstream #200, nowhere near sufficient.

## Risks

- **Scriban syntax vs user expectations.** Authors coming from ClosedXML.Report lose C#
  method calls and lambdas (`item.Name.Substring(0,3)` → `item.Name | string.slice 0 3`;
  LINQ → `array.filter`/`map` or the function bridge). Mitigation: docs migration page, and
  the `XLibur.Report.DynamicLinq` package runs upstream-syntax templates
  unmodified for those who want it. **Both mitigations are in place**: the migration table is in
  `docs/report-templating.md`, and the package landed with Task 10.
- **Two engines double the behavioural surface.** Tag-parameter expressions (`over=`,
  `source=`, `test=`) evaluate through whichever engine the template uses, so every tag has
  two syntax audiences. Mitigation: tags receive evaluated *values* through the seam, never
  engine-specific ASTs; the parameterized fixture suite keeps structural behaviour identical
  across engines. **Held as built**: `BothEnginesTests` generates the same report under each engine and
  asserts the two agree cell for cell, tags and all — including `<<If test=…>>`, the tag most exposed to
  the difference.
- ~~**Buffer-sheet performance at scale.**~~ Upstream shows cell-by-cell buffer rendering strains
  past ~100K rows (#341, #68). There is no buffer sheet here (finding 1), but the risk landed anyway in
  a different form and the Task 9 benchmark caught it exactly as intended — see finding 10. Now flat to
  100K rows.
- **Pivot re-point depends on internal `XLPivotCache.Source` semantics** (area vs table vs
  name sources) — characterized and now relied on by `PivotRewriter`. If those internals shift
  under spec work in the core, the IVT coupling makes `XLibur.Report` a same-repo build break —
  visible immediately in CI, which is the point of same-repo versioning.
- ~~**Picture-anchor behaviour on row insert is unverified**~~ — resolved by Task 6's
  characterization tests: anchors are live ranges and move on their own. The risk that replaced it
  was the one nobody listed — that chart references could not be *written back* at all.
- **Tag-in-cell parsing ambiguity** (a literal `<<` in report text). Kept upstream-compatible;
  documented escape hatch if it bites.

## References

- ClosedXML.Report source (develop @ 2025-05-22): `XLTemplate`, `RangeInterpreter`,
  `RangeTemplate`, `TempSheetBuffer`, `FormulaEvaluator`, `Options/*`, `Excel/Subtotal.cs` —
  https://github.com/ClosedXML/ClosedXML.Report ; docs
  https://closedxml.io/ClosedXML.Report/docs/en/
- Upstream issues driving the gap-fills: #123/#351 (charts), #200/#399 (pivots), #216/#355
  (conditional formatting), #354/#281/#249 (pictures), #93, #225, #340, #341, #303.
- Scriban: https://github.com/scriban/scriban — `ScriptMode.ScriptOnly`, `Template.Evaluate`,
  ScriptObject import, safe-runtime docs; DoS advisories fixed in 6.6.0
  (GHSA-wgh7-7m3c-fx25, GHSA-xw6w-9jjh-p9cr, GHSA-xcx6-vp38-8hr5).
- System.Linq.Dynamic.Core: https://github.com/zzzprojects/System.Linq.Dynamic.Core —
  upstream pins 1.6.0.2; CVE-2023-32571 (arbitrary method invocation) fixed in 1.3.0.
- XLibur seams: `XLibur/Excel/CalcEngine/FunctionRegistry.cs` (internal),
  `XLibur/Excel/PivotTables/XLPivotCache.cs` (`Source`, `Refresh`, `RefreshDataOnOpen`),
  `XLibur/Excel/Charts/IXLChartSeries.cs` (`ValueReferences`/`CategoryReferences`),
  `XLibur/Properties/AssemblyInfo.cs` (IVT grants), spec 05 (ClosedXML.Parser reference
  shifting), spec 07 (function registry), spec 10 (chart write path).

## Implementation notes

Written after Tasks 1–5 landed and updated when grouping did, for whoever picks up the rest.
Everything here was established by running the code, not by reading it.

### The code as it stands

`XLibur.Report/`

| File | Role |
|---|---|
| `XLTemplate.cs` | Public entry point. Holds variables, owns the engine, registers the function bridge in its constructor, delegates generation to `RangeInterpreter`. |
| `IXLTemplate.cs`, `XLGenerateResult.cs`, `TemplateError(s).cs` | Public surface and the error model. Errors are collected, never thrown. |
| `ExpressionText.cs` | Recognises `{{ }}` and the `&=` formula prefix. `TryGetSingleExpression` is what decides whether a cell keeps its value's type. |
| `Expressions/` | `IExpressionEngine` (the seam), `ExpressionScope` (parent-linked name lookup), `ScribanExpressionEngine`, `ExpressionEvaluationException`. |
| `Excel/ReportValueConverter.cs` | `object?` → `XLCellValue`. |
| `Functions/ExcelFunctionBridge.cs` | Registers every calc-engine function under its upper-case name. |
| `Ranges/RangeInterpreter.cs` | Orchestrates: resolve bound ranges → evaluate cells outside them → expand each. Exposes `Expansions` (the ledger). |
| `Ranges/RangeBinder.cs` | Defined name → collection, including `Parent_Child` property paths. |
| `Ranges/RangeExpander.cs` | The heart. Read tags → capture column expressions → transform items → insert rows → copy → evaluate → restore CFs → run tags → drop the options row → re-point the name. |
| `Ranges/BoundRange.cs` | `BoundRange` and `RangeArea` (a plain row/column rectangle). |
| `Ranges/RangeAxis.cs` | Which way a range repeats, and every sheet operation that depends on it. The slot-and-line vocabulary is defined here. |
| `Ranges/GroupRenderer.cs` | Grouping. `Prepare` orders items by the levels' keys before any row exists; `Render` inserts the subtotal rows and outlines the block afterwards. |
| `Ranges/GroupOptions.cs` | The range-wide grouping options, read out of the `<<SummaryAbove>>`-style tags. |
| `Tags/` | `OptionTag` + `ProcessingContext`, `TagsRegister`, `TagParser`/`TagToken`, and the built-in tags. |
| `Tags/IRangeSummaryTag.cs` | The seam a summary is written through, so the options row's total and each group's use one implementation. |
| `Tags/IfTag.cs` | Conditional inclusion, at row scale in a repeated row and at range scale in the options row. |
| `Expressions/ExpressionTruth.cs` | The one rule for reading a value as a yes or a no, shared by every tag that asks a question. |
| `Rewriting/ExpansionMap.cs` | Where a row ends up after an expansion, shared by everything that refers to a range by address. |
| `Rewriting/ReferenceRewriter.cs` | Consumes the expansion ledger and re-points chart series. No picture code — see below. |
| `Rewriting/SheetReference.cs` | Takes a sheet-qualified A1 reference apart and puts it back. Refuses the forms it cannot safely rewrite. Read by the chart rewriter and by `<<Pivot dest=…>>`. |
| `Rewriting/PivotRewriter.cs` | Re-points area-sourced pivot caches, refreshes every touched cache, and moves pivot tables out of the way. |
| `Rewriting/PivotBuilder.cs` | Builds the pivot tables `<<Pivot>>` asked for, after every range has settled. `PivotRequest`/`PivotField` are the request it holds until then. |
| `Tags/PivotTag.cs` | `<<Pivot>>` and the four field tags. Resolves its destination and its fields to live ranges before expansion, so template coordinates stay template coordinates. |

`XLibur.Report.Tests/Infrastructure/` holds `WorkbookComparer` (semantic diff),
`ReportFixture`/`GoldenFile` (fixture runner), `ReportResources` (template files, regeneration).

`XLibur.Report.Examples/` is ten worked examples behind one `ReportExample` base class: an example says
what its template looks like and what data it binds, and the base class saves the pair, validates both
and prints what the generated workbook says when read back. `AllExamples.Ordered` is the menu, and
`ExampleSmokeTests` runs every one of them. See its `README.md`.

`XLibur.Report.DynamicLinq/` is one class. `DynamicLinqExpressionEngine` flattens the scope chain into
the flat parameter list Dynamic LINQ compiles against — which is where the shadowing a scope chain
expresses has to be resolved, and where the `@`-prefixed aliases are added — and caches the compiled
lambda under the expression text **plus the parameter names and types**, because the same text against a
different scope shape is a different lambda. Two details earned their comments: a parameter is compiled
against the value's *concrete* type, without which `item.Name.ToUpper()` cannot resolve; and a
non-generic `IEnumerable` is re-cast to the element type its first item reports and materialised, without
which LINQ has no element type to work with.

`XLibur.Report.Benchmarks/` holds three instruments, and which to reach for matters:
`ScalingProbe` (`-- scaling`) times one generation per row count and reports **cost per row**, which is
the only cheap way to see the shape of the curve — it is what found and then confirmed the fix for
finding 10; `ExpansionPhaseProbe` (`-- phases`) splits expansion into its copy and evaluate halves using
no report code at all, to tell a core-library scaling property from an engine one; and
`ReportGenerateBenchmarks` is the BenchmarkDotNet suite for comparing implementations at one size, which
costs the better part of an hour to run in full.

### Running it

```
dotnet build XLibur.slnx -c Release -v q
XLibur.Report.Tests/bin/Release/net10.0/XLibur.Report.Tests.exe
XLibur.Report.Tests/bin/Release/net10.0/XLibur.Report.Tests.exe --treenode-filter "/*/*/TagBehaviourTests/*"
XLIBUR_REPORT_REGEN=1 XLibur.Report.Tests/bin/.../XLibur.Report.Tests.exe   # rewrite fixture templates
```

`--treenode-filter`, **not** `--filter` — MTP ignores the latter and reports "Zero tests ran".

**`dotnet test` currently discovers nothing here** — every invocation reports "Zero tests ran", exit
code 5, in ~170 ms. It is not this package: `XLibur.Tests` does the same, filtered or not. Running
the built test executable directly discovers and runs the whole suite normally, which is what the
commands above do. Worth re-checking against a newer SDK before spending time on it.

### Core APIs this depends on, that are not obvious from their signatures

- **`IXLConditionalFormat.Ranges` is a fresh projection.** Mutating the returned collection does
  nothing. Rewrite coverage with `((XLConditionalFormat)format).SetAreas(XLAreaList.FromRanges(...))`
  (`XLibur.Excel.ConditionalFormats`, `XLibur.Excel.Coordinates`). The public `Range` setter handles
  the single-area case.
- **`RangeUsed()` defaults to contents.** A cell differing only in style, hyperlink or merge falls
  outside it. `WorkbookComparer` uses `RangeUsed(XLCellsUsedOptions.All)`; two of its own tests
  failed until it did.
- **`IXLRange.CopyTo` costs more the larger the target worksheet is**, independently of how much is
  being copied: measured tripling per call over 30,000 rows, by `ExpansionPhaseProbe`, with no report
  code involved. Any per-item call to it is therefore quadratic in the item count — which is what
  finding 10 was. Copy in as few calls as the work allows. **Now diagnosed and specced separately**: the
  cost is `XLRangeBase.Clear` creating and deleting a data validation on every call even when the sheet
  has none, which is ~85% of `CopyTo`'s cost and all of its growth. See spec 14 and
  [#271](https://github.com/XLibur/XLibur/issues/271); a two-line guard takes `CopyTo` from 420 µs to
  13 µs at 30,000 rows. The report package's doubling workaround stands either way.
- **Calling a calc-engine function**: `FunctionDefinition.CallFunction(CalcContext, Span<AnyValue>)`.
  Construct `new CalcContext(engine, culture, workbook: null, worksheet: null, formulaAddress: null)`
  for a no-grid evaluation; functions needing a grid throw `MissingContextException`
  (`XLibur.Excel.CalcEngine.Exceptions`). Convert with `AnyValue.From(...)` in and
  `AnyValue.TryPickScalar` + `ScalarValue.Match(...)` out; blank is `ScalarValue.Blank.ToAnyValue()`.
- **Charts**: `IXLWorksheet.Charts` is an `IEnumerable<IXLChart>`; a chart has `Series` and
  `SecondarySeries`, and a series' `ValueReferences`/`CategoryReferences` are plain sheet-qualified
  A1 strings (`Data!$B$3:$B$8`), not ranges. `IXLChartSeriesCollection.Add` **throws** for a chart
  loaded from a file, so a rewriter can only edit the series a template already has. Nothing shifts
  a chart's own anchor on row insert either — `IXLDrawingPosition.Row` is a plain `int`.
- **Defined names**: enumerate `DefinedNames.ValidNamedRanges()` on both the workbook and each
  worksheet; re-point with `SetRefersTo(range)`; `Delete()` removes one. Row inserts and deletes
  shift and shrink names automatically, which the expander relies on to process several ranges on
  one sheet top to bottom.
- **`XLHelper.GetColumnLetterFromNumber` / `GetColumnNumberFromLetter`** for column arithmetic; the
  latter throws `ArgumentException` on nonsense.
- **Scriban**: `LexerOptions { Mode = ScriptMode.ScriptOnly }` for a bare expression; one
  `TemplateContext` reused across evaluations with `PushGlobal`/`PopGlobal` per scope;
  `PushCulture` for formatting. Pushing the function `ScriptObject` **once, by reference** is what
  makes functions registered after the first evaluation visible. A delegate whose last parameter is
  `params object?[]` registers as a variadic function — that is how the bridge fits through
  `AddFunction(string, Delegate)`.

### Conventions this code established

- **The options row is the last row of a multi-row range**; a single-row range has none. It is
  removed after generation only if nothing is left in it, so a row holding a total survives and a
  row holding only tags does not. Tag text is `Clear`ed rather than set to `""`, because an empty
  string still counts as content.
- **Tags act at two named moments** — `TransformItems` before any row exists, `Execute` after — 
  rather than relying on priority alone to imply ordering. Priority still orders within a moment
  (`If` is 1 so nothing sorts or groups a row that was dropped; `Delete` is 250 so a column can be
  sorted by and then removed).
- **A tag's placement carries meaning.** The options row describes the range; a repeated row
  describes a row. `OptionTag.InRepeatedRow` is how a tag that means something at both scales tells
  which was meant, and it is why the expander reads tags from every row of the range rather than
  only the last.
- **A tag parameter may be an expression.** `ProcessingContext.IsTrue` takes a bare expression, an
  interpolated one or a literal, so a template author writes whichever reads best.
- **A column-placed tag learns its column's meaning** from `ProcessingContext.LineExpressions`,
  captured before evaluation overwrites the template text. That is why `<<Sort>>` and `<<Group>>`
  need no `by`.
- **A tag can see what else the range asked for**, through `ProcessingContext.Tags`. That is how
  `<<Pivot>>` finds the field tags that describe its columns, and how it refuses to run at all when a
  `<<Group>>` is present.
- **Nothing throws out of generation.** Failures become `TemplateError`s and, for cells, red text
  in the offending cell.
- **A grouped range is ordered by the engine, stably.** `<<Group>>` alone is enough to make its
  groups contiguous; because the ordering is stable, a `<<Sort>>` on another column still decides
  the order within a group, and `nosort` opts out for data that arrives arranged.
- **A subtotal row takes the options row's cell styling.** It is the only styling a template has a
  way to express for a row that does not exist until generation, and it is what makes a group total
  and the grand total look alike.
- **Group subtotals and the grand total are the same declaration.** `<<Sum>>` in the options row is
  written into every group's subtotal row over that group's rows and again into the options row over
  the lot; `SUBTOTAL` ignoring nested `SUBTOTAL`s is what stops the grand total counting the data
  twice, and `<<DisableGrandTotal>>` is how a report keeps the group totals without a report total.

### Notes for the remaining tasks

- **A pivot's source cannot be set to a defined name through the public API.** `XLPivotCache.Source`
  is internal, so `IXLPivotTables.Add` only ever gives an area source or — when the range is exactly
  a table's area — a table source. A template author picks a named source in Excel, not in code, so
  a *code-built* template demonstrating the refresh-only branch has to use a table. The engine
  treats name and table sources identically, so nothing is lost; it is the docs and the Task 11
  examples that need to know.
- **A template address only stays a template address if the core holds it.** `<<Pivot>>` keeps its
  destination and its fields' heading cells as live one-cell ranges from before expansion, which is
  what makes `dest="A10"` mean row 10 *of the template* however many rows the report generated, and
  what makes `<<Delete>>` remove a column and its pivot field together. Anything else that needs to
  refer to a template address after generation should do the same rather than compute where it went;
  `RangeShiftMechanicsCharacterizationTests` records exactly what the core does and does not carry.
- **Anything whose position is a rectangle has to be created last.** A pivot table's position and its
  cache source are both plain values, so `PivotBuilder` runs at the end of a generation run rather
  than in the tag. It runs *after* the rewriters on purpose: a generated pivot then shares a cache
  with a template pivot re-pointed to the same rows, and nothing re-points what it just built.
  `PivotTagTests.AGeneratedPivotIsNotRePointedASecondTime` fails if that order is disturbed.
- **Horizontal ranges are the axis, not a flag.** `RangeAxis` names what varies — a **slot** is what
  repeats (a row, or a column), a **line** is the other direction — and the expander, the tags, the
  ledger and both rewriters are written against it. Adding a third orientation, if one ever made
  sense, means a third subclass and nothing else. The vertical suite passing unchanged through the
  refactor is what says the abstraction is faithful rather than merely plausible.
- **Task 3b (nested subranges).** `RangeBinder` resolves `Parent_Child` paths from *workbook*
  variables. A child range inside a parent's rows needs resolving per parent item instead, and
  expanding inside each copied block — before the parent's own evaluation pass, or the child's
  `{{ }}` cells will be evaluated against the parent's scope.

### Environment gotchas

- TUnit's global usings make bare `Assembly` ambiguous with `HookType.Assembly`; qualify
  `System.Reflection.Assembly`.
- The `TUnitAssertions0015` analyser rejects `.IsEqualTo(true)`; use `.IsTrue()`. It is a warning
  locally but `TreatWarningsAsErrors` makes it fail the Release build.
- Golden-file templates are read from the **source tree**, not embedded resources: a regeneration
  run writes a template and reads it back in the same pass, which an embedded copy cannot satisfy
  until the next build.
- Commit messages here contain parentheses and quotes that break shell quoting — write the message
  to a file and use `git commit -F`.
