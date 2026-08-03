---
id: migration
title: Migration from ClosedXML
sidebar_label: Migration from ClosedXML
sidebar_position: 3
description: Move an existing ClosedXML project to XLibur — namespace changes, font engine packaging, and every breaking or behaviour change since the 0.105 fork, each with its mitigation.
---

# Migration from ClosedXML

XLibur was forked from [ClosedXML v0.105.0](https://github.com/ClosedXML/ClosedXML/), and the
public API surface is largely unchanged. To migrate:

1. Install the NuGet package (see [Getting Started](./getting-started.md#installation)).
2. Replace `using ClosedXML` namespace references with `using XLibur`.
3. Read [Breaking API changes](#breaking-api-changes) below — there are seven, and most projects
   hit none of them.

Namespaces are prefixed with `XLibur`, so both libraries can be referenced from the same project
while you port.

Everything on this page is relative to ClosedXML 0.105. The
[changelog](https://github.com/XLibur/XLibur/blob/main/CHANGELOG.md) records the same ground
release by release, with the pull request behind each entry.

## Font engine configuration (different from ClosedXML)

This is the one area where XLibur's packaging differs from the base ClosedXML package. ClosedXML
bundles [SixLabors.Fonts](https://github.com/SixLabors/Fonts) directly into its core assembly for text
measurement (column auto-fit, row heights, glyph metrics). XLibur instead keeps the **core assembly
free of any font library** and ships the font engine as a **separate, swappable package**. This lets
you pick a font library with a license that suits you and avoids forcing a font dependency on library
authors who don't need one.

What this means when migrating:

- **Install `XLibur.Bundle` (or `XLibur` + `XLibur.Fonts.SkiaSharp`) and no code changes are needed.**
  The default [SkiaSharp](https://github.com/mono/SkiaSharp) engine (MIT-licensed) is **auto-registered
  by XLibur core the first time you create a workbook** — there is no startup call to add:

  ```csharp
  using var wb = new XLWorkbook(); // font engine resolved automatically
  ```

  The default resolves system fonts and falls back to an embedded, metric-only Calibri-compatible font,
  so text measurement works even in headless/serverless environments with no system fonts installed.

- **If you install the bare `XLibur` package with no font engine**, creating a workbook throws an
  `InvalidOperationException` telling you to add a font engine package. This is intentional — it's how
  the core stays font-library-agnostic.

- **To choose a different engine**, install its package and either register it at startup (it takes
  precedence over the auto-registered default) or pass it per workbook:

  | Package | Font library | License | Notes |
  |---|---|---|---|
  | `XLibur.Fonts.SkiaSharp` | SkiaSharp | MIT | **Default.** Auto-registers; ships native binaries. |
  | `XLibur.Fonts.SixLabors.V1` | SixLabors.Fonts 1.x | Apache 2.0 | Pure-managed; matches ClosedXML 0.105's engine. |
  | `XLibur.Fonts.SixLabors` | SixLabors.Fonts 2.x | Six Labors Split License | Commercial restrictions over $1M revenue. |

  ```csharp
  // Override globally at startup (e.g. keep ClosedXML's SixLabors 1.x behavior):
  SixLaborsV1FontBootstrap.Register();

  // Or override per workbook:
  var options = new LoadOptions { FontEngine = new SkiaSharpFontEngine("Arial") };
  using var wb = new XLWorkbook(options);
  ```

Resolution order for the font engine is: `LoadOptions.FontEngine` (per workbook) →
`LoadOptions.DefaultFontEngine` (explicitly registered global) → the auto-registered default engine.
See [Fonts and Font Engines](./fonts.md) for the full picture, including headless environments and
loading fonts from streams, or
[docs/font-architecture.md](https://github.com/XLibur/XLibur/blob/main/docs/font-architecture.md)
for the design rationale.

## Breaking API changes

Seven changes since 0.105 can break a build or throw where the old code returned. At a glance:

| Change | ClosedXML 0.105 | XLibur | What to do |
|---|---|---|---|
| [`IXLRange.Cell(string)`](#ixlrangecellstring-throws-instead-of-returning-null) | returned `null` for an unresolvable address | throws `ArgumentException` | Catch `ArgumentException` where you tested for `null` |
| [Range address validation](#a-bad-range-address-throws-a-typed-exception) | `IndexOutOfRangeException` / `NullReferenceException` | `ArgumentNullException`, `ArgumentException`, `FormatException` | Catch the typed exception instead |
| [`XLColorType` ordinals](#xlcolortype-members-are-renumbered) | `Color`=0, `Theme`=1, `Indexed`=2 | `Automatic`=0, `Color`=1, `Theme`=2, `Indexed`=3 | Remap any persisted numeric value; recompile against XLibur |
| [`XLColor.NoColor.Color`](#reading-a-component-off-an-automatic-colour-throws) | returned ARGB `(0,0,0,0)` | throws | Test `IsAutomatic` before reading `Color`/`Indexed`/`ThemeColor` |
| [`XLColor.NoColor`](#xlcolornocolor-is-deprecated) | current | `[Obsolete]`, aliases `XLColor.Automatic` | Rename to `XLColor.Automatic` |
| [`XLError` gained `SpillRange`](#xlerror-gained-a-member) | 7 members | 8 members (`SpillRange = 7`) | Handle the new member in exhaustive `switch`es |
| [`IXLPivotCache` / `IXLConditionalFormat`](#two-interfaces-gained-members) | — | new members | Only affects types outside XLibur that *implement* these interfaces |

### `IXLRange.Cell(string)` throws instead of returning null

The interface has always been annotated non-null, so the old `null` contradicted its own
signature and usually surfaced as a `NullReferenceException` somewhere downstream. It now fails
at the call site, matching `IXLWorksheet.Cell(string)`, which already behaved this way.

```csharp
// ClosedXML 0.105 — a null return was the only signal the address was unresolvable
var cell = range.Cell(name);
if (cell is null)
    return;

// XLibur
try
{
    var cell = range.Cell(name);
}
catch (ArgumentException)
{
    return;
}
```

This is not a mechanical swap. If you relied on the `null` to mean "not found", the equivalent is
catching `ArgumentException`; code that trusted the non-null annotation needs no change at all.
`IXLRange.Range(string)` gained the same guard, but nothing reaches it — a bad address already
throws while being parsed.

### A bad range address throws a typed exception

`Range("")` and half-written addresses such as `"A1:"`, `":"` and `"$"` used to throw
`IndexOutOfRangeException`, and a null address threw `NullReferenceException` — internal failures
escaping a public API rather than anything a caller could act on. Each now throws something a
caller can catch deliberately:

| Address | XLibur throws |
|---|---|
| `null` | `ArgumentNullException` |
| `""` and `" "` | `ArgumentException` |
| `"A1:"`, `":"`, `"$"`, `"not an address"` | `FormatException` |
| past the sheet limits | `OverflowException` |
| an unknown defined name | `ArgumentOutOfRangeException` |

This affects every entry point that parses an address string, including `IXLWorksheet.Range`,
`IXLWorkbook.Range` and `IXLRange.Range`. If you were catching `IndexOutOfRangeException` or
guarding against a `NullReferenceException` around an address parse, switch to
`FormatException` / `ArgumentException`.

### `XLColorType` members are renumbered

OOXML has four kinds of colour and ClosedXML modelled three. The automatic colour (ECMA-376
`CT_Color/@auto`, what Excel's font colour picker labels *Automatic*) was disguised as a fully
transparent black; it is now a member in its own right, and takes ordinal 0 so that a default
colour key is automatic rather than a transparent black no Excel file means to express.

```text
Color   0 → 1
Theme   1 → 2
Indexed 2 → 3
```

This is source compatible — recompiling is enough — but **binary breaking**, and it breaks
anything that persisted the numeric value. Stored settings or serialised styles written by
ClosedXML need remapping on read.

### Reading a component off an automatic colour throws

`XLColor.Automatic.Color` (and `.Indexed` / `.ThemeColor`) throw rather than returning a
meaningless all-zero ARGB. There is no ARGB that means "automatic" — the application resolves it
at display time — so the property refuses to invent one.

```csharp
// ClosedXML 0.105 — returned Color.FromArgb(0, 0, 0, 0) for an automatic colour
var rgb = cell.Style.Font.FontColor.Color;

// XLibur
var colour = cell.Style.Font.FontColor;
var rgb = colour.IsAutomatic ? defaultRgb : colour.Color;
```

Decide what your code should render for a colour Excel leaves to the application — generally black
for a font or border, white for a fill.

### `XLColor.NoColor` is deprecated

Some Excel pickers (sheet tab colour, fill background) label the same value *No Color*; that is a
GUI convention, not a different value. `NoColor` still compiles and still returns `Automatic`, but
now warns — which is an **error** for consumers building with `TreatWarningsAsErrors`.

```csharp
cell.Style.Font.FontColor = XLColor.NoColor;    // before
cell.Style.Font.FontColor = XLColor.Automatic;  // after — the same value
```

### `XLError` gained a member

`XLError.SpillRange` (`#SPILL!`, ordinal 7) is appended for the new
[dynamic-array spill engine](./formulas.md). Existing ordinals are unchanged, so nothing persisted
needs remapping — but a `switch` over `XLError` that was exhaustive under 0.105 is no longer
exhaustive, and a `switch` *expression* will throw at runtime rather than fail to compile. Add an
arm for it, or a discard.

### Two interfaces gained members

`IXLPivotCache` gained `SourceKind`, `SourceRange`, `SourceName`, `SourceWorksheet` and
`SetSourceRange`; `IXLConditionalFormat` gained `SetRanges`. Adding a member to a public interface
is source-breaking for any type outside the library that implements it.

Neither interface is designed to be implemented externally — each has a single implementation with
an internal constructor — so in practice this affects test doubles. Consumers that only *use* these
interfaces are unaffected, and gain the new members. See
[Pivot tables](./pivot-tables.md) and [Conditional formatting](./conditional-formatting.md).

## Deprecations

Three groups of `[Obsolete]` members that shipped in ClosedXML 0.105 are **removed in the next
minor version**, so port them now rather than on the upgrade after this one:

| Deprecated | Replacement |
|---|---|
| `IXLWorkbook.NamedRanges`, `IXLWorksheet.NamedRanges`, `IXLDefinedNames.NamedRange` | `DefinedNames` / `DefinedName` |
| `IXLCell.DataValidation`, `IXLRangeBase.DataValidation` | `GetDataValidation()` to read, `CreateDataValidation()` to create |
| `IXLRanges.SetDataValidation()` | `CreateDataValidation()` |
| `XLFontCharSet.Hangeul` | `XLFontCharSet.Hangul` |

`XLColor.NoColor` is the exception: it is deprecated but stays.

One interface is newly deprecated in XLibur. `IXLBaseCollection<TSingle, TMultiple>` is an orphan —
nothing in the library implements, extends or consumes it, and the collections it looks like it
should describe (`IXLColumns`, `IXLRows`, `IXLCells`, `IXLRangeColumns`, `IXLRangeRows`) all derive
from `IEnumerable<T>` alone. Use whichever of those you actually hold.

## Behaviour changes that need no code change

These compile unchanged but can produce a different result or a different file than ClosedXML
0.105 did. Almost all of them are bug fixes — the old behaviour was silently wrong — but if you
pin output with byte-comparison tests or golden files, expect those to move.

### Formulas and references

- **A reference whose rows or columns are all deleted becomes `#REF!`.** Endpoints used to be
  clamped to row 1, so deleting rows 1–5 turned `Sheet1!$A$1:$B$2` into a phantom one-row range
  over whatever data had moved up into it. `=SUM(A1:A2)` with those rows deleted now reads
  `SUM(#REF!)` instead of quietly summing the wrong cells.
- **Row-only and column-only references shift correctly.** `3:5` with row 4 deleted became `2:4` —
  a reference that had walked onto a row it never covered. Inserting two rows at row 4 moved `3:5`
  to `5:7` rather than expanding it to `3:7`. Both axes now follow the same boundary rules as an
  equivalent cell range.
- **A deletion that removes the tail of a reference no longer inverts it.** `3:5` with rows 5–7
  deleted came back as `3:2`, which is not a valid formula; `A2:A8` with rows 5–9 deleted came back
  as `A2:A3`, dropping row 4, which survived.
- **`=SUM(B2:A1)` evaluates** instead of throwing `ArgumentException("Range address must be
  normalized")`. Each axis is ordered independently, carrying its own fixed marker.
- **`SUBTOTAL` no longer counts a nested post-2007 function twice.** Excel stores every function
  added after 2007 under an `_xlfn.` namespace, so a nested `AGGREGATE` was never recognised by the
  check that stops a subtotal counting a subtotal. Totals that were wrong are now right.
- **A formula reading a table keeps up with the table's contents.** `=SUM(Table1[Amount])`
  registered no precedents, so nothing invalidated it and it served a stale cached value until a
  full recalculation. A structured reference naming a table on another sheet is now resolved
  against *that* sheet rather than the calling one, where it previously read the same coordinates
  on the wrong sheet and quietly returned 0.
- **A defined name holding a structured reference resolves to what the reference says.**
  `Sales[[#Headers],[Amount]]` pointed at the data instead of the header, a column span
  `Sales[[Amount]:[Tax]]` lost everything but its first column, and `Sales[#All]` resolved to
  nothing. An unknown column threw `ArgumentOutOfRangeException` out of a property getter —
  reachable just by loading a workbook whose table column had since been renamed — and now
  contributes no range instead. `Sales[Amount]`, the common form, resolves as before, so any
  `try`/`catch` you wrapped around `IXLDefinedName.Ranges` can go.
- **Copying a worksheet repoints the copy's self-references at the copy.** Copying a sheet named
  `Original` holding `Original!A1 * 3` produced a sheet whose formula still pointed at the
  original. References to *other* sheets are left alone.
- **Named ranges shrink correctly when their first row or column is deleted.** `A3:A4` became
  `A2:A3`, expanding the range over a row that was never part of it; it is now `A3:A3`, as Excel
  produces.
- **Array and dynamic-array formulas survive row and column shifts.** A shift used to rebuild every
  formula cell through the `FormulaA1` setter, splitting one shared array formula into a normal
  formula per cell — and a spilled `=UNIQUE(...)` into several implicit-intersection
  `=@UNIQUE(...)` cells, even when the edit happened on an unrelated sheet.

:::note Dynamic arrays now spill
XLibur adds `SEQUENCE`, `UNIQUE`, `SORT`, `SORTBY`, `FILTER`, `XLOOKUP` and `XMATCH` together with
a spill engine, so one of these written into a single cell auto-fills its computed footprint into
the neighbouring cells. A footprint blocked by existing content, or one running past the sheet
edge, collapses to `#SPILL!` on the anchor. ClosedXML 0.105 had none of these functions, so this
cannot break existing formulas — but a workbook authored in Excel that uses them behaves
differently once XLibur can evaluate it. See [Formulas](./formulas.md).
:::

### Text and number parsing

**Coercion is stricter, and rejects things the BCL parse accepted.** Group-separator and currency
placement are now enforced the way Excel enforces them, so under `en-US` the strings `1,00`,
`1,00,000` and `1$` no longer coerce to a number. If you were relying on the looser behaviour,
parse the string yourself and assign the typed value:

```csharp
// Instead of letting a loosely-formatted string coerce
cell.Value = decimal.Parse(raw, NumberStyles.Any, culture);
```

In the other direction, coercion now succeeds where it used to fail: date-times carry a seconds
component (`8/22/2008 3:30:45 PM` failed entirely before), overflowing time components carry into
the date, parenthesised and sign-separated numbers such as `(100%)` and `- 100 %` read as negative,
a month matches on any prefix from three letters up, and a shortened or dot-suffixed AM/PM
designator is accepted.

### What gets written to the file

- **Cached formula values are preserved on save** whenever they exist and the formula has not been
  dirtied, regardless of `EvaluateFormulasBeforeSaving`, and the data-type attribute is kept. This
  fixes round-trip loss of dynamic-array results and spill cell values.
- **An automatic colour is written as `auto="1"`, not `rgb="00000000"`.** Colour writers switched on
  the colour type and the automatic colour fell into the RGB arm, pinning down a colour the source
  deliberately left to the application. The three conditional-format colour converters gained
  explicit automatic arms too, where they would otherwise have dropped the colour silently.
- **Rich-text runs keep an absent or automatic colour.** A plain load → `SaveAs` wrote every
  colour-less run back with an explicit `<color rgb="FF000000"/>`, which cannot then be overridden
  by a theme change or by conditional formatting. A run read with no `<rPr>` is written back
  without one.
- **Saving a plain shared string carrying a phonetic guide no longer throws** an `ArgumentException`
  naming an invalid XML character — common in Japanese workbooks.
- **Page breaks no longer inflate the used range.** `AddHorizontalPageBreak()` /
  `AddVerticalPageBreak()` wrote `brk@max` as the sheet's full row count, so a file with ~2,000
  rows of data rendered in Excel with a scrollbar spanning all 1,048,576.
- **Totals-row formulas escape column names containing spaces.** A header such as `Feb 2023`
  produced a structured reference Excel could not parse.
- **Chart XML passes OpenXML schema validation.** Three violations in the chart writer are fixed
  (a series name written as `c:strRef` with no `c:f`, a `c:doughnutChart` missing `c:holeSize`,
  and `c:marker` written after `c:cat`/`c:val`). Excel tolerated all three; stricter readers and
  `SaveOptions.ValidatePackage` did not.
- **Grouped pictures and shapes survive a round trip** instead of being dropped, and pictures
  nested in `xdr:grpSp` groups are now a first-class API.

### Data validation

- **A data validation is never written covering nothing.** Adding a validation over a range that
  wholly contained an existing rule left that rule with no coverage, and `ClearRanges` /
  `RemoveRange` let a caller empty one directly. The schema requires a non-empty `sqref`, and Excel
  reads `sqref=""` as corruption — it repairs the workbook and drops *every* validation on the
  sheet. A rule left with no coverage is now deleted, and the writer skips any rule covering
  nothing.
- **Validations no longer vanish when inserting at row 1 or column 1.** The index was keyed by
  address at insert time and never re-keyed, so Excel rejected the saved file with *"Removed
  Records: Data validation"*.
- **Criteria formulas are shifted with the sheet.** An insert or delete relocated each rule's
  `sqref` but left cell references inside `formula1`/`formula2` pointing at the pre-shift location,
  silently breaking any `List`, `Custom` or comparison rule that referenced other cells — most
  visibly dependent dropdown pairs driven by `OFFSET`/`MATCH`. The in-memory value was wrong
  immediately, before any save.

### Comments

- **A cell carries either a note or a threaded comment, never both.** Creating one over the other
  throws rather than silently discarding it. Threads previously read lossily — the whole
  conversation was flattened into the legacy note's text joined by newlines — and had no write path
  at all. See [Comments and hyperlinks](./comments-and-hyperlinks.md).
- **`IXLComment.Delete()` removes the note from where it is now, not where it was created.** A note
  remembered its construction cell, so deleting a note on `A5` after two rows were inserted above
  cleared `A5` — by then empty — and left the note sitting on `A7`.

### Charts

Charts were stubs in ClosedXML 0.105 and are now fully implemented across all 78 `XLChartType`
values, so most of this is new rather than changed. Two things to know if you load and re-save
files containing charts:

- **`Series.Add(...)` on a chart loaded from a file throws `NotSupportedException`.** A loaded chart
  is patched rather than regenerated, so a new series had nowhere to be written and used to vanish
  on save without a word.
- **Charts that used to be dropped on load are now read**, which changes what a round trip
  produces: one-cell and absolute anchors (previously only `xdr:twoCellAnchor` was read), 3D and
  of-pie chart groups, and second and subsequent plot groups of the same type — which is how Excel
  stores a secondary axis.

### Pivot tables and autofilter

- **Pivot field filters survive a round trip.** The reader skipped the `filters` element, so loading
  and saving silently un-filtered every pivot table in the workbook — a change to what the workbook
  *shows*, not just what it remembers. If a downstream process depended on that accidental
  un-filtering, it now sees the filtered view. (Not the report-filter axis, which is `pageFields`
  and was already supported.)
- **A PivotChart keeps its manual series and point formatting**, via the `chartFormats` collection
  that ties each formatting record to the pivot area it applies to.
- **Pivot table alignment in differential formats (DXF) round-trips** instead of being lost on load.
- **A worksheet autofilter keeps the parts XLibur does not model** — `iconFilter`, button
  attributes, `extLst`, the dynamic filter types beyond the two averages. A column that has not been
  changed is written back from the criteria it was loaded with; every mutation drops them, so an
  edit is never discarded.
- **Loading a relative-date filter no longer throws `KeyNotFoundException`.** Any of the ~38
  relative date types (`thisMonth`, `yearToDate`, `lastQuarter`, …) failed a two-entry map lookup.

### Conditional formatting

**A rule's ranges shift once, not twice.** Inserting rows or columns below the first line doubled
the shift for any rule whose shifted target address collided with another rule's existing range —
a rule at `K13` that should move to `K23` landed at `K33`, while rules whose targets happened to be
empty shifted correctly.
