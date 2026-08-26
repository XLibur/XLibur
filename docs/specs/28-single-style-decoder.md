# Spec 28 — One OOXML style decoder instead of three

**Area:** Architecture · Refactor · **Defect** (dxf fonts lose three fields on load)
**Effort:** M (~1 week)
**Dependencies:** None hard. File-disjoint from spec 23. Reads `XL*Key.cs`, which spec 20 may relayout
— see Conflicts.
**Status:** ✅ Merged — PR #411 as `17d74943`, 2026-08-26. See Results.

## Goal

Make one module the single implementation of "OOXML style XML → `XLStyleKey`", so which element the
XML arrived in stops deciding which fields get read. Delete the second decoder family, collapse three
`numFmtId` lookups into one, and take style decoding out of the sheet-data reader.

## Why this spec exists

`XLibur/Utils/OpenXmlHelper.cs` (513 lines) contains **two complete decoders for the same XML**,
chosen by provenance:

| Aspect | Mutating form — used for `<dxfs>` | Key form — used for `<cellXfs>` |
|---|---|---|
| Number format | `LoadNumberFormat:81` | `WorksheetSheetDataReader.LoadStyleNumberFormat:1299` |
| Alignment | `LoadAlignment:91` | `AlignmentToXLibur:284` |
| Border | `LoadBorder:117` | `BorderToXLibur:302` |
| Fill | `LoadFill:148` | *reuses `LoadFill` through a throwaway `XLFill`* — `WorksheetSheetDataReader:1272-1281` |
| Font | `LoadFont:202` | `FontToXLibur:358` |
| Protection | *not decoded at all* | `ProtectionToXLibur:410` |

The mutating form takes an `IXLFontBase` / `IXLBorder` / `IXLFill` / `IXLNumberFormat` /
`IXLAlignment` and writes through it. The key form takes and returns an immutable `XL*Key`. Five
public entry points and nine private helpers on one side (`OpenXmlHelper.cs:81-282`), four public
entry points and one private helper on the other (`:284-420`).

Fill is the tell. The key path has no fill decoder of its own; `LoadStyleFill` allocates an `XLFill`,
mutates it through `OpenXmlHelper.LoadFill`, and takes `.Key`
(`WorksheetSheetDataReader.cs:1272-1281`). One implementation, reached from both provenances, through
one allocation. That is the shape this spec generalises to the other five aspects.

### Divergence 1 — a dxf font loses its name, family and charset

This is the headline, and it is worse than a single dropped field.

`FontToXLibur` reads the typed properties of a `Font`:

```csharp
var fontName = f.FontName?.Val?.Value ?? string.Empty;          // :392
var fontFamilyNumbering = f.FontFamilyNumbering?.Val?.Value;    // :396
var fontCharSet = f.FontCharSet?.Val?.Value;                    // :400
```

`LoadFont` takes an untyped `OpenXmlElement` and searches by child type:

```csharp
var runFont = fontSource.Elements<RunFont>().FirstOrDefault();       // :234
var fontFamilyNumbering = fontSource.Elements<FontFamily>()...;      // :226
// nothing looks for a charset at all
```

Those are the **rich-text** spellings, not the **font** spellings, and the CLR types are unrelated.
Verified by reflection over `DocumentFormat.OpenXml.dll` 3.4.1:

```
FontName            : OpenXmlLeafElement          <x:name>     child of <x:font>
RunFont             : OpenXmlLeafElement          <x:rFont>    child of <x:rPr>
FontFamilyNumbering : OpenXmlLeafElement          <x:family>   child of <x:font>
FontFamily          : InternationalPropertyType   <x:family>   child of <x:rPr>
FontCharSet         : OpenXmlLeafElement          <x:charset>  child of <x:font>
RunPropertyCharSet  : InternationalPropertyType   <x:charset>  child of <x:rPr>
```

No inheritance links any pair. `Elements<RunFont>()` over a `Font` returns nothing, because that
element's name child is a `FontName` instance.

The three dxf call sites all pass a `Font` (`DifferentialFormat.Font`). So **`LoadFont` applied to a
dxf silently drops the font name, the font family numbering and the charset** — three of the fifteen
children `Font` can carry. The nine it does read (bold, italic, shadow, strikethrough, colour, size,
underline, vertical alignment, scheme) use types shared by both schemas, which is why the omission has
never shown up as a crash.

The writer emits all three. `AppendFontScalarElements` (`WorkbookStylesPartWriter.cs:1123-1130`) is
shared by cell fonts and dxf fonts alike:

```csharp
if (f.FontName != d.FontName || ignoreMod)
    font.AppendChild(new FontName { Val = f.FontName });
if (f.FontFamilyNumbering != d.FontFamilyNumbering || ignoreMod)
    font.AppendChild(new FontFamilyNumbering { Val = (int)f.FontFamilyNumbering });
if ((f.FontCharSet != d.FontCharSet || ignoreMod) && f.FontCharSet != XLFontCharSet.Default)
    font.AppendChild(new FontCharSet { Val = (int)f.FontCharSet });
```

**User-visible effect, stated precisely:** set a conditional format's font to `Arial` with
`XLFontCharSet.Arabic`, save, reopen. The `<dxf><font>` in the saved file carries `<name val="Arial"/>`
and `<charset val="178"/>` — the file is correct and Excel renders it correctly. XLibur reads the
format back with the workbook default font name (`Calibri`) and `XLFontCharSet.Default`. The next save
writes the degraded form, so the second round trip loses it from the file too. The same applies to
pivot-table formats (`PivotTableDefinitionPartReader.cs:182`).

There is no test for any of the three. `grep -rln FontCharSet XLibur.Tests` matches one file,
`XLibur.Tests/Graphics/FontTests.cs`, which is about font metrics.

**Premise, and how to disprove it.** The reflection output above is fact; the round-trip consequence
is inference from it. Task 1 is the test that settles it. If a conditional format's font name *does*
survive a round trip today, the inference is wrong, and that is the more interesting result — record
it and say why before continuing.

### Divergence 2 — the diagonal flags

`BorderToXLibur:306-316` reads `diagonalUp` / `diagonalDown` **only inside** the `<diagonal>` guard:

```csharp
var diagonalBorder = b.DiagonalBorder;
if (diagonalBorder is not null)
{
    nb = ApplyBorderStyleAndColor(nb, diagonalBorder, ...);
    if (b.DiagonalUp is not null)
        nb = nb with { DiagonalUp = b.DiagonalUp.Value };
    if (b.DiagonalDown is not null)
        nb = nb with { DiagonalDown = b.DiagonalDown.Value };
}
```

`LoadBorder:121-126` reads them unconditionally:

```csharp
LoadBorderValues(borderSource.DiagonalBorder, border.SetDiagonalBorder, border.SetDiagonalBorderColor);

if (borderSource.DiagonalUp != null)
    border.DiagonalUp = borderSource.DiagonalUp.Value;
if (borderSource.DiagonalDown != null)
    border.DiagonalDown = borderSource.DiagonalDown.Value;
```

In ECMA-376 `CT_Border`, `diagonalUp` and `diagonalDown` are **attributes of the `<border>` element**,
siblings of the `<diagonal>` child rather than dependents of it. On that reading `LoadBorder` is right
and `BorderToXLibur` is the bug: a `<border diagonalUp="1"/>` with no `<diagonal>` child decodes to a
key that differs from the same border decoded through the other path, so the two interned keys are not
equal and the writer emits a duplicate `<border>`. **Task 3 owns this decision** — read the schema,
pick the correct behaviour, and implement only that one. Do not preserve both.

`BorderToXLibur:342` also ends with `nb.Normalize()`; `LoadBorder` does not normalize. The comment at
`:338-341` explains what normalizing is for — a file may state a colour for an edge with no style, and
an un-normalized key would compare unequal to the interned form of the same border. `XLBorderKey.Normalize`
(`XLBorderKey.cs:164-187`) leaves `DiagonalUp`/`DiagonalDown` alone, so it does not paper over
divergence 2. Dxf borders currently skip normalization entirely.

### Divergence 3 — the three dxf callers read different subsets

`CT_Dxf` permits `font`, `numFmt`, `fill`, `alignment`, `border`, `protection`. What each caller reads:

| Caller | font | fill | border | numFmt | alignment | protection |
|---|:-:|:-:|:-:|:-:|:-:|:-:|
| `ConditionalFormatReader.cs:63-66` | ✓ | ✓ | ✓ | ✓ | — | — |
| `PivotTableDefinitionPartReader.cs:182-186` | ✓ | ✓ | ✓ | ✓ | ✓ | — |
| `WorkbookStylesPartWriter.cs:497-500` | ✓ | ✓ | ✓ | ✓ | — | — |

`<alignment>` is read by one of three. `<protection>` by none — there is no mutating protection
decoder at all, only the key-form `ProtectionToXLibur:410`.

The third row is a **writer** calling the load-side decoder, and that is not an accident.
`FillDifferentialFormatsCollection:488-505` decodes every dxf already present in the stylesheet — a
template's, or a previous save's — to build `Dictionary<XLStyleValue, int>`, so that
`AddConditionalFormatDxfs:355-364`, `AddPivotTableFormatDxfs:407-417` and their four siblings can call
`ContainsKey` and reuse an existing dxf index instead of appending a duplicate.

That makes the dedup key a *decode* of what the writer itself *encoded*. The two must round-trip
exactly or the map misses. They do not: the pivot reader reads `<alignment>` into
`XLPivotFormat.DxfStyleValue` (`:186`, `:194`) and the writer's dedup decode does not (`:497-500`), so
an alignment-bearing pivot dxf never matches its own entry and a duplicate is appended **on every
save**. `<dxfs count>` grows by one per round trip.

**Premise, and how to disprove it.** The two subsets differ; that is read from the code. That the
mismatch produces unbounded dxf growth is inference. Task 1 step 2 measures it directly by
round-tripping three times and counting `<dxf>` elements. A flat count disproves it and the
alignment row becomes cosmetic rather than a leak.

`FontsAreEqual:1137-1143` is a fourth writer call into load-side code, this time the *key* form
(`OpenXmlHelper.FontToXLibur`). So the writer already depends on both families being correct, and
gets one of them from each.

### Three implementations of "numFmtId → number format"

| # | Site | Input | Lookup | Custom hit | Miss / null |
|---|---|---|---|---|---|
| 1 | `WorksheetSheetDataReader.LoadStyleNumberFormat:1299-1322` (cellXfs) | `cellFormat.NumberFormatId` | **linear scan** of `<numFmts>` children per style, `FirstOrDefault` on id equality | `XLNumberFormatKey.ForFormat(code)` → id `-1` | falls through to `key with { NumberFormatId = id }` |
| 2 | `OpenXmlHelper.LoadNumberFormat:81-89` (dxf) | the dxf's own inline `<numFmt>` | **none** — the element is already in hand | `id < 164` sets the id and **discards the format code**; otherwise `FormatCode != null` sets the format | `nfSource == null` → the style keeps whatever it had |
| 3 | `LoadContext.GetNumberFormat:72-95` (pivot) | `int?` | **dictionary** `_numberFormats`, prefilled by `LoadNumberFormats:56-70` | `{ NumberFormatId = -1, Format = code }` | `{ NumberFormatId = id, Format = "" }`; `null` in → `null` out |

Sites 1 and 3 agree on the *representation* of a custom format — `XLNumberFormatKey.ForFormat`
(`XLNumberFormatKey.cs:21-30`) sets `NumberFormatId = CustomFormatNumberId` = `-1`, which is what site
3 writes by hand. That agreement is what makes unification safe. Everything else differs:

- **Complexity.** Site 1 scans the `<numFmts>` element once per distinct cell style; site 3 does a
  dictionary lookup. Same question, different complexity class, in the same load.
- **What gets indexed.** Site 3's dictionary only admits entries with both an id and a non-empty
  format code (`LoadContext.cs:63-68`), silently skipping the rest. Site 1 scans the live element and
  accepts a `<numFmt>` with no format code, falling through to the id branch.
- **Built-in versus custom.** Site 2 uses `XLConstants.NumberOfBuiltInStyles` (164,
  `XLConstants.cs:9`) as an `else if` discriminator, so `<numFmt numFmtId="5" formatCode="0.00"/>`
  in a dxf takes the id branch and **throws the format code away**. Sites 1 and 3 prefer the code
  whenever one is present. The id-versus-code exclusivity is not arbitrary — the `IXLNumberFormat`
  setters enforce it (`XLNumberFormat.cs:62-80`: assigning `NumberFormatId` resets `Format`) — but
  site 2 resolves the tie the opposite way from the other two.
- **A missing id.** Sites 1 and 3 both fall back to "it must be a built-in id", even for an id above
  164 that names a custom format the file forgot to declare. Site 2 can leave the number format
  entirely untouched.

### Style decoding inside the sheet-data reader

`XLibur/Excel/IO/WorksheetSheetDataReader.cs` is 1,460 lines. Its own summary admits the overreach
(`:19-21`):

```csharp
/// <summary>
/// Reads cell, row, and column data from a worksheet part, including style application and formula handling.
/// </summary>
```

Ten members have nothing to do with sheet data. Seven are style:

- `ApplyStyle:878` — `internal`, called from `XLWorkbook_Load.cs:306` and from `:868`, `:1097`, `:1264`
- `ResolveStyleValue:898` — `internal`, called from `StyleValueCache.cs:45` and `:65`
- `LoadStyle:905` — `internal`, called from `XLWorkbook_Load.cs:150`
- `UInt32HasValue:1054` — `internal`, used only by `LoadStyle` (`:930`, `:940`, `:943`, `:946`)
- `LoadStyleFill:1272`, `LoadStyleBorder:1283`, `LoadStyleFont:1291`, `LoadStyleNumberFormat:1299` — `private`

Two more are column structure, not sheet data:

- `LoadColumns:856` — `internal`, called from `XLWorkbook_Load.cs:438`
- `LoadColumn:1238` — `private`

(The survey this spec was drafted from said "five style entry points" and then listed seven. Seven is
correct; `UInt32HasValue` and the two column methods bring the total to ten.)

## Non-goals

- **Not touching `XLibur/Excel/Style/XLDeferred*.cs` or `XLStyle.cs`.** That is spec 23's territory.
  Same shape of finding — two implementations of one thing that must agree by hand — one storey down.
- **Not changing the style repository or the transition cache.** `XLStyleValue.FromKey`,
  `StyleValueCache` and the interning path stay exactly as they are. This spec changes what produces
  a key, never what is done with it.
- **Not renaming any field on `XL*Key.cs`.** Spec 20 owns that layout. If a unification appears to
  need a field renamed, stop and report — see Conflicts.
- **Not a performance spec.** Removing a duplicate decode must not make load *slower*; task 7
  confirms that and nothing more.
- **No public API change.** Every type named here is `internal`.

## Current state

Verified against the tree at `1b41cadd` (2026-08-24).

- `XLibur/Utils/OpenXmlHelper.cs` — 513 lines
  - mutating family: `LoadNumberFormat:81`, `LoadAlignment:91`, `LoadBorder:117`, `LoadFill:148`,
    `LoadFont:202`; helpers `LoadBorderValues:134`, `LoadSolidFill:172`, `LoadPatternedFill:191`,
    `LoadFontFamilyNumbering:224`, `LoadFontName:232`, `LoadFontSize:244`, `LoadFontUnderline:251`,
    `LoadFontVerticalAlignment:258`, `LoadFontScheme:265`
  - key family: `AlignmentToXLibur:284`, `BorderToXLibur:302`, `FontToXLibur:358`,
    `ProtectionToXLibur:410`; helper `ApplyBorderStyleAndColor:345`
  - colour conversion (`FromXLiburColor:26`, `:43`, `ToXLiburColor:66`, `:76`,
    `ConvertToXLiburColor:432`, `FillFromXLiburColor:468`), `GetBoolean:272`,
    `GetBooleanValue:51`, `GetBooleanValueAsBool:56`, `GetXLiburTextRotation:498` — **all stay**
- `XLibur/Excel/IO/WorksheetSheetDataReader.cs` — 1,460 lines; the ten members listed above
- `XLibur/Excel/IO/LoadContext.cs` — `StylesheetData` record `:12-18`; `LoadNumberFormats:56`;
  `GetNumberFormat:72`
- `XLibur/Excel/IO/ConditionalFormatReader.cs` — `LoadConditionalFormatStyle:59-67`
- `XLibur/Excel/IO/PivotTableDefinitionPartReader.cs` — `LoadFormats:170-198`
- `XLibur/Excel/IO/WorkbookStylesPartWriter.cs` — `FillDifferentialFormatsCollection:488-505`,
  `FontsAreEqual:1137`, `AppendFontScalarElements:1115-1134`
- `XLibur/Excel/IO/StyleValueCache.cs` — 69 lines, the only consumer of `ResolveStyleValue`
- Rich-text and phonetics uses of `LoadFont`: `WorksheetSheetDataReader.cs:819` (a `RunProperties`),
  `XLWorkbook_Load.cs:694` (a `RunProperties`), `WorksheetSheetDataReader.cs:1222` (a
  `PhoneticProperties`)
- Apply-a-key mechanism, already present: `WorksheetSheetDataReader.ApplyStyle:878-892` does
  `xlStylized.InnerStyle = new XLStyle(xlStylized, xlStyleKey)`

`PhoneticProperties` deserves a note. Per the SDK it has no child elements at all — `FontId`, `Type`
and `Alignment` are attributes. So `LoadFont(pp, ...)` at `:1222` finds nothing to read and merely
writes `false` into `Bold`, `Italic`, `Shadow` and `Strikethrough` on the phonetics font. It is a
no-op with a side effect. **Premise:** deleting that call changes nothing observable. Task 6 step 3 is
where that gets tested; if the suite goes red, the reset was load-bearing and must be written out
explicitly rather than left as a decoder artefact.

## File structure

```
XLibur/Excel/IO/StyleDecoder.cs                       new  — the single decoder
XLibur.Tests/Excel/IO/StyleDecoderTests.cs            new  — characterization + unit
XLibur/Utils/OpenXmlHelper.cs                         modified — Load*/  *ToXLibur families removed;
                                                                 colour + boolean helpers stay
XLibur/Excel/IO/WorksheetSheetDataReader.cs           modified — ten members lifted out
XLibur/Excel/IO/LoadContext.cs                        modified — StylesheetData gains the numFmt map;
                                                                 GetNumberFormat delegates
XLibur/Excel/IO/ConditionalFormatReader.cs            modified — four calls become one
XLibur/Excel/IO/PivotTableDefinitionPartReader.cs     modified — five calls become one
XLibur/Excel/IO/WorkbookStylesPartWriter.cs           modified — dedup decode becomes one call
XLibur/Excel/XLWorkbook_Load.cs                       modified — three call sites re-pointed
```

Nothing is deleted as a file. `OpenXmlHelper` keeps its colour layer, which is a genuinely shared
converter used by both load and save.

## The design

One module, one direction, one shape:

```csharp
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// The single decoder from OOXML style XML to XLibur style keys.
/// </summary>
/// <remarks>
/// Before spec 28 the same XML was decoded by two families chosen by provenance: a mutating one for
/// <c>&lt;dxfs&gt;</c> that wrote through <c>IXLFontBase</c> and friends, and a key-returning one for
/// <c>&lt;cellXfs&gt;</c>. They had diverged — a dxf font lost its name, family and charset, and the
/// diagonal border flags were read under different conditions. One implementation cannot diverge
/// from itself.
/// </remarks>
internal static class StyleDecoder
{
    /// <summary>Decodes one <c>&lt;xf&gt;</c> from <c>&lt;cellXfs&gt;</c>.</summary>
    internal static XLStyleKey Decode(CellFormat cellFormat, StylesheetData styles, XLStyleKey defaults);

    /// <summary>
    /// Decodes one <c>&lt;dxf&gt;</c>. Differential formats state only what they override, so every
    /// absent child leaves the corresponding part of <paramref name="defaults"/> in place.
    /// </summary>
    internal static XLStyleKey Decode(DifferentialFormat dxf, XLStyleKey defaults);

    internal static XLAlignmentKey AlignmentKey(Alignment alignment, XLAlignmentKey defaults);
    internal static XLBorderKey BorderKey(Border border, XLBorderKey defaults);
    internal static XLFillKey FillKey(Fill fill, bool differential, XLFillKey defaults);
    internal static XLFontKey FontKey(Font font, XLFontKey defaults);
    internal static XLProtectionKey ProtectionKey(Protection protection, XLProtectionKey defaults);

    /// <summary>Resolves a <c>numFmtId</c> against the workbook's declared custom formats.</summary>
    internal static XLNumberFormatKey NumberFormatKey(int numberFormatId, StylesheetData styles,
        XLNumberFormatKey defaults);

    /// <summary>Reads a <c>&lt;numFmt&gt;</c> stated inline, as a dxf states it.</summary>
    internal static XLNumberFormatKey NumberFormatKey(NumberingFormat inline, XLNumberFormatKey defaults);

    /// <summary>
    /// Decodes a rich-text run's <c>&lt;x:rPr&gt;</c>. Separate from <see cref="FontKey"/> on
    /// purpose: <c>CT_RPrElt</c> and <c>CT_Font</c> spell three children with different CLR types
    /// (<c>rFont</c>/<c>name</c>, and two each for <c>family</c> and <c>charset</c>), so one
    /// element-typed decoder cannot serve both. Conflating them is what dropped three fields from
    /// every dxf font before spec 28.
    /// </summary>
    internal static XLFontKey RunFontKey(RunProperties runProperties, XLFontKey defaults);
}
```

Two decisions carry the design.

**1. The dxf path gets its `IXL*` mutation by applying the decoded key, not by decoding twice.**
The mechanism already exists — `ApplyStyle:878-892` does exactly this for cell styles:

```csharp
xlStylized.InnerStyle = new XLStyle(xlStylized, decodedKey);
```

`XLConditionalFormat` and `XLPivotFormat`'s style container are both `XLStylizedBase`, so the same
line serves them. The three callers collapse from four or five decoder calls to one `Decode` plus one
assignment, and gain the aspects they were each missing.

**2. `RunFontKey` stays separate, and its name says why.** The rich-text provenance is genuinely a
different schema type. Keeping it distinct and typed is the fix for divergence 1, not a compromise
against it: the reason the charset was ever dropped is that `OpenXmlElement` let one function pretend
to serve two schemas.

The number-format map moves onto `StylesheetData`:

```csharp
internal sealed record StylesheetData(
    Stylesheet? Stylesheet,
    NumberingFormats? NumberingFormats,
    Fills? Fills,
    Borders? Borders,
    Fonts? Fonts,
    Dictionary<int, DifferentialFormat> DifferentialFormats)
{
    /// <summary>
    /// <c>numFmtId</c> → format code for every custom format the workbook declares. Built once, so
    /// resolving a style's number format is a dictionary hit rather than a scan of
    /// <see cref="NumberingFormats"/> per style.
    /// </summary>
    internal Dictionary<int, string> CustomNumberFormats { get; init; } = [];
}
```

`LoadContext.GetNumberFormat` then wraps `StyleDecoder.NumberFormatKey`, and its private
`_numberFormats` dictionary and `LoadNumberFormats` go away. One map, one rule, one complexity class.

## Global constraints

- Warnings are errors (`TreatWarningsAsErrors=true`); nullable is enabled — every new member must be
  null-annotated.
- Branch per spec; never commit to main. Commit prefixes: `test:` for task 1, `refactor:` for tasks
  2, 5, 6, `fix:` for tasks 3 and 4, `docs:` for task 7's Results.
- No compound shell commands (`&&`, `||`, `;`) in agent tool calls.
- **Do not use `sed -i` on tracked files.** `.gitattributes` checks out CRLF and Git Bash's `sed -i`
  rewrites the file as LF, turning a one-line change into a whole-file diff. Use the Edit/Write
  tools; verify with `git diff --numstat` — a file whose changed-line count approaches its total line
  count was rewritten, not edited.
- Build: `dotnet build XLibur/XLibur.csproj -c Release -v q`
- Tests: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
- Filtering uses `--treenode-filter`, never `--filter`. Exit 5 = invalid option; exit 8 = zero tests
  matched. Never filter at solution level — name the `.csproj`.
- Pass `-f net10.0` while iterating; run without it before opening the PR.
- Tests are TUnit and assertions are awaitable: `await Assert.That(actual).IsEqualTo(expected)`. A
  missing `await` silently passes. `[Test]`, `[Arguments(...)]`, `[MethodDataSource(...)]`. The suite
  is serial.

## Work plan

Sequential, one owner. Tasks 3 and 4 both edit `StyleDecoder`, and task 4 is what turns task 1 green.

| # | Task | Size | Gate |
|---|------|------|------|
| 1 | Characterization tests for all three divergences — **lands red on purpose** | S | Suite compiles; the new tests fail with the recorded messages |
| 2 | Build `StyleDecoder` with the seven key functions, inert | M | Build green; unit tests agree with both existing families where they agree |
| 3 | Route the `<cellXfs>` path through it; settle the diagonal question | M | Full suite green; the diagonal decision recorded |
| 4 | Route the `<dxfs>` path through it by applying the key | M | Task 1's tests go green |
| 5 | Collapse the three `numFmtId` lookups into one | S–M | Full suite green |
| 6 | Lift the ten members out of `WorksheetSheetDataReader` | S | Grep gates return nothing; suite green |
| 7 | Confirm load is not slower | S | Within benchmark noise of the merge base |

---

### Task 1 — Prove the divergences, red

This spec's central claim is that two decoders for the same XML have already diverged. That claim is
worth nothing until a test fails. **These tests land failing, on a commit whose message says so**, and
task 4 turns them green.

**Files:**
- Create: `XLibur.Tests/Excel/IO/StyleDecoderTests.cs`

**Interfaces:**
- Produces: `A_conditional_format_font_keeps_its_name_family_and_charset`,
  `Round_tripping_does_not_grow_the_dxf_table`, `The_diagonal_flags_decode_the_same_from_both_paths`.

- [ ] **Step 1: Write the font round-trip test**

```csharp
using System.IO;
using System.Threading.Tasks;
using XLibur.Excel;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// Spec 28: the same style XML is decoded by two implementations chosen by which element it came
/// from, and they have diverged. These tests pin the divergences. They fail on the tree at
/// 1b41cadd and are made to pass by spec 28 task 4.
/// </summary>
public class StyleDecoderTests
{
    /// <summary>
    /// OpenXmlHelper.LoadFont takes an untyped OpenXmlElement and looks for RunFont, FontFamily and
    /// no charset at all — the &lt;x:rPr&gt; spellings. A dxf hands it a &lt;x:font&gt;, whose
    /// corresponding children are the unrelated types FontName, FontFamilyNumbering and FontCharSet.
    /// The writer emits all three (WorkbookStylesPartWriter.AppendFontScalarElements), so they reach
    /// the file and are dropped on the way back.
    /// </summary>
    [Test]
    public async Task A_conditional_format_font_keeps_its_name_family_and_charset()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            var cf = ws.Range("A1:A5").AddConditionalFormat().WhenGreaterThan(5);
            cf.Font.FontName = "Arial";
            cf.Font.FontCharSet = XLFontCharSet.Arabic;
            cf.Font.FontFamilyNumbering = XLFontFamilyNumberingValues.Swiss;
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using var reloaded = new XLWorkbook(ms);
        var format = reloaded.Worksheet("Sheet1").ConditionalFormats.Single();

        await Assert.That(format.Style.Font.FontName).IsEqualTo("Arial");
        await Assert.That(format.Style.Font.FontCharSet).IsEqualTo(XLFontCharSet.Arabic);
        await Assert.That(format.Style.Font.FontFamilyNumbering)
            .IsEqualTo(XLFontFamilyNumberingValues.Swiss);
    }
}
```

- [ ] **Step 2: Write the dxf-growth test**

```csharp
    /// <summary>
    /// WorkbookStylesPartWriter.FillDifferentialFormatsCollection decodes the dxfs already in the
    /// stylesheet to build the reuse map, but reads a different set of children than the pivot
    /// reader does. A dxf carrying &lt;alignment&gt; therefore never matches its own map entry and a
    /// duplicate is appended on every save.
    /// </summary>
    [Test]
    public async Task Round_tripping_does_not_grow_the_dxf_table()
    {
        var bytes = BuildWorkbookWithAnAlignedConditionalFormat();
        var counts = new List<int> { CountDxfs(bytes) };

        for (var i = 0; i < 3; i++)
        {
            bytes = ReSave(bytes);
            counts.Add(CountDxfs(bytes));
        }

        await Assert.That(counts.Distinct().Count())
            .IsEqualTo(1)
            .Because($"dxf count per round trip: {string.Join(", ", counts)}");
    }
```

`CountDxfs` opens the package with `SpreadsheetDocument`, reads
`WorkbookStylesPart.Stylesheet.DifferentialFormats`, and returns
`ChildElements.Count`. `ReSave` opens the bytes with `XLWorkbook` and saves to a fresh
`MemoryStream` without touching anything.

Build the fixture with a conditional format **and** a pivot-table format, since the alignment
mismatch is on the pivot path. If constructing a pivot format from the public API is more than a few
lines, reuse the pivot fixture from `XLibur.Tests/Excel/PivotTables/`.

- [ ] **Step 3: Write the diagonal test**

```csharp
    /// <summary>
    /// A &lt;border diagonalUp="1"/&gt; with no &lt;diagonal&gt; child decodes one way through
    /// BorderToXLibur (flags read only inside the &lt;diagonal&gt; guard) and another through
    /// LoadBorder (flags read unconditionally). One of the two is wrong per ECMA-376 CT_Border;
    /// spec 28 task 3 decides which.
    /// </summary>
    [Test]
    [Arguments(true, false)]
    [Arguments(false, true)]
    [Arguments(true, true)]
    public async Task The_diagonal_flags_decode_the_same_from_both_paths(bool up, bool down)
    {
        var border = new DocumentFormat.OpenXml.Spreadsheet.Border
        {
            DiagonalUp = up,
            DiagonalDown = down,
        };

        var throughKeyPath = OpenXmlHelper.BorderToXLibur(border, XLBorderValue.Default.Key);

        var mutated = new XLStyle(null);
        OpenXmlHelper.LoadBorder(border, mutated.Border);
        var throughMutatingPath = ((XLBorder)mutated.Border).Key.Normalize();

        await Assert.That(throughKeyPath).IsEqualTo(throughMutatingPath);
    }
```

`XLBorder.Key` is `internal`; the test project already has `InternalsVisibleTo`. If reaching the key
off an `IXLBorder` is awkward, compare the four observable properties instead — the assertion that
matters is that the two paths agree.

- [ ] **Step 4: Run them and record exactly how they fail**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/StyleDecoderTests/*"`

Expected: **FAIL.** Record the actual values in a Results section of this spec:
- what `FontName` / `FontCharSet` / `FontFamilyNumbering` came back as;
- the dxf count sequence across four saves;
- which diagonal argument rows differ, and by which field.

**If any of the three passes, that premise is disproved.** Say so in the Results section and adjust
the spec before continuing. A disproved premise is a real result here, not a setback — the divergence
argument survives on whichever of the three still fails, and the design does not change.

- [ ] **Step 5: Commit, red on purpose**

```bash
git add XLibur.Tests/Excel/IO/StyleDecoderTests.cs
git commit -m 'test(io): pin three style-decoder divergences - these FAIL on purpose until spec 28 task 4'
```

The message must say the tests fail, so a bisect or a CI run over this commit is not read as a
regression.

---

### Task 2 — Build `StyleDecoder`, inert

Nothing calls it yet. This task is only about getting the seven key functions right in one place.

**Files:**
- Create: `XLibur/Excel/IO/StyleDecoder.cs`
- Modify: `XLibur.Tests/Excel/IO/StyleDecoderTests.cs`

**Interfaces:**
- Produces: `StyleDecoder.Decode(CellFormat, StylesheetData, XLStyleKey)`,
  `Decode(DifferentialFormat, XLStyleKey)`, `AlignmentKey`, `BorderKey`, `FillKey`, `FontKey`,
  `ProtectionKey`, two `NumberFormatKey` overloads, `RunFontKey`.

- [ ] **Step 1: Move the four key functions across verbatim**

`AlignmentToXLibur:284`, `BorderToXLibur:302`, `FontToXLibur:358` and `ProtectionToXLibur:410` move
from `OpenXmlHelper` into `StyleDecoder` as `AlignmentKey`, `BorderKey`, `FontKey` and
`ProtectionKey`, with their bodies unchanged and `ApplyBorderStyleAndColor:345` alongside them. Leave
`OpenXmlHelper` forwarding stubs in place for this task only — task 3 deletes them.

Do not "fix" the diagonal guard here. That is task 3, and it needs the schema decision first.

- [ ] **Step 2: Write `FillKey`, the first genuinely new one**

Port `LoadFill:148-200` — including `LoadSolidFill:172` and `LoadPatternedFill:191` — from mutation
into a key:

```csharp
    /// <summary>
    /// Decodes a <c>&lt;fill&gt;</c>.
    /// </summary>
    /// <param name="differential">
    /// Differential fills store background in <c>bgColor</c> and pattern in <c>fgColor</c>, which is
    /// the sane reading. Ordinary fills store the background in <c>fgColor</c> when the pattern is
    /// solid. The flag selects between them; it is not a style choice.
    /// </param>
    internal static XLFillKey FillKey(Fill fill, bool differential, XLFillKey defaults)
    {
        if (fill.PatternFill is not { } patternFill)
            return defaults;

        var patternType = patternFill.PatternType is not null
            ? patternFill.PatternType.Value.ToXLibur()
            : XLFillPatternValues.Solid;

        var key = defaults with { PatternType = patternType };

        return patternType switch
        {
            XLFillPatternValues.None => key,
            XLFillPatternValues.Solid => key with
            {
                BackgroundColor = SolidFillBackground(patternFill, differential),
            },
            _ => key with
            {
                PatternColor = patternFill.ForegroundColor is not null
                    ? patternFill.ForegroundColor.ToXLiburColor().Key
                    : key.PatternColor,
                BackgroundColor = patternFill.BackgroundColor is not null
                    ? patternFill.BackgroundColor.ToXLiburColor().Key
                    : XLColor.FromIndex(64).Key,
            },
        };
    }
```

Check the exact field names on `XLFillKey` before writing this — `PatternType`, `PatternColor` and
`BackgroundColor` are read from `XLFill.Key` usage, not confirmed against `XLFillKey.cs`. If a name
differs, use the real one; **do not rename it** (spec 20's territory).

Note the asymmetry `LoadFill` has today and preserve it: the `None` branch touches no colour, while
both other branches default a missing background to `XLColor.FromIndex(64)`.

- [ ] **Step 3: Write the two `NumberFormatKey` overloads**

The inline overload reproduces `LoadNumberFormat:81-89` exactly, including the `164` threshold and the
`else if`:

```csharp
    internal static XLNumberFormatKey NumberFormatKey(NumberingFormat inline, XLNumberFormatKey defaults)
    {
        if (inline.NumberFormatId is { Value: var id } && id < XLConstants.NumberOfBuiltInStyles)
            return new XLNumberFormatKey { NumberFormatId = (int)id, Format = string.Empty };

        if (inline.FormatCode?.Value is { Length: > 0 } code)
            return XLNumberFormatKey.ForFormat(code);

        return defaults;
    }
```

The id branch clearing `Format` is not a change: `IXLNumberFormat.NumberFormatId`'s setter
(`XLNumberFormat.cs:62-80`) resets `Format` to the default, so the mutating path already produced
this key. Confirm that by reading the setter before trusting the sentence.

The id overload reproduces `LoadStyleNumberFormat:1299-1322` but over the map instead of a scan —
see task 5, which is where the map arrives. Until then, take a `NumberingFormats?` and keep the scan.

- [ ] **Step 4: Write `RunFontKey`**

Port `LoadFont:202-222` and its five private helpers, but typed to `RunProperties` and using that
schema's names: `RunFont`, `FontFamily`, `RunPropertyCharSet`. **Add the charset** — the rich-text
path drops it today for exactly the same reason the dxf path drops three fields, and one decoder that
reads the schema it is given is the whole point of this spec.

Record in the commit message that `RunFontKey` reads one field more than `LoadFont` did.

- [ ] **Step 5: Write the two `Decode` composites**

`Decode(CellFormat, …)` follows `LoadStyle:905-948` branch for branch, including the
`ApplyProtection != null` guard and the `IncludeQuotePrefix` line. `Decode(DifferentialFormat, …)`
reads **all six** children the schema permits — font, fill, border, numFmt, alignment, protection —
which is the union of what the three current callers read plus protection, which none read.

- [ ] **Step 6: Unit-test the decoder against both old families where they agree**

Add tests to `StyleDecoderTests` that feed the same constructed OOXML element to `StyleDecoder` and
to the surviving stub, and assert equality — for fill, alignment, protection and the nine font fields
both old paths already handled. These must pass immediately. They are the proof that task 2 moved
code without changing it.

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/StyleDecoderTests/*"`
Expected: the task 2 unit tests PASS; task 1's three still FAIL.

- [ ] **Step 7: Commit**

```bash
git add XLibur/Excel/IO/StyleDecoder.cs XLibur/Utils/OpenXmlHelper.cs XLibur.Tests/Excel/IO/StyleDecoderTests.cs
git commit -m 'refactor(io): add StyleDecoder as the single OOXML style decoder (spec 28 task 2)'
```

---

### Task 3 — Route `<cellXfs>` through it, and settle the diagonal question

**Files:**
- Modify: `XLibur/Excel/IO/WorksheetSheetDataReader.cs:905-948`, `:1272-1322`
- Modify: `XLibur/Excel/IO/StyleDecoder.cs`
- Modify: `XLibur/Utils/OpenXmlHelper.cs` — delete the forwarding stubs from task 2 step 1

- [ ] **Step 1: Decide the diagonal question from the schema**

Read ECMA-376 Part 1, `CT_Border` (§18.8.4 `border`). Establish whether `diagonalUp` and
`diagonalDown` are attributes of `border` itself — in which case they are independent of the
`<diagonal>` child and `LoadBorder:123-126` is correct — or whether the schema ties them to it.

Implement **one** behaviour in `StyleDecoder.BorderKey`. Write the verdict into the method's
`<remarks>` with the section reference, and into this spec's Results section.

The expected answer is that the attributes are independent, making `BorderToXLibur:312-315`'s guard
the bug. **Confirm it; do not assume it.** If the schema says otherwise, fix towards the schema and
say so — the point of the task is a single correct rule, not a particular one.

- [ ] **Step 2: Collapse `LoadStyle` onto `Decode`**

```csharp
    internal static void LoadStyle(ref XLStyleKey xlStyle, int styleIndex, StylesheetData styles)
    {
        if (styles.Stylesheet is not { CellFormats: not null } s)
            return; // No stylesheet, no styles.

        xlStyle = StyleDecoder.Decode((CellFormat)s.CellFormats.ElementAt(styleIndex), styles, xlStyle);
    }
```

`LoadStyleFill:1272`, `LoadStyleBorder:1283`, `LoadStyleFont:1291` and `LoadStyleNumberFormat:1299`
are deleted; their bodies now live in `StyleDecoder`. `UInt32HasValue:1054` moves with them and
becomes `private` on `StyleDecoder` — grep first to confirm nothing outside `LoadStyle` uses it:

Run: `grep -rn 'UInt32HasValue' XLibur --include=*.cs`
Expected: only `WorksheetSheetDataReader.cs` (4 call sites plus the declaration) before the move.

- [ ] **Step 3: Delete the forwarding stubs and the mutating border/alignment/protection decoders**

`OpenXmlHelper.LoadBorder:117`, `LoadBorderValues:134` and `LoadAlignment:91` still have dxf callers
until task 4. Leave those three. Delete the four `*ToXLibur` stubs added in task 2 step 1 and
re-point their callers:

Run: `grep -rn 'OpenXmlHelper\.\(AlignmentToXLibur\|BorderToXLibur\|FontToXLibur\|ProtectionToXLibur\)' XLibur --include=*.cs`
Expected before: `WorkbookStylesPartWriter.cs:1139` (`FontsAreEqual`) and the sheet-data reader.
Expected after: nothing.

- [ ] **Step 4: Build and run the full suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`

Expected: PASS, except task 1's three, which still fail — the dxf path has not moved yet. The
diagonal one may go green here if the decision matched the key path's current behaviour; note which.

**If a cell-style test fails, the port is not verbatim.** Diff `Decode(CellFormat, …)` against
`LoadStyle:905-948` line by line before changing anything else. The four guards
(`ApplyProtection != null`, and the three `UInt32HasValue` calls) each suppress a decode, and dropping
one changes which defaults survive.

- [ ] **Step 5: Verify the diff is an edit, not a rewrite**

Run: `git diff --numstat`
Expected: `WorksheetSheetDataReader.cs` loses roughly 70 lines and gains under 10. A changed-line
count near the file's 1,460 total means the line endings were rewritten — see the constraints.

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel/IO/StyleDecoder.cs XLibur/Excel/IO/WorksheetSheetDataReader.cs XLibur/Utils/OpenXmlHelper.cs XLibur/Excel/IO/WorkbookStylesPartWriter.cs
git commit -m 'fix(io): decode cellXfs through StyleDecoder; one rule for the diagonal flags (spec 28 task 3)'
```

---

### Task 4 — Route `<dxfs>` through it by applying the key

This is the task that closes the defect.

**Files:**
- Modify: `XLibur/Excel/IO/ConditionalFormatReader.cs:59-67`
- Modify: `XLibur/Excel/IO/PivotTableDefinitionPartReader.cs:170-198`
- Modify: `XLibur/Excel/IO/WorkbookStylesPartWriter.cs:488-505`
- Modify: `XLibur/Utils/OpenXmlHelper.cs` — delete `LoadNumberFormat`, `LoadAlignment`, `LoadBorder`,
  `LoadFill`, `LoadFont` and their nine helpers

- [ ] **Step 1: Conditional formats**

```csharp
    private static void LoadConditionalFormatStyle(ConditionalFormattingRule fr,
        XLConditionalFormat conditionalFormat, Dictionary<int, DifferentialFormat> differentialFormats)
    {
        var dxf = differentialFormats[(int)fr.FormatId!.Value];

        // One decode of the whole dxf, then apply the key - rather than four per-aspect decodes
        // through a second implementation. This is what makes a conditional format's font name,
        // family, charset and alignment survive a load; see spec 28.
        var key = StyleDecoder.Decode(dxf, conditionalFormat.StyleValue.Key);
        conditionalFormat.InnerStyle = new XLStyle(conditionalFormat, key);
    }
```

Check the member names on `XLConditionalFormat` before writing this — it derives from
`XLStylizedBase`, which is where `InnerStyle` and `StyleValue` come from
(`WorksheetSheetDataReader.ApplyStyle:878-892` is the working example).

- [ ] **Step 2: Pivot formats**

`PivotTableDefinitionPartReader.LoadFormats:179-187` becomes:

```csharp
            var dxfStyle = XLStyle.Default;
            if (format.FormatId is not null)
            {
                var df = differentialFormats[checked((int)format.FormatId.Value)];
                dxfStyle = new XLStyle(null, StyleDecoder.Decode(df, XLStyle.Default.Key));
            }
```

`XLStyle.Default` is a property returning a fresh instance each call (`XLStyle.cs:10`), so the
existing code is not mutating a shared singleton — but the replacement does not depend on that either
way.

- [ ] **Step 3: The writer's dedup decode**

```csharp
        foreach (var df in differentialFormats.Elements<DifferentialFormat>())
        {
            // The reuse map is a decode of what this writer encodes, so it must read exactly the
            // children AddStyleAsDifferentialFormat writes. Before spec 28 it read four of six and
            // the pivot reader read five, so an alignment-bearing dxf never matched its own entry
            // and a duplicate was appended on every save.
            var key = StyleDecoder.Decode(df, DefaultStyle.Key);
            var styleValue = XLStyleValue.FromKey(ref key);

            if (!dictionary.ContainsKey(styleValue))
                dictionary.Add(styleValue, id++);
        }
```

This removes the `XLStylizedEmpty` allocation per dxf as a side effect. `DefaultStyle` is already in
scope at `:496`; confirm whether it is an `IXLStyle` or an `XLStyle` and reach its key accordingly.

- [ ] **Step 4: Delete the mutating family**

Run: `grep -rn 'OpenXmlHelper\.Load' XLibur --include=*.cs`

Expected remaining: `WorksheetSheetDataReader.cs:819`, `XLWorkbook_Load.cs:694` (both
`RunProperties`) and `WorksheetSheetDataReader.cs:1222` (`PhoneticProperties`). Re-point the first two
to `StyleDecoder.RunFontKey`; the third is handled in task 6 step 3.

`RunFontKey` returns a key while the two rich-text call sites mutate an `IXLFontBase` (an
`XLRichString`). Applying a key there is not the same move as applying one to an `IXLStylized` —
`XLRichString` has no `InnerStyle`. If a clean apply is not available, keep a small
`ApplyRunFont(RunProperties, IXLFontBase)` on `StyleDecoder` that calls `RunFontKey` and writes the
fields through. **One decode, one shape** is the invariant; a thin applier over it is fine, a second
decode is not.

Then delete `LoadNumberFormat:81`, `LoadAlignment:91`, `LoadBorder:117`, `LoadFill:148`,
`LoadFont:202`, `LoadBorderValues:134`, `LoadSolidFill:172`, `LoadPatternedFill:191`,
`LoadFontFamilyNumbering:224`, `LoadFontName:232`, `LoadFontSize:244`, `LoadFontUnderline:251`,
`LoadFontVerticalAlignment:258`, `LoadFontScheme:265`.

Keep `GetBoolean:272`, `GetBooleanValue:51`, `GetBooleanValueAsBool:56`, `GetXLiburTextRotation:498`
and the whole colour layer — those are shared converters with save-side callers.

- [ ] **Step 5: Run task 1's tests — this is the gate**

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/StyleDecoderTests/*"`
Expected: **PASS**, all three.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS.

**Expect conditional-format and pivot tests to fail here, and read each failure carefully.** The dxf
path now reads two aspects it did not before (alignment for conditional formats, protection for
both), so a test asserting the *old* dropped-field behaviour will break. That is the fix working. A
test asserting a *value* that changed — a colour, a border style — is a port error. Do not weaken an
assertion without deciding which of the two it is, and record any assertion that changes.

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel/IO/ConditionalFormatReader.cs XLibur/Excel/IO/PivotTableDefinitionPartReader.cs XLibur/Excel/IO/WorkbookStylesPartWriter.cs XLibur/Excel/IO/StyleDecoder.cs XLibur/Utils/OpenXmlHelper.cs XLibur/Excel/IO/WorksheetSheetDataReader.cs XLibur/Excel/XLWorkbook_Load.cs
git commit -m 'fix(io): decode dxfs through StyleDecoder - conditional format fonts keep name, family and charset (spec 28 task 4)'
```

---

### Task 5 — One `numFmtId` lookup

**Files:**
- Modify: `XLibur/Excel/IO/LoadContext.cs:12-18`, `:56-95`
- Modify: `XLibur/Excel/IO/StyleDecoder.cs`
- Modify: `XLibur/Excel/XLWorkbook_Load.cs` — wherever `LoadNumberFormats` is called

- [ ] **Step 1: Put the map on `StylesheetData`**

Add `CustomNumberFormats` as shown in The design, built once where `StylesheetData` is constructed.
Find that site first:

Run: `grep -rn 'new StylesheetData' XLibur --include=*.cs`

Build it with `LoadContext.LoadNumberFormats:56-70`'s rule — id present, format code non-empty —
because that is the rule the pivot path already ships and the one this task keeps. Note the
difference from `LoadStyleNumberFormat`: the scan accepted a `<numFmt>` with no format code and fell
through to the id branch. With a map that skips such entries the fall-through happens anyway, by
missing. **Confirm that equivalence rather than assuming it** — it is the one behavioural claim in
this task.

- [ ] **Step 2: One resolver on `StyleDecoder`**

```csharp
    /// <summary>
    /// Resolves a <c>numFmtId</c>. A workbook-declared custom format wins; anything else is taken to
    /// be a built-in id, including an id above <see cref="XLConstants.NumberOfBuiltInStyles"/> that
    /// no <c>&lt;numFmt&gt;</c> declares — such a file is malformed and this is what both the cell
    /// and pivot paths did before spec 28 unified them.
    /// </summary>
    internal static XLNumberFormatKey NumberFormatKey(int numberFormatId, StylesheetData styles,
        XLNumberFormatKey defaults)
    {
        if (styles.CustomNumberFormats.TryGetValue(numberFormatId, out var formatCode))
            return XLNumberFormatKey.ForFormat(formatCode);

        return defaults with { NumberFormatId = numberFormatId, Format = string.Empty };
    }
```

Check `defaults with { Format = string.Empty }` against `LoadStyleNumberFormat:1319`, which writes
`xlNumberFormat with { NumberFormatId = ... }` and leaves `Format` inherited. `LoadContext:89-94`
writes `Format = string.Empty`. **They differ, and this task must pick one.** Picking
`string.Empty` matches the pivot path and matches what a built-in id means (`XLNumberFormatKey.cs:15`:
`-1` means custom, anything else means the format lives in `XLPredefinedFormat`); picking the
inherited value matches the cell path. Test both against the suite and record which one is compatible.

- [ ] **Step 3: Make `LoadContext.GetNumberFormat` a thin wrapper**

```csharp
    internal XLNumberFormatValue? GetNumberFormat(int? numberFormatId)
    {
        if (numberFormatId is not { } id)
            return null;

        var key = StyleDecoder.NumberFormatKey(id, Styles, XLNumberFormatValue.Default.Key);
        return XLNumberFormatValue.FromKey(ref key);
    }
```

Delete `_numberFormats` (`:31`) and `LoadNumberFormats` (`:56-70`), and remove the call to it.

- [ ] **Step 4: Build and run**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`
Expected: PASS.

Pay attention to the number-format and pivot suites:

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*NumberFormat*/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*Pivot*/*"`

- [ ] **Step 5: Commit**

```bash
git add XLibur/Excel/IO/LoadContext.cs XLibur/Excel/IO/StyleDecoder.cs XLibur/Excel/XLWorkbook_Load.cs
git commit -m 'refactor(io): one numFmtId lookup instead of three (spec 28 task 5)'
```

---

### Task 6 — Get style decoding out of the sheet-data reader

**Files:**
- Modify: `XLibur/Excel/IO/WorksheetSheetDataReader.cs`
- Modify: `XLibur/Excel/IO/StyleDecoder.cs`
- Create: `XLibur/Excel/IO/WorksheetColumnReader.cs`
- Modify: `XLibur/Excel/IO/StyleValueCache.cs`, `XLibur/Excel/XLWorkbook_Load.cs`

- [ ] **Step 1: Move the three style entry points onto `StyleDecoder`**

`ApplyStyle:878`, `ResolveStyleValue:898` and `LoadStyle:905` move as-is. `LoadStyle` is now three
lines (task 3 step 2) and can be inlined into the other two, leaving `XLWorkbook_Load.cs:150`'s
`ref XLStyleKey` call to be re-pointed at `StyleDecoder.Decode` directly.

Re-point the four external callers:
- `StyleValueCache.cs:45`, `:65` and its two `<see cref>` references at `:17`, `:56`
- `XLWorkbook_Load.cs:150`, `:306`

- [ ] **Step 2: Move the two column methods**

`LoadColumns:856` and `LoadColumn:1238` go to a new `WorksheetColumnReader`. `<cols>` is a worksheet
element, not sheet data; the loader already treats it as one (`XLWorkbook_Load.cs:438` dispatches on
`typeof(Columns)`).

**If spec 24 has landed**, `TryLoad` calls `WorksheetSheetDataReader.LoadColumns` from inside
`WorksheetElementReader` — re-point that call instead of the one in `XLWorkbook_Load`. Check with:

Run: `grep -rn 'LoadColumns' XLibur --include=*.cs`

- [ ] **Step 3: Delete the phonetics font decode**

`LoadPhonetics:1222` calls the font decoder with a `PhoneticProperties`, which per the SDK has no
child elements — only `FontId`, `Type` and `Alignment` attributes. The call reads nothing and writes
`false` into four booleans.

Delete the line. Run the suite.

Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*Phonetic*/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0 --treenode-filter "/*/*/*RichText*/*"`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj -f net10.0`

Expected: PASS. **If it goes red, the reset was load-bearing** — some path relies on a phonetics font
arriving with those four flags cleared. Restore the effect as an explicit four-line assignment with a
comment naming the test that demanded it, and record the finding. Do not restore it by reinstating
the decoder call.

- [ ] **Step 4: Confirm the reader is sheet data only**

Run: `grep -nE 'Style|Column' XLibur/Excel/IO/WorksheetSheetDataReader.cs`

Expected: matches only where a cell or row *uses* a resolved style —
`GetInheritedStyleFast:707`, the `StyleIndex` fields on `RowProperties`/`CellProperties`, and
`ApplyRowCustomProps:1064`. No `LoadStyle*`, no `ApplyStyle`, no `ResolveStyleValue`, no
`UInt32HasValue`, no `LoadColumn`.

- [ ] **Step 5: Build both frameworks and run the whole suite**

Run: `dotnet build XLibur/XLibur.csproj -c Release -v q`
Run: `dotnet test XLibur.Tests/XLibur.Tests.csproj`
Expected: PASS on net8.0 and net10.0.

- [ ] **Step 6: Commit**

```bash
git add XLibur/Excel/IO/StyleDecoder.cs XLibur/Excel/IO/WorksheetColumnReader.cs XLibur/Excel/IO/WorksheetSheetDataReader.cs XLibur/Excel/IO/StyleValueCache.cs XLibur/Excel/XLWorkbook_Load.cs
git commit -m 'refactor(io): lift style and column decoding out of the sheet data reader (spec 28 task 6)'
```

---

### Task 7 — Confirm load is not slower

This spec removes a duplicate decode and replaces a per-style linear scan with a dictionary hit. Both
should make load marginally faster or leave it flat. Neither is the point, and neither is claimed —
what must be shown is that it did not get **worse**.

- [ ] **Step 1: Measure the merge base**

```
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --filter '*XLiburReadBenchmarks*'
dotnet run -c Release --project XLibur.Benchmarks/XLibur.Benchmarks.csproj -- --filter '*TemplateRoundTripBenchmarks*'
```

Record `LoadWorkbook`, `LoadAndReadAllCells`, `Open` and `OpenAndSaveUnchanged`, mean and allocated.

- [ ] **Step 2: Measure the branch**

Same two commands, same machine, nothing else running.

- [ ] **Step 3: Compare**

Expected: within noise. This machine shows roughly 40% run-to-run variance on wall-clock, so **only
the allocation column is a reliable signal**, and only BenchmarkDotNet numbers count — a stopwatch in
a test proves nothing here.

Allocations should be flat or slightly down: task 4 removes one `XLStylizedEmpty` per existing dxf
and task 3 removes one `XLFill` per distinct cell fill.

**Decision rule.** An allocation *increase* on any of the four must be explained before this spec
lands. The likeliest cause is `StylesheetData.CustomNumberFormats` being rebuilt per worksheet
instead of per workbook — check where `StylesheetData` is constructed.

- [ ] **Step 4: Write the Results section and commit it**

Record: the task 1 failure values, the diagonal verdict and its schema citation, the
`string.Empty`-versus-inherited number-format decision from task 5 step 2, the phonetics outcome from
task 6 step 3, and this task's four benchmark rows. Any premise this spec named and the work
disproved goes here too, stated as a result.

```bash
git add docs/specs/28-single-style-decoder.md
git commit -m 'docs(specs): record spec 28 results - decoder divergences, diagonal verdict, load numbers'
```

---

## Acceptance criteria

1. `XLibur/Utils/OpenXmlHelper.cs` contains no style decoder. Gate:
   `grep -nE 'LoadNumberFormat|LoadAlignment|LoadBorder|LoadFill|LoadFont|ToXLibur\(' XLibur/Utils/OpenXmlHelper.cs`
   returns only the colour conversions (`ToXLiburColor`, `FromXLiburColor`).
2. `grep -rn 'OpenXmlHelper\.Load' XLibur --include=*.cs` returns **nothing**.
3. `StyleDecoder` is the only module that constructs an `XLFontKey`, `XLBorderKey`, `XLFillKey`,
   `XLAlignmentKey` or `XLProtectionKey` from an OOXML element. Gate:
   `grep -rln 'DocumentFormat.OpenXml.Spreadsheet.Font\|Elements<Bold>\|Elements<RunFont>' XLibur --include=*.cs`
   lists `StyleDecoder.cs` and the writers, not the readers.
4. One `numFmtId` resolver. Gate: `grep -rn 'NumberOfBuiltInStyles' XLibur --include=*.cs` matches
   `XLConstants.cs`, `StyleDecoder.cs` and `WorkbookStylesPartWriter.cs` only —
   `LoadContext.cs` and `WorksheetSheetDataReader.cs` no longer appear.
5. `LoadContext` has no `_numberFormats` field and no `LoadNumberFormats` method.
6. `WorksheetSheetDataReader.cs` declares none of `LoadStyle`, `ApplyStyle`, `ResolveStyleValue`,
   `LoadStyleFill`, `LoadStyleBorder`, `LoadStyleFont`, `LoadStyleNumberFormat`, `UInt32HasValue`,
   `LoadColumns`, `LoadColumn`. Gate: `grep -cE '(LoadStyle|ApplyStyle|ResolveStyleValue|UInt32HasValue|LoadColumn)' XLibur/Excel/IO/WorksheetSheetDataReader.cs`
   returns `0`. Its class summary no longer says "including style application".
7. All three of task 1's tests pass, none of them weakened.
8. `<dxfs count>` is identical after 1, 2, 3 and 4 saves of the same workbook.
9. A conditional format's `FontName`, `FontCharSet` and `FontFamilyNumbering` survive a round trip.
10. The diagonal-flag rule is implemented once, and the `<remarks>` on `StyleDecoder.BorderKey`
    cites the ECMA-376 section that decided it.
11. Full suite green on net8.0 and net10.0.
12. No public API change. Gate: `git diff --stat` touches no file whose types are `public`.
13. Load allocations on `LoadWorkbook` and `Open` are not higher than the merge base.
14. Every premise this spec labels as such is either confirmed or recorded as disproved in a Results
    section.

## Conflicts

**Spec 23 — file-disjoint, and the same finding one storey down.** 23 deletes the seven `XLDeferred*`
twins under `XLibur/Excel/Style/` and modifies the seven facades beside them; verified against its
File structure block (`23-single-style-facade.md:169-186`), which lists only
`XLibur/Excel/Style/*.cs`. It never touches `XLibur/Utils/OpenXmlHelper.cs`,
`XLibur/Excel/IO/WorksheetSheetDataReader.cs` or any file this spec modifies. Worth reading first
anyway: 23 found that every style *interface* has two implementations that must agree by hand and had
already disagreed on `InsideBorder`; 28 finds the same thing about every style *decoder*. The two can
run concurrently.

One bookkeeping note: `23-single-style-facade.md` carries `**Status:** Done — see Results` while
`docs/specs/README.md:36` still lists it as Proposed, and the file is untracked in git as of
`1b41cadd`. Treat it as landed on a branch, not on main, and do not rebase onto it.

**Spec 20 — reads its types, must not rename its fields.** 20 relayouts `XLColorKey` and the four
keys that embed it (`20-style-key-struct-size.md`). This spec reads `XLFontKey`, `XLBorderKey`,
`XLFillKey`, `XLAlignmentKey`, `XLProtectionKey`, `XLNumberFormatKey` and `XLStyleKey` through `with`
expressions, exactly as the code it replaces already does. It adds no field and renames none.

**If a unification appears to require a field rename, stop and report.** That is spec 20's territory,
and doing it here would collide with a spec whose whole subject is those declarations. The two are
otherwise sequential in either order — 28 first is marginally better, since it leaves 20 one decoder
to re-verify against instead of two.

Note also that spec 20's measured sizes are stale: it records `XLBorderKey` at 264 bytes, while
`XLBorderKey.cs:55-73` documents a field-grouping change that brought it to 88. Do not cite 20's
table as current.

**Spec 02 — done, and adjacent.** 02 rewrote the load path for allocations and owns the shape of
`WorksheetSheetDataReader`'s sheet-data half — `SheetDataReadContext`, the `XmlReader` buffers,
`GetInheritedStyleFast:707`. This spec removes the *style* half and touches none of that. `LoadStyle`
and `ResolveStyleValue` are what 02's `StyleValueCache` calls, so the cache's two call sites move;
its caching behaviour does not.

**Spec 24 — one call site, no conflict.** 24 moves the worksheet element dispatch into
`WorksheetElementReader`, and its `TryLoad` calls `WorksheetSheetDataReader.LoadColumns`. Task 6 step
2 moves that method. Whichever lands second re-points one call. Task 6 step 2 says how to check.

**Spec 18 task 5** owns per-sheet load cost. This spec's task 7 confirms it did not move; it does not
try to improve it.

No other spec in `docs/specs/` touches `XLibur/Utils/OpenXmlHelper.cs`,
`XLibur/Excel/IO/ConditionalFormatReader.cs`, `XLibur/Excel/IO/LoadContext.cs` or
`XLibur/Excel/IO/PivotTableDefinitionPartReader.cs`.

---

## Results

Implemented on `task/28`, base `c569b95a`, tip `c89d3e04`. Eight commits — one per task, plus one
unplanned fix found in review. Suite green on both TFMs: 28,292 tests, 0 failed, 10 skipped.

### The headline defect is real and is fixed

`A_conditional_format_font_keeps_its_name_family_and_charset` failed on the tree at `c569b95a` with
`FontName` coming back `"Calibri"` instead of `"Arial"`. The inference in *Divergence 1* was
correct: `LoadFont` searched the `<x:rPr>` spellings while all three dxf callers passed a `<x:font>`,
and the three fields died on the way back in. Green since task 4.

`FontCharSet` and `FontFamilyNumbering` were not separately observed failing, because the test
asserts `FontName` first and TUnit stops at the first failed assertion. All three assert green now.

### The diagonal question, settled

**ECMA-376 Part 1 §18.8.4 (`border`, `CT_Border`)** declares `diagonalUp`, `diagonalDown` and
`outline` as *attributes of the `border` element*, while `diagonal` is one of nine `CT_BorderPr`
child elements in the type's sequence. An attribute of `border` is a sibling of the `diagonal`
child, not a dependent of it.

So the unconditional read — the old `LoadBorder` — is correct, and `BorderToXLibur`'s
`<diagonal>` guard was the bug. `StyleDecoder.BorderKey` implements only the unconditional rule;
the guarded read is deleted, not kept behind a flag. The verdict and the section reference are in
the method's `<remarks>`.

The evidence is pinned as a test rather than left as prose. Reflecting over
`DocumentFormat.OpenXml.dll` standalone did not work, but serialising a `Border` that carries only
the two flags gives `<x:border diagonalUp="1" diagonalDown="1" />` — both on the border element,
no `<diagonal>` child in sight. The SDK's element model is generated from the schema, so that is
the schema speaking. See `The_diagonal_flags_are_attributes_of_border_not_of_its_diagonal_child`.

**The spec expected this to go green in task 4; it went green in task 3.** The decision alone
settles it — the dxf path did not need to move first. The test now also asserts the flags equal
what the element stated, rather than only that two paths agree with each other.

### PREMISE DISPROVED — the dxf table does not grow

`<dxf>` count after 1, 2, 3 and 4 saves: **1, 1, 1, 1**. Flat.

The subset mismatch described in *Divergence 3* is real — the writer's reuse-map decode read four
of six `CT_Dxf` children, the pivot reader five. But it cannot produce growth, because
`AddDifferentialFormats` calls `differentialFormats.RemoveAllChildren()` on the line **immediately
before** `FillDifferentialFormatsCollection`. That method therefore always iterates an empty
collection, the reuse map is always empty, and `ContainsKey` can never miss because there is
nothing to miss. The dxf table is rebuilt from the live object model on every save.

The fixture was checked for vacuity before the premise was declared dead: the alignment-bearing
pivot dxf does reach `XLPivotFormat.DxfStyleValue` with `Alignment.Horizontal == Center` and a
non-default style value, pinned by `An_alignment_bearing_pivot_dxf_reaches_the_pivot_format`. A
flat count is a real result, not an artefact of nothing being loaded.

`Round_tripping_does_not_grow_the_dxf_table` is kept rather than deleted. It is what would catch
the growth if that `RemoveAllChildren` call were ever moved, which would put the reuse map back
into service and make the two decodes have to agree exactly.

The finding that replaces the premise: **`FillDifferentialFormatsCollection` is dead in its stated
role**, and its doc comment ("Populates the differential formats that are currently in the file")
is untrue. Filed as a defect rather than fixed here — it is not this spec's work.

### PREMISE CONFIRMED — the phonetics font decode was a no-op

`LoadPhonetics` called the font decoder with a `PhoneticProperties`, which per the SDK has no child
elements. Deleting the call changed nothing: 44 rich-text tests including the phonetics ones pass
without it. The reset of `Bold`/`Italic`/`Shadow`/`Strikethrough` was not load-bearing.

Landed in task 4 rather than task 6, because deleting `LoadFont` forced the issue — the call could
not survive its decoder.

### The number-format decision the spec asked us to make

On the built-in branch the two paths disagreed: `LoadStyleNumberFormat` left `Format` inherited,
`LoadContext.GetNumberFormat` wrote `string.Empty`.

**Both variants pass the full suite** — measured, not assumed; 14,140 tests green each way. The
suite does not decide it.

Taken: `string.Empty`. `NumberFormatId` is `-1` exactly when the format is custom, so any other id
says the format lives in `XLPredefinedFormat` and no literal belongs in the key beside it; an
inherited custom string next to a built-in id describes a state that cannot occur. Because the
suite does not discriminate, the choice is pinned by
`A_built_in_numFmtId_clears_an_inherited_custom_format_string` so it cannot be reversed by accident.

### Two defects closed that the spec did not name

**An indent in a pivot dxf could fail the load outright.** The pivot path decoded dxf alignments by
writing through `IXLAlignment`, whose `Indent` setter rewrites a `General` horizontal alignment to
`Left` and **throws `ArgumentException`** for any indent above zero on an alignment that is not
left, right or distributed. A workbook whose pivot dxf carried
`<alignment horizontal="center" indent="2"/>` — legal OOXML — could not be opened. Decoding to a key
touches no setter. Pinned by `An_indent_with_a_centred_horizontal_alignment_no_longer_throws` and
`An_indent_alone_no_longer_forces_the_horizontal_alignment_to_left`.

**A duplicated `numFmtId` made a workbook unopenable.** `LoadContext.LoadNumberFormats` used
`Dictionary.Add`, which throws on a duplicate key, and it ran on every load. The unified map uses
`TryAdd`, so the first declaration wins — which is what the linear scan's `FirstOrDefault` already
did. Pinned by `A_duplicated_numFmtId_keeps_the_first_declaration_and_does_not_throw`. Worth noting
that using `Add` for a now workbook-wide map would have *promoted* this from a latent crash to one
on every load; the equivalence was checked rather than assumed.

### The one thing that cost two red runs: write order in `ApplyRunFont`

Two rich-text call sites mutate an `IXLFontBase` (an `XLRichString`), which has no `InnerStyle` to
assign a key to, so `StyleDecoder.ApplyRunFont` writes fields through one at a time.

A first cut wrote all twelve fields; a second wrote only the fields the decode had changed. Both
produced semantically identical workbooks and both broke golden-file tests — `UsingRichText`, then
`UsingRichText` and `AdjustToContents` — with every `<rPr>` byte-identical and only the **order** of
`<si>` entries in `sharedStrings.xml` different.

The cause: a rich run is part of its shared string's identity, so every property write on one
dereferences that string's shared-string-table entry and interns a new one.
`SharedStringTable.GetConsecutiveMap` emits entries in insertion order, so a different set or order
of intermediate writes reorders `sharedStrings.xml` for a file whose content never changed.

`ApplyRunFont` now gates and orders its writes exactly as the decoder it replaces did — the four
booleans unconditional, everything else only when the run states the element — with the
newly-read charset last so that adding it cannot disturb the intermediate states of a run that has
none. **No golden file was regenerated and no assertion was weakened.**

This is worth knowing for spec 29 and anything else that touches rich-text load: the order of
property writes on a rich run is observable in the saved file.

### One unplanned commit: the cellXf fill did not inherit

Re-reading `Decode(CellFormat, …)` against `LoadStyle` branch by branch, as task 3 step 4 says to
do, turned up a discrepancy the suite could not see. Five of the six aspects inherit from the key
passed in; **fill did not** — the original allocated a fresh `XLFill`, so its starting point was the
default fill, and the port passed `key.Fill`.

The original is right on the merits too: a cellXf's `fillId` points at a complete `<fill>`
definition rather than an override, so there is nothing to inherit. The dxf path still inherits,
which is what the `differential` flag distinguishes. Fixed in `ffeea134`; latent rather than live,
because every reachable caller passes a key whose fill is already the default.

### Acceptance criteria

Thirteen of fourteen pass outright. **Criterion 6's literal gate returns 1, not 0.**

The match is `StyleDecoder.ApplyStyle(xlRow, …)` at `WorksheetSheetDataReader.cs:1003` — a
qualified *call* inside `ApplyRowCustomProps`, not a declaration. The criterion's prose
("declares none of …") is satisfied: all ten declarations are gone and the class summary no longer
claims style application.

The spec contradicts itself here: task 6 step 4 lists `ApplyRowCustomProps` among the matches it
**expects** to survive, on the grounds that it is a row *using* a resolved style rather than
decoding one — and that method calls `ApplyStyle`. Removing the match means copying `ApplyStyle`'s
body to the call site. Duplicating logic to make a grep count zero is the wrong trade, so the call
stays and the discrepancy is recorded here instead.

Two criteria come out **stricter** than written: criterion 1's gate returns nothing at all rather
than the permitted colour conversions, and criterion 3's gate lists only `StyleDecoder.cs` rather
than the decoder plus the writers.

### Load is not slower

Allocated, base → branch: `LoadWorkbook` 334.54 → 334.54 MB; `LoadRowHeavy` 55.40 → 55.40 MB;
`Open` 1.67 → 1.66 MB; `OpenAndSaveUnchanged` 3.18 → 3.17 MB; `OpenAndSaveRowHeavyUnchanged`
88.84 → 88.83 MB; `RefreshLookupColumn` 3.82 → 3.81 MB. **Not one increased.**

The reductions are the two allocations this spec removes — a throwaway `XLFill` per distinct cell
fill and an `XLStylizedEmpty` per existing dxf — and they are small because these fixtures carry
few distinct fills and no dxfs. The spec's named risk does not apply: `StylesheetData` is
constructed once per workbook load and `CustomNumberFormats` is derived in its initialiser.

Means all moved inside their own reported margins, which on this machine means nothing (see
DEFECTS D11).

Two forced deviations from the spec's benchmark command: `-f net10.0` is mandatory (the benchmark
project multi-targets and `dotnet run` refuses without it), and the default job cannot be used —
`Program.cs` pins `InProcessEmitToolchain`, which aborts `XLiburReadBenchmarks` with *"takes too
long to run"* after ~50 iterations at ~4.7 s each. Run with `--warmupCount 1 --iterationCount 3
--launchCount 1` instead; allocation figures are exact per operation, and the timings were already
noise-dominated. `LoadAndReadAllCells` was not measured for the same reason and is not named by
criterion 13.

### What the next consumer inherits

One decoder, `StyleDecoder` (556 lines), with three `Decode` entry points — style index,
`CellFormat`, `DifferentialFormat` — over seven key functions. `OpenXmlHelper` drops from 513 lines
to 183 and keeps only its colour layer, the boolean helpers and `GetXLiburTextRotation`.
`WorksheetSheetDataReader` drops from 1,460 to 1,283 and reads sheet data again.

Two asymmetries are now explicit rather than accidental, and both are commented in place:

- **Fill and protection decode against defaults on the cellXfs path**, because an `<xf>` points at
  a complete definition; everything else inherits. The dxf path inherits everything, because a dxf
  is genuinely differential.
- **`RunFontKey` stays separate from `FontKey`**, because `CT_RPrElt` and `CT_Font` spell three
  children with unrelated CLR types. Conflating them is what caused the original defect.
  `RunFontKey` now also reads `<charset>` — one field more than the decoder it replaced, which had
  dropped it on the rich-text path for the same reason it was dropped on the dxf path.

Spec 20 now has one decoder to re-verify its key layout against instead of two.
