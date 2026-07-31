# Changelog

## Contents

- [Unreleased](#unreleased)
- [v0.106.0](#v01060---2026-07-25)

## Unreleased

### ⚠️ Breaking Changes

#### Colours and styles

- **`XLColorType` members are renumbered.** `Automatic` takes ordinal 0, so `Color` moves 0 → 1, `Theme` 1 → 2 and `Indexed` 2 → 3. This is source compatible but binary breaking, and breaks anything that persists the numeric value — stored settings or serialised styles written by an earlier version need remapping. ([#232](https://github.com/XLibur/XLibur/pull/232) by [@jafin](https://github.com/jafin))

- **`XLColor.NoColor.Color` (and `.Indexed`/`.ThemeColor`) now throw** instead of returning a meaningless all-zero ARGB. Test with `IsAutomatic` before reading a colour component. ([#232](https://github.com/XLibur/XLibur/pull/232) by [@jafin](https://github.com/jafin))

  ```csharp
  // Before — returned Color.FromArgb(0, 0, 0, 0) for an automatic colour
  var rgb = cell.Style.Font.FontColor.Color;

  // After
  var colour = cell.Style.Font.FontColor;
  var rgb = colour.IsAutomatic ? defaultRgb : colour.Color;
  ```

  There is no ARGB that means "automatic", which is why the property now throws rather than
  inventing one — decide what your code should render for a colour Excel resolves at display time.

### ✨ New Features

#### Writing large files

- **Streaming write API (`XLStreamingWorkbook`)**: a forward-only writer for exports too large to hold in memory. Rows are serialised straight into the file as they are appended, so nothing is retained per row — a million rows by ten columns costs about 108 MB of peak managed heap, where `XLWorkbook` needs roughly that much for a *tenth* as many rows. On the 50K-row benchmark it is also 1.6× faster than `XLWorkbook` and allocates a fifth as much.

  ```csharp
  using XLibur.Excel.Streaming;

  using var workbook = XLStreamingWorkbook.Create("Large.xlsx");
  var sheet = workbook.AddWorksheet("Data");
  sheet.FreezeRows(1);
  sheet.AppendRow("Name", "Amount");

  for (var i = 0; i < 1_000_000; i++)
      sheet.AppendRow($"Item {i}", i * 1.5);

  workbook.Finish();   // required — disposing without it abandons the write
  ```

  The trade is that it is append-only: rows go in ascending order, one worksheet at a time, nothing can be read back or revised, and formulas are stored verbatim rather than evaluated. Streamed sheets support column widths, freeze panes, an autofilter range, row and per-cell styles, and formulas with cached values. Anything beyond that — tables, merges, conditional formats, pivot tables, drawings — still needs `XLWorkbook`.

  Memory is flat in the number of rows, but not unconditionally flat: under the default `SharedStrings` mode each distinct string is held until `Finish()`, so cost tracks how many *distinct* text values there are rather than how many rows. A million rows of repeating labels is cheap; a million distinct ones is not. `XLStreamingStringStorage.Inline` removes that term entirely at the cost of a larger file, and takes the worst case — every row carrying a distinct string — from 108 MB to 14 MB.

  Because the writer assembles the package itself rather than going through `System.IO.Packaging`, the destination stream does not have to be seekable — a workbook can be written straight to an HTTP response, which `XLWorkbook.SaveAs` cannot do. ([#263](https://github.com/XLibur/XLibur/pull/263) by [@jafin](https://github.com/jafin))

- **`SaveOptions.CompressionLevel`**: choose how hard the package is compressed on an ordinary save. `CompressionLevel.Fastest` trades a larger file for a quicker save, `NoCompression` skips it entirely. Applies to parts the save creates; re-saving a workbook loaded from an existing file leaves its existing parts alone. `XLStreamingOptions.CompressionLevel` does the same for streamed writes, where `Fastest` is about 1.7× quicker than the default. ([#263](https://github.com/XLibur/XLibur/pull/263) by [@jafin](https://github.com/jafin))

#### Formula functions

- **Regression and descriptive statistics — 23 functions**: `LINEST`, `LOGEST`, `TREND`, `GROWTH`, `FREQUENCY`, `FORECAST`, `FORECAST.LINEAR`, `CORREL`, `PEARSON`, `COVARIANCE.P`, `COVARIANCE.S`, `COVAR`, `SLOPE`, `INTERCEPT`, `RSQ`, `STEYX`, `SKEW`, `SKEW.P`, `KURT`, `PROB`, `TRIMMEAN`, `HARMEAN` and `AVEDEV`. The five array-returning ones spill.

  `LINEST` handles several predictors, not just one, and reports the full five-row statistics block on request. Both orientations work — y down a column with each predictor in its own column, or y across a row with each in its own row. ([#257](https://github.com/XLibur/XLibur/pull/257) by [@jafin](https://github.com/jafin))

- **The statistical distributions — 69 functions**. The modern dotted set (`NORM.DIST`, `NORM.INV`, `NORM.S.DIST`, `NORM.S.INV`, `LOGNORM.*`, `CHISQ.*`, `F.*`, `T.DIST*`, `EXPON.DIST`, `POISSON.DIST`, `WEIBULL.DIST`, `GAMMA`, `GAMMA.DIST`, `GAMMA.INV`, `GAMMALN(.PRECISE)`, `BETA.DIST`, `BETA.INV`, `HYPGEOM.DIST`, `NEGBINOM.DIST`, `BINOM.INV`, `CONFIDENCE.NORM`, `CONFIDENCE.T`, `PERCENTILE.EXC`, `QUARTILE.EXC`, `RANK.AVG`, `MODE.MULT`, `PERCENTRANK(.INC/.EXC)`), the hypothesis tests (`CHISQ.TEST`, `F.TEST`, `T.TEST`, `Z.TEST`) and the 26 pre-2010 names (`NORMDIST`, `CHIDIST`, `FINV`, `TDIST`, `CRITBINOM`, …), which are registered against the same implementations rather than copies of them.

  `MODE.MULT` spills. Everything reduces to one of four special functions — the regularized incomplete gamma, the regularized incomplete beta, the error function, or elementary functions — so the inverses invert their own distributions to full double precision rather than to the accuracy of a rational approximation. `T.TEST` with unequal variances truncates the Welch degrees of freedom, as the rest of the `T.DIST` family truncates its own. ([#256](https://github.com/XLibur/XLibur/pull/256) by [@jafin](https://github.com/jafin))

- **`AGGREGATE`, `NETWORKDAYS.INTL` and `WORKDAY.INTL`**. `AGGREGATE` covers all nineteen function numbers — including `PERCENTILE.EXC` and `QUARTILE.EXC`, which have no standalone registration yet — and all eight option values, so ignoring hidden rows and error values both work rather than being documented limitations. The `.INTL` date functions take a weekend as either one of Excel's numbered codes or a seven-character Monday-to-Sunday mask. ([#255](https://github.com/XLibur/XLibur/pull/255) by [@jafin](https://github.com/jafin))

- **20 modern text and array-shaping functions**: `TEXTSPLIT`, `TEXTBEFORE`, `TEXTAFTER`, `VALUETOTEXT`, `ARRAYTOTEXT`, `UNICHAR`, `UNICODE`, `DBCS`, `ENCODEURL`, and the array-shaping set `VSTACK`, `HSTACK`, `TOROW`, `TOCOL`, `WRAPROWS`, `WRAPCOLS`, `CHOOSEROWS`, `CHOOSECOLS`, `TAKE`, `DROP`, `EXPAND`. The array-shaping functions and `TEXTSPLIT` spill.

  `DBCS` derives its mapping by inverting `ASC`'s, so the two are exact inverses. Where Excel would report `#CALC!` — a `DROP` that leaves nothing, a `TOCOL` that ignores every value — XLibur reports `#VALUE!` instead, because the value model has no `#CALC!`. ([#254](https://github.com/XLibur/XLibur/pull/254) by [@jafin](https://github.com/jafin))

- **42 engineering functions**: the complex-number family (`COMPLEX` and all 26 `IM*` functions), `CONVERT` with the full unit table, `BESSELI`/`BESSELJ`/`BESSELK`/`BESSELY`, `ERF`/`ERF.PRECISE`/`ERFC`/`ERFC.PRECISE`, `DELTA`, `GESTEP` and the bitwise set (`BITAND`, `BITOR`, `BITXOR`, `BITLSHIFT`, `BITRSHIFT`).

  A complex number in Excel is text, so the `IM*` functions parse `"3+4i"` and write their result back the same way — echoing whichever of `i` or `j` the input used, and refusing to mix the two. `CONVERT` unit names are case sensitive, as Excel's are: `Pica` is a point and `pica` is six to the inch. Prefixes are accepted on metric units, binary prefixes on `bit` and `byte` only, and on no temperature unit — a scale with an offset has no meaningful "milli". ([#253](https://github.com/XLibur/XLibur/pull/253) by [@jafin](https://github.com/jafin))

- **24 more financial functions**: depreciation (`SLN`, `SYD`, `DB`, `DDB`, `VDB`), rate conversion and growth (`EFFECT`, `NOMINAL`, `RRI`, `PDURATION`), fractional dollar notation (`DOLLARDE`, `DOLLARFR`), loan schedules (`ISPMT`, `CUMIPMT`, `CUMPRINC`), discount securities (`TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`, `DISC`, `INTRATE`, `RECEIVED`) and irregular cash flows (`FVSCHEDULE`, `MIRR`, `XNPV`, `XIRR`). `XIRR` solves with Newton–Raphson and falls back to bisection when that wanders off, so a poor guess still converges.

  Two deliberate limitations. The day-count-basis bond family (`PRICE`, `YIELD`, `DURATION`, `MDURATION`, `ACCRINT`, `COUP*`, `ODD*`, `AMOR*`) is not included — it needs a coupon-period engine rather than another day-count fraction. And `DISC`/`INTRATE`/`RECEIVED` take their year fraction from the same code as `YEARFRAC`, so basis 1 (actual/actual) uses `YEARFRAC`'s average-year-length rule. ([#252](https://github.com/XLibur/XLibur/pull/252) by [@jafin](https://github.com/jafin))

#### Security and encryption

- **Password-protected workbooks (ECMA-376 encryption)**: `LoadOptions.Password` opens an encrypted workbook and `SaveOptions.Password` writes one. Reading covers both schemes in the wild — agile encryption (Office 2010 and later) and standard encryption (Office 2007); writing uses agile encryption with the parameters Excel itself writes (AES-256-CBC, SHA-512, 100,000 spins and a fresh random key per save). Previously an encrypted file could not be opened at all and failed opaquely.

  A wrong or missing password throws `XLInvalidPasswordException`, which is what a caller re-prompts on; an altered package whose HMAC fails throws `XLEncryptionException`, because there the password was right. A legacy `.xls` is detected and named rather than mistaken for an encrypted workbook. The password is never carried from load to save — a plain `SaveAs` cannot silently produce a file the caller can no longer open. ([#245](https://github.com/XLibur/XLibur/pull/245) by [@jafin](https://github.com/jafin))

#### Charts

- **Chart series formatting**: `IXLChartSeries` gained `FillColor`, `LineColor`, `LineWidthPt`, `MarkerStyle` (new `XLMarkerStyle` enum), `MarkerSize`, `MarkerFillColor` and `Smooth`, so a generated chart can be styled instead of relying on Excel's automatic theme colours. Leaving a property `null` omits its element, which keeps the automatic colour — nothing is ever written as an explicit black. ([#220](https://github.com/XLibur/XLibur/pull/220) by [@jafin](https://github.com/jafin))

- **Secondary value axis per series**: `IXLChartSeries.UseSecondaryAxis` plots a series against a value axis on the right, so a percentage can share a chart with values in the thousands. It applies to series of the primary chart type as well as to a combo chart's `SecondarySeries`. ([#220](https://github.com/XLibur/XLibur/pull/220) by [@jafin](https://github.com/jafin))

- **Chart data labels**: `IXLDataLabels` on both `IXLChart.DataLabels` (chart-wide) and `IXLChartSeries.DataLabels` (per series, overriding the chart's), with `ShowValue`, `ShowCategoryName`, `ShowSeriesName`, `ShowPercentage`, `NumberFormat` and `Position`. `Position` is validated against the chart type — Excel refuses to open a file that uses a position it does not offer for that type, so the setter throws with the allowed values listed rather than producing a workbook Excel has to repair. ([#221](https://github.com/XLibur/XLibur/pull/221) by [@jafin](https://github.com/jafin))

- **Chart legend**: `IXLChart.Legend` with `Visible`, `Position` (right, bottom, left, top, top-right) and `Overlay`. Charts XLibur creates still have no legend unless one is asked for; setting `Visible = false` on a chart read from a file removes the legend it came with. ([#222](https://github.com/XLibur/XLibur/pull/222) by [@jafin](https://github.com/jafin))

- **Chart axes**: `IXLChart.CategoryAxis`, `ValueAxis` and `SecondaryValueAxis`, each with `Title`, `NumberFormat`, `Min`, `Max`, `MajorUnit`, `MinorUnit`, `Visible`, `MajorGridlines`, `Orientation` (reversed axes) and `LogScale`/`LogBase`. The unit and log-scale properties belong to a value axis in the file format and are skipped on a category axis — except on scatter and bubble charts, whose horizontal axis holds numbers. ([#222](https://github.com/XLibur/XLibur/pull/222) by [@jafin](https://github.com/jafin))

- **Charts loaded from a file can be restyled**: setting the series formatting, data labels, legend or axes on a loaded chart now writes back on save. Only the properties actually assigned are patched into the existing chart part, so trendlines, error bars, gradient fills, per-point colours and label overrides, label and axis fonts, tick marks and the chart's style/colour parts are all preserved — and a chart nobody edited is left byte for byte as it was. ([#220](https://github.com/XLibur/XLibur/pull/220), [#221](https://github.com/XLibur/XLibur/pull/221), [#222](https://github.com/XLibur/XLibur/pull/222) by [@jafin](https://github.com/jafin))

- **Chart anchoring**: `IXLChart.Anchor` (`MoveAndSizeWithCells`, `MoveWithCells`, `Absolute`) with `Width`, `Height`, `Left` and `Top` in pixels, so a chart can keep its size as rows are inserted or be pinned to a spot on the sheet. Two-cell anchoring via `Position`/`SecondPosition` remains the default. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

#### Colours and styles

- **`XLColorType.Automatic`**, with `XLColor.Automatic` and `XLColor.IsAutomatic`. OOXML has four kinds of colour and XLibur modelled three; the automatic colour (ECMA-376 `CT_Color/@auto`, what Excel's font colour picker labels "Automatic") was disguised as a fully transparent black. `auto="1"` is now read and written explicitly. ([#232](https://github.com/XLibur/XLibur/pull/232) by [@jafin](https://github.com/jafin))

### ⚡ Performance

#### Workbook creation

- **40% fewer allocations building a workbook** (create phase of the 50K x 10 benchmark, 305.5 → 183.6 MB; whole benchmark 423.4 → 301.3 MB). Every `cell.Value = x` ran a merged-range membership test that allocated ~250 bytes even on a sheet with no merged ranges at all, boxing the address and building a LINQ iterator chain over an empty list. That test is now skipped outright when the sheet has no merges, and made allocation-free when it does — a merged title row is a common layout and used to cost *more* than no merge at all, because the predicate then actually ran. A cell write drops 375.5 → 103.6 bytes on both paths. ([#185](https://github.com/XLibur/XLibur/pull/185) by [@jafin](https://github.com/jafin))

- **Repeated access to the same cell is cheaper**: `ws.Cell(r, c)` minted a fresh `XLCell` every call, so the per-object style cache never hit and a `.Style` access threw away an 80-byte `XLStyle` one statement later. A small direct-mapped cache now hands back the same wrapper for the same address. ([#185](https://github.com/XLibur/XLibur/pull/185) by [@jafin](https://github.com/jafin))

#### Styling

- **86% fewer allocations styling a range, row or column in bulk** (~234 → ~33 bytes per cell): setting a style on a container collected every child cell into a `HashSet` of wrappers before writing anything. Containers that can enumerate exactly those cells now walk the addresses and write the style slice directly. ([#185](https://github.com/XLibur/XLibur/pull/185) by [@jafin](https://github.com/jafin))

#### Saving

- **50% fewer allocations on save** (237.1 → 117.9 MB for the same benchmark), with wall time improving from 1187–1290 ms to 1051–1167 ms. Four save helpers materialised a cell wrapper for every used cell just to read one property and now read the underlying storage directly; `int`/`uint` cell values are formatted into a span instead of allocating a string each; style-key hashes are memoised, cutting `StyleKey.GetHashCode` by 95% over 100K styles. Saved output is byte-identical to before. ([#179](https://github.com/XLibur/XLibur/pull/179) by [@jafin](https://github.com/jafin))

### 🐛 Bug Fixes

#### Formulas and references

- **`SUBTOTAL` no longer counts a nested post-2007 function twice.** The check that stops a subtotal counting a subtotal inside its own range compared the function name as the formula stores it, and Excel stores every function added after 2007 under an `_xlfn.` namespace — so a nested `AGGREGATE` was never recognised. The namespace is now stripped before the comparison. ([#255](https://github.com/XLibur/XLibur/pull/255) by [@jafin](https://github.com/jafin))

- **A reference whose rows or columns are all deleted becomes `#REF!`** ([ClosedXML #880](https://github.com/ClosedXML/ClosedXML/issues/880)): endpoints were shifted by the deleted height and clamped to row 1, so deleting rows 1–5 turned `Sheet1!$A$1:$B$2` into `Sheet1!$A$1:$B$1` — a phantom one-row range over whatever data had moved up into it. The same shifter serves cell formulas, so `=SUM(A1:A2)` with those rows deleted now reads `SUM(#REF!)` rather than quietly summing the wrong cells. Deleting the sheet afterwards drops the stale sheet prefix instead of leaving a defined name pointing at a sheet that no longer exists. ([#243](https://github.com/XLibur/XLibur/pull/243) by [@jafin](https://github.com/jafin))

- **Row-only and column-only references are positioned correctly when rows or columns are inserted or deleted**: `3:5` with row 4 deleted became `2:4` — a reference that had walked onto row 2, which it never covered, while losing row 5, which survived. Because a multi-row delete is applied one row at a time, the reference kept drifting up instead of shrinking and never became small enough for the `#REF!` check to fire. Insertions had the mirror problem: inserting two rows at row 4 moved `3:5` to `5:7` rather than expanding it to `3:7`. Both axes now follow the same boundary rules as an equivalent cell range. ([#244](https://github.com/XLibur/XLibur/pull/244) by [@jafin](https://github.com/jafin))

- **Copying a worksheet repoints the copy's self-references at the copy** ([ClosedXML #836](https://github.com/ClosedXML/ClosedXML/issues/836)): copying a sheet named `Original` holding `Original!A1 * 3` produced a sheet whose formula still pointed back at the original. References to *other* sheets are left alone, as they should be. ([#241](https://github.com/XLibur/XLibur/pull/241) by [@jafin](https://github.com/jafin))

#### Text and number parsing

- **Text-to-number and text-to-date coercion matches Excel much more closely.** The date-time patterns had no seconds component, so `8/22/2008 3:30:45 PM` failed to coerce at all. Beyond that: time components that overflow their range now carry into the date (`11/30/2022 24:59` is one minute to one in the morning of December 1st); parenthesised, spaced and sign-separated numbers such as `(100%)`, `(   1,000.54  )` and `- 100 %` are read as negative; a month is matched by any prefix from three letters up, as Excel does; a shortened or dot-suffixed AM/PM designator is accepted. In the other direction, group-separator and currency placement are now enforced the way Excel enforces them — `1,00`, `1,00,000` and `1$` are rejected under en-US where the underlying BCL parse accepted them. ([#241](https://github.com/XLibur/XLibur/pull/241) by [@jafin](https://github.com/jafin))

#### Rich text and shared strings

- **Rich-text runs keep their automatic or absent colour through a round-trip**: a plain load → `SaveAs` wrote every colour-less run back with an explicit `<color rgb="FF000000"/>`, and promoted a plain string carrying a phonetic guide (`<t>` + `<rPh>`, common in Japanese workbooks) into a synthetic run that inherited the same injection. An explicit black cannot be overridden by a theme colour change or by conditional formatting, where the automatic colour it replaced can. A run that was read with no `<rPr>` is now written back without one, so inheritance stays inheritance; splitting such a run keeps that property rather than materialising the cell font as explicit formatting. ([#225](https://github.com/XLibur/XLibur/pull/225), [#227](https://github.com/XLibur/XLibur/pull/227) by [@jafin](https://github.com/jafin))

- **Saving a plain shared string that carries a phonetic guide no longer throws.** Its text is decoded on the way in, so a decoded `_xHHHH_` escape is a literal control character — not valid XML content — and the writer emitted it raw, failing with an `ArgumentException` naming the invalid character. ([#225](https://github.com/XLibur/XLibur/pull/225), [#227](https://github.com/XLibur/XLibur/pull/227) by [@jafin](https://github.com/jafin))

#### Colours and conditional formatting

- **An automatic colour is no longer written as an explicit `rgb="00000000"`.** Colour writers switch on the colour type, and the automatic colour fell into the RGB arm — pinning down a colour the source deliberately left for the application to resolve. The three conditional-format colour converters gained explicit automatic arms too, where they would otherwise have dropped the colour silently. ([#232](https://github.com/XLibur/XLibur/pull/232) by [@jafin](https://github.com/jafin))

#### Charts

- **Charts anchored with a one-cell or absolute anchor are no longer dropped on load.** The reader only looked at `xdr:twoCellAnchor`, so a chart Excel had anchored either of the other two ways was missing from `IXLWorksheet.Charts` entirely (its XML survived a round trip, but the chart was invisible to the API). ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

- **3D and of-pie chart groups are read.** `c:pie3DChart`, `c:line3DChart`, `c:area3DChart`, `c:surface3DChart` and `c:ofPieChart` were not recognised, so an Excel-authored chart using one loaded with no series and the wrong chart type. Their series and series formatting now read the same as the 2D groups', and pie-of-pie and bar-of-pie are told apart. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

- **Chart XML now passes OpenXML schema validation.** Three long-standing violations in the chart writer are fixed: series names were written as a `c:strRef` with no required `c:f` (a literal name now uses `<c:tx><c:v>`, and both forms are read back), `c:doughnutChart` omitted the required `c:holeSize`, and `c:marker` was written after `c:cat`/`c:val` instead of before. Excel tolerated all three, but stricter readers and `SaveOptions.ValidatePackage` did not. ([#220](https://github.com/XLibur/XLibur/pull/220) by [@jafin](https://github.com/jafin))

- **A `Line` chart whose markers are switched off no longer reads back as `LineWithMarkers`.** The reader treated the presence of a `c:marker` element as "has markers", even when it held `<c:symbol val="none"/>`. ([#220](https://github.com/XLibur/XLibur/pull/220) by [@jafin](https://github.com/jafin))

- **Charts with more than one plot group of the same type now read all of their series.** The reader took only the first `c:barChart` (or `c:lineChart`, …) of a plot area, so the series of a second group — which is how Excel stores a secondary axis — were dropped. ([#220](https://github.com/XLibur/XLibur/pull/220) by [@jafin](https://github.com/jafin))

- **Which of those groups is on the secondary axis no longer depends on the order they appear in the file.** The primary axis pair was taken from whichever group came first, so a file that wrote its secondary group ahead of the primary one read back with `UseSecondaryAxis` inverted on every series and the two axis models swapped. The group whose value axis crosses at the maximum — how a secondary axis comes to be drawn on the right — is now passed over instead. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

- **`Smooth` is honoured on a new stock chart.** A stock chart's series are `CT_LineSer` and take `c:smooth`, but the writer never emitted it, so the property worked on a stock chart read from a file and was silently dropped on one XLibur created. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

- **`IXLChart.SetTitle` on a chart loaded from a file is no longer silently discarded** ([#272](https://github.com/XLibur/XLibur/issues/272)): the patcher rewrote the legend, axes, data labels and series of a loaded chart but never its title, so the model accepted the new title, `chart.Title` read it back, and the saved part kept the old one — with no exception and no warning. The title is now patched in like everything else. Only the text is replaced, so a title that came out of Excel keeps its font, colour, layout and overlay; setting it to `null` removes the title and sets `c:autoTitleDeleted` so Excel does not put an automatic one back in its place. This covers extended (ChartEx) charts — waterfall, funnel, treemap, sunburst, box &amp; whisker — as well as the standard ones; the title is the only edit those take, since they carry none of the formatting XLibur models. ([#283](https://github.com/XLibur/XLibur/pull/283) by [@jafin](https://github.com/jafin))

- **Positioning a legend that is not there no longer creates one.** `IXLChartLegend.Position` and `Overlay` are documented as ignored while `Visible` is `false`, and a new chart gets no legend from them — but assigning one of them on a *loaded* chart that had no legend added one. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

- **Chart-wide data labels reach every group of a loaded combo chart.** `IXLChart.DataLabels` applies to the whole chart, and a new combo chart gets them on both of its plot groups, but a loaded one was patched on the primary group only — so turning labels on left the secondary series unlabelled. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

- **`Series.Add(...)` on a chart loaded from a file throws instead of being discarded on save.** A loaded chart is patched, not regenerated, so a new series had nowhere to be written and vanished without a word. It now throws `NotSupportedException`, as `UseSecondaryAxis` already did. ([#230](https://github.com/XLibur/XLibur/pull/230) by [@jafin](https://github.com/jafin))

#### Comments

- **`IXLComment.Delete()` removes the note from where it is now, not where it was created** ([#259](https://github.com/XLibur/XLibur/issues/259)): a note remembered the cell it was constructed with, and inserting or deleting rows or columns moves the note without telling it. Deleting a note on `A5` after two rows were inserted above cleared `A5` — by then an empty cell — and left the note sitting on `A7`. The note is now confirmed to still be the one at the remembered address before that address is cleared, and located by identity when it is not. ([#268](https://github.com/XLibur/XLibur/pull/268) by [@jafin](https://github.com/jafin))

### 🗑️ Deprecations

#### Colours and styles

- **`XLColor.NoColor` is deprecated in favour of `XLColor.Automatic`.** Some Excel pickers (sheet tab, fill background) label the same value "No Color" — that is a GUI convention, not a different value. `NoColor` still compiles but now warns, which is an error for consumers building with `TreatWarningsAsErrors`. ([#232](https://github.com/XLibur/XLibur/pull/232) by [@jafin](https://github.com/jafin))

  ```csharp
  // Before
  cell.Style.Font.FontColor = XLColor.NoColor;

  // After
  cell.Style.Font.FontColor = XLColor.Automatic;
  ```

  A straight rename — `NoColor` returns `Automatic`, so the two are the same value.

## v0.106.0 - 2026-07-25

First XLibur release since forking [ClosedXML v0.105.0](https://github.com/ClosedXML/ClosedXML/)
(May 2025). Everything below is relative to that baseline.

### Added

- **Charts — all 78 `XLChartType` values**: End-to-end chart creation, saving, loading and round-tripping. Covers bar/column (clustered, stacked, percent, 2D and 3D), the 21 Bar3D cone/cylinder/pyramid shapes, line, area, pie/doughnut (including pie-to-pie and pie-to-bar), radar, scatter/XY, bubble, and surface types, plus data series, chart titles, combo charts and positioning. The previous `IXLChart`/`XLChart` stubs are now backed by a real implementation.

- **Dynamic arrays**: Modern array functions — `SEQUENCE`, `UNIQUE`, `SORT`, `SORTBY`, `FILTER`, `XLOOKUP` and `XMATCH` — together with a **spill engine**. A dynamic-array formula written into a single cell now auto-fills its computed footprint into the neighbouring cells, grows and shrinks as the result changes, and round-trips through save/load. Only the anchor cell holds the formula; spilled cells stay formula-less, matching Excel. A footprint blocked by existing content, or one that would run past the sheet edge, collapses to the new `#SPILL!` error (`XLError.SpillRange`) on the anchor.

- **New worksheet functions**:
  - Conditional aggregates: `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`
  - Logical: `IFS`, `SWITCH`
  - Statistical: `SMALL`, `RANK`, `PERCENTILE`, `QUARTILE`, `MODE`
  - Financial: `PV`, `NPV`, `IRR`, `RATE`, `NPER`, `PPMT`
  - Reference: `INDIRECT`

- **Wildcard support in `HLOOKUP` and `VLOOKUP`**: `*` and `?` patterns now match in lookup values, as they do in Excel.

- **Swappable font engine, and a font-library-free core**: Text measurement (column auto-fit, row heights, glyph metrics) moved behind `IXLFontEngine`, and the font library ships as a separate package rather than being compiled into the core assembly. The MIT-licensed SkiaSharp engine is the default and auto-registers the first time a workbook is created, so no startup call is needed. This lets you choose a font library whose licence suits you, and stops library authors inheriting a font dependency they don't need. See the Upgrade Guide below.

- **`XLibur.Bundle` meta-package**: Installs the core library together with the default font engine, so a single package reference behaves like ClosedXML out of the box.

- **Editable pictures inside group shapes**: Pictures nested in `xdr:grpSp` groups — at any nesting depth — can be read, resized, moved, added, removed and grouped through a first-class public API. Geometry is computed through the composed group transform, and moves operate in sheet space.

- **DataBar conditional formats can be modified after creation**, including axis settings, rather than being write-once at creation time.

- **Pivot table improvements**: Named ranges resolve as a pivot cache source, and `autoSortScope` on pivot fields round-trips through load/save.

### Fixed

- **Array and dynamic-array formulas no longer break on row/column shifts**: Inserting or deleting rows/columns anywhere in a workbook used to rebuild every formula cell through the `FormulaA1` setter, which turned a single array formula (shared across its whole range) into one *normal* formula per cell. For dynamic arrays this split a single spilled formula such as `=UNIQUE(...)` into multiple implicit-intersection `=@UNIQUE(...)` cells, even when the edit happened on an unrelated sheet. Shifts now update the shared formula instance in place — preserving its array/dynamic-array nature — and relocate the spill range for same-sheet inserts/deletes.

- **Deleting through an array no longer corrupts its stored range**: When a delete overlapped an array formula, relocating the array's top edge could drive the coordinate below 1. `Point` does not bounds-check, so the value silently overflowed and corrupted the stored range.

- **Data-validation formulas are shifted with the sheet**: Inserting or deleting rows/columns relocated each rule's ranges (`sqref`) but left cell references *inside* the criteria formulas (`formula1`/`formula2`) pointing at the pre-shift location. Any `List`, `Custom` or comparison rule referencing other cells silently broke — most visibly dependent dropdown pairs driven by `OFFSET`/`MATCH`. The in-memory value was wrong immediately after the shift, before any save.

- **Data validations no longer vanish when inserting at row 1 or column 1**: The data-validation index was keyed by address at insert time and never re-keyed, so an insert at the first row/column left it stale. At save time the split logic then treated a rule's own out-of-date entry as a competing rule and stripped its ranges, emitting `<dataValidation sqref="">`. Excel rejected the file on open with *"Removed Records: Data validation"*. The index is now reconciled before consolidation.

- **Conditional-format ranges shift once, not twice** ([ClosedXML #2850](https://github.com/ClosedXML/ClosedXML/issues/2850)): Inserting rows or columns below the first line doubled the shift for any rule whose shifted target address collided with another rule's existing range. A rule at `K13` that should move to `K23` landed at `K33`, while rules whose targets happened to be empty shifted correctly.

- **Page breaks no longer inflate the used range** ([ClosedXML #2842](https://github.com/ClosedXML/ClosedXML/issues/2842)): `AddHorizontalPageBreak()`/`AddVerticalPageBreak()` wrote `brk@max` as the sheet's full row/column count. Excel read that as a huge used range, so a file with ~2000 rows of data rendered with a scrollbar spanning all 1,048,576.

- **Named ranges shrink correctly when their first row or column is deleted**: Deleting the first row of a named range shifted both endpoints up instead of removing the deleted row and shifting the survivors, so `A3:A4` became `A2:A3` — expanding the range to include a row that was never part of it. Excel produces `A3:A3`.

- **Totals-row formulas escape column names containing spaces**: Structured references for headers such as `Feb 2023` used the single-bracket form, producing a formula Excel could not parse.

- **Grouped pictures and shapes survive a load/save round-trip** instead of being dropped.

- **Cached formula values are preserved on save**: Cached values are now written whenever they exist and the formula has not been dirtied, regardless of `EvaluateFormulasBeforeSaving`, and the data-type attribute is preserved. This fixes round-trip loss of dynamic-array results (`SORT`, `UNIQUE`, `FILTER`) and spill cell values.

- **Pivot table alignment formatting round-trips**: Alignment in pivot table differential formats (DXF) was silently lost on load/save.

### Performance

- **61% fewer allocations and 16.5% less wall time on load** (250K rows x 15 columns benchmark), from removing per-cell and per-entry garbage in the shared-string reader, cell value/attribute reads, and a new style cache.

- **`<sheetData>` is read with a raw `XmlReader`**: Worksheet loading — the dominant cost when opening a workbook — no longer goes through the OpenXML SDK's `OpenXmlPartReader`, which rebuilt a `ReadOnlyCollection<OpenXmlAttribute>` and materialized text through its object model for every `<c>`, `<row>` and `<f>` element. Measured in isolation on a 250K x 15 sheet (3.75M cells), that reader accounted for ~67% of load time and ~80% of load allocations — roughly 4x slower and 5x more garbage than an equivalent raw `XmlReader` traversal.

- **Faster string cell reads**: `GetValue<string>()`/`GetString()` — the most common cell read — no longer runs a compiled regex over the whole string (allocating a `MatchCollection`) to find the rare `_xHHHH_` escape sequence.

- **Reduced allocations in 10 per-cell, per-formula and per-address hot-path methods**, with no public API or behaviour change.

- **Load and save hot paths**: The shared-string reader is pre-allocated from the SST count, merged cells stream instead of building a full DOM, worksheet attributes are parsed in a single pass, calc-engine overhead is skipped for formula cells during load, and `uint` boxing was removed from the XML writer.

- **`XmlEncoder.EncodeString` fast-path**: Added a character scan that short-circuits before the `Regex` and `StringBuilder` when a string contains no characters that need encoding (the common case for plain text). For workbooks with ~50K unique shared strings this eliminates ~50K `StringBuilder` allocations, ~50K regex evaluations, and ~50K string copies on save.

- **`IXLWorksheet.SetCellValue(int row, int column, XLCellValue value)`** (new API): Sets a cell value directly on the worksheet's internal storage without allocating an intermediate `XLCell` object. For bulk data population (e.g. 50K rows x 3 columns) this eliminates ~150K object allocations that the `Cell(row, col).SetValue(...)` pattern would create.

### Upgrade Guide

#### Migrating from ClosedXML

The public API surface is largely unchanged from ClosedXML 0.105. To migrate:

1. Install `XLibur.Bundle` from NuGet.
2. Replace `using ClosedXML` namespace references with `using XLibur`.

Namespaces are prefixed with `XLibur` so both libraries can be referenced in the same project.

#### Font engine packaging

This is the one area where XLibur's packaging differs from ClosedXML. ClosedXML compiles
[SixLabors.Fonts](https://github.com/SixLabors/Fonts) into its core assembly; XLibur keeps the core
assembly free of any font library and ships the engine as a separate, swappable package.

- **Installing `XLibur.Bundle` (or `XLibur` + `XLibur.Fonts.SkiaSharp`) requires no code changes.**
  The default SkiaSharp engine auto-registers on first workbook creation. It resolves system fonts
  and falls back to an embedded, metric-only Calibri-compatible font, so text measurement works in
  headless and serverless environments with no system fonts installed.
- **Installing the bare `XLibur` package with no font engine** throws an `InvalidOperationException`
  when a workbook is created, telling you to add a font engine package. This is intentional — it is
  how the core stays font-library-agnostic.
- **To keep ClosedXML 0.105's exact engine**, install `XLibur.Fonts.SixLabors.V1` and call
  `SixLaborsV1FontBootstrap.Register()` at startup.

See [docs/font-architecture.md](docs/font-architecture.md) for the full design and the list of
available engines.

#### Using `SetCellValue` for bulk writes

The existing `Cell(row, col).SetValue(value)` API continues to work and remains the correct choice when you need full cell semantics (formula clearing, merged-range checks, table header refresh). No code changes are required.

For **performance-critical bulk data population** where you are writing values into empty or freshly-created cells, you can switch to the new direct API:

```csharp
// Before (allocates an XLCell per call):
for (int row = 1; row <= 50_000; row++)
{
    ws.Cell(row, 1).SetValue(row);
    ws.Cell(row, 2).SetValue($"Item {row}");
    ws.Cell(row, 3).SetValue(row * 1.5);
}

// After (zero intermediate allocations):
for (int row = 1; row <= 50_000; row++)
{
    ws.SetCellValue(row, 1, row);
    ws.SetCellValue(row, 2, $"Item {row}");
    ws.SetCellValue(row, 3, row * 1.5);
}
```

`SetCellValue` handles date/time number format application and quote-prefix stripping, so the resulting cell content and formatting is identical for data values. The following behaviors are **not** performed by `SetCellValue` — use `Cell().SetValue()` if you need them:

| Behavior | `Cell().SetValue()` | `SetCellValue()` |
|---|---|---|
| Set value and number format | Yes | Yes |
| Clear existing formula | Yes | No |
| Check merged range (inferior cell skip) | Yes | No |
| Refresh table header fields | Yes | No |
