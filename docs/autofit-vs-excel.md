# Autofit compared with Excel

Measurements behind #789: why `IXLColumn.AdjustToContents` and `IXLRow.AdjustToContents` give
other sizes than Excel's AutoFit, which part was fixed, and which part is left.

## How it was measured

- Excel: Microsoft 365, 16.0 build 20430, driven over COM. The display runs at 120 DPI. No
  measurement was made at 96 DPI.
- The Normal style was Calibri 11 (digit 9 px at 120 DPI) unless stated.
- Column: type the text into a cell, set its font, `Columns(n).AutoFit()`, read `ColumnWidth`
  and `Width` (points; px = pt × 120 / 72).
- Row: put text with hard line breaks in a wrapped cell, `Rows(n).AutoFit()`, read `Height`.
- For comparison, the text width that GDI gives: `CreateFont` with a height of
  `-round(size × 120 / 72)` px and `GetTextExtentPoint32`.

## Column width = text + padding

For text `W` px wide (as GDI measures it), Excel's autofit width in px is

```
W + 2 * ceil(cellMdw / 4) + 1 + max(2, floor(W / (4 * mdw)))
```

- `cellMdw`: the digit width of the cell's font (Calibri 11: 9 px, Calibri 20: 17 px).
- `mdw`: the digit width of the column, which is the first font's (#786).

This formula with the GDI width gave Excel's width to the pixel for all 108 strings of the survey
(12 strings in Calibri 8, 11, 14 and 20, Arial 10 and 20, Times New Roman 12, Consolas 11 and
Segoe UI 9), and for runs of 1 to 120 glyphs of `i` (Calibri 11), `0` (Calibri 11 and 20, Times
New Roman 12).

The last term steps up at multiples of `4 × mdw`. With the Normal style of Calibri 11 the steps
are at 108, 144, 180, … px. With a Normal style of Calibri 20 (17 px digits) they are at 204,
272, 340, … px.

Wrapping uses the same rule. In a column of 79, 151, 295 and 439 px, the longest run of `i` that
Excel kept on one line was 17, 35, 70 and 105 glyphs (Calibri 11), 8, 17, 34 and 52 glyphs
(Calibri 20), and 7, 14, 28 and 42 zeros (Times New Roman 12): the longest whose padded width
fits the column.

XLibur used `W + 2 * ceil(0.03 * W + cellMdw / 4) + 1`, which pads more as the text gets longer.
It now uses the measured formula (`XLColumn.GetCellWidthPx`, tests in `AutofitPaddingTests`).

## Text measurement (left as it is)

Excel measures text with GDI: the font is made at an em of whole pixels (Calibri 11 at 120 DPI is
18 px, not 18.33) and the advance widths are hinted to whole pixels. The font engines scale the
outline advance widths to the exact em and do not hint. So the engine's text width differs from
Excel's, most often wider:

| Text, font | GDI (Excel) | engine |
|---|---|---|
| `The quick brown fox jumps`, Calibri 11 | 195 px | 201 px |
| same, Calibri 20 | 357 px | 366 px |
| 16 × `0`, Calibri 11 | 144 px | 149 px |
| 10 × `i`, Calibri 8 | 40 px | 31 px |
| 10 × `W`, Arial 20 | 340 px | 315 px |

Over the 108 strings of the survey (test font engine, Calibri installed):

| | mean difference | mean abs. difference | within 1 px |
|---|---|---|---|
| before #789 | +6.4 px | 7.3 px (4.6 %) | 27 of 108 |
| after #789 | +2.1 px | 3.5 px (2.7 %) | 49 of 108 |

The rest is the hinting. Scaling the advance widths to the whole-pixel em alone does not fix it:
it moves the mean difference to +0.9 px, but the mean absolute difference only to 3.1 px, and it
makes some fonts worse (Calibri 8, Arial 10). Matching Excel exactly needs hinted advance widths
(a TrueType hinting engine, or the font's `hdmx` table), which the engines do not have.

The digit width shows the whole-pixel em too: Arial 20 at 120 DPI is 18 px in Excel (em 33 px)
but 19 px in the engine (em 33.33 px). That is a change to the column unit, not to autofit, and is
not made here.

## Row height (left as it is)

In the wrapped fixture of `FirstFontDigitWidthTests` XLibur breaks the text into as many lines as
Excel does. The height differs only because a line is lower: XLibur makes a line the em plus the
descent, each rounded (`GlyphBox.LineHeight`), Calibri 11 at 120 DPI 18 + 5 = 23 px, where Excel's
lines are 24 px.

Excel's line pitch, from rows of 2, 3 and 4 lines, Calibri, 120 DPI:

| pt | 7 | 8 | 9 | 10 | 11 | 12 | 14 | 16 | 18 | 19 | 20 | 24 | 31 | 32 | 36 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Excel line, px | 16 | 17 | 20 | 23 | 24 | 26 | 30 | 35 | 39 | 42 | 43 | 52 | 67 | 68 | 77 |
| GDI `tmHeight`, px | 14 | 15 | 18 | 21 | 22 | 24 | 28 | 33 | 37 | 39 | 40 | 49 | 64 | 64 | 73 |
| XLibur line, px | 15 | 16 | 19 | 21 | 23 | 25 | 29 | 34 | 38 | 40 | 41 | 50 | 65 | 66 | 75 |

Excel's line is GDI's `tmHeight + tmExternalLeading` plus 1 to 6 px. The extra depends on the font
and on the size: in Calibri it is 2 px up to 18 pt, 3 px from 19 to 31 pt and 4 px from 32 pt, and
Calibri 31 pt and 32 pt have the same GDI metrics (64 px) but lines of 67 px and 68 px. In Consolas
it steps up at 12, 19, 27 and 34 pt; in Arial, Times New Roman, Georgia and Segoe UI it goes up and
down with the size. No formula was found that fits all eight fonts surveyed (Calibri, Arial, Times
New Roman, Segoe UI, Cambria, Verdana, Consolas, Georgia, 6 to 36 pt). A formula would also need
the font's Windows ascent and descent and its line gap, which `IXLFontEngine` does not give.

One more effect: the empty cells of a row take part in its height with the Normal style's font.
With a Normal style of Calibri 11, a wrapped Calibri 8 cell of n lines is `17n + 2` px, not `17n`,
because the row's descent below the last baseline is Calibri 11's (5 px), not Calibri 8's (3 px).
A one-line row is never lower than the Normal style's row (24 px). XLibur does not model this.
