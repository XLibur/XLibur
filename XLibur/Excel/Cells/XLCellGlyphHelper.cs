using System;
using System.Collections.Generic;
using System.Globalization;
using XLibur.Graphics;

namespace XLibur.Excel;

internal static class XLCellGlyphHelper
{
    /// <summary>
    /// Get glyph bounding boxes for each grapheme in the text. Box size is determined according to
    /// the font of a grapheme. New lines are represented as the default (all dimensions zero) box.
    /// A line without any text (i.e., contains only new line) should be represented by a box
    /// with zero advance width, but with a line height of corresponding font.
    /// </summary>
    /// <param name="cell">Cell to get glyph boxes for.</param>
    /// <param name="engine">Engine used to determine box size.</param>
    /// <param name="dpi">DPI used to determine the size of glyphs.</param>
    /// <param name="output">List where items are added.</param>
    internal static void GetGlyphBoxes(XLCell cell, IXLGraphicEngine engine, Dpi dpi, List<GlyphBox> output)
    {
        var richText = cell.RichText;
        if (richText is not null)
        {
            foreach (var richTextRun in richText.Runs)
            {
                var text = richText.GetRunText(richTextRun);
                var font = new XLFont(richTextRun.Font.Key);
                AddGlyphs(text, font, engine, dpi, output);
            }
        }
        else
        {
            var text = cell.GetFormattedString();
            AddGlyphs(text, cell.Style.Font, engine, dpi, output);
        }
    }

    private static void AddGlyphs(string text, IXLFontBase font, IXLGraphicEngine engine, Dpi dpi, List<GlyphBox> output)
    {
        Span<int> zeroWidthJoiner = [0x200D];
        var prevWasNewLine = false;
        var graphemeStarts = StringInfo.ParseCombiningCharacters(text);
        var textSpan = text.AsSpan();

        // If we have more than 1 code unit per grapheme, the code units can
        // be distributed through multiple grapheme. In the worst case, all extra
        // code units are in exactly one grapheme -> allocate buffer of that size.
        Span<int> codePointsBuffer = stackalloc int[1 + text.Length - graphemeStarts.Length];
        for (var i = 0; i < graphemeStarts.Length; ++i)
        {
            var startIdx = graphemeStarts[i];
            var slice = textSpan.Slice(startIdx);
            if (slice.TrySliceNewLine(out var eolLen))
            {
                i += eolLen - 1;
                if (prevWasNewLine)
                {
                    // If there are consecutive new lines, we need height of new the lines between them
                    var box = engine.GetGlyphBox(zeroWidthJoiner, font, dpi);
                    output.Add(box);
                }

                output.Add(GlyphBox.LineBreak);
                prevWasNewLine = true;
            }
            else
            {
                var codeUnits = i + 1 < graphemeStarts.Length
                    ? textSpan.Slice(startIdx, graphemeStarts[i + 1] - startIdx)
                    : textSpan[startIdx..];
                var count = codeUnits.ToCodePoints(codePointsBuffer);
                ReadOnlySpan<int> grapheme = codePointsBuffer.Slice(0, count);
                var box = engine.GetGlyphBox(grapheme, font, dpi);
                output.Add(box);
                prevWasNewLine = false;
            }
        }
    }
}
