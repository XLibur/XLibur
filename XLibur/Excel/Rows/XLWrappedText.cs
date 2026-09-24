using System;
using System.Collections.Generic;
using XLibur.Graphics;

namespace XLibur.Excel;

/// <summary>
/// Measures the height of horizontal text that is soft-wrapped to fit a cell, the way Excel
/// lays out a cell with <see cref="IXLAlignment.WrapText"/> set.
/// </summary>
/// <remarks>
/// Hard line breaks start a new paragraph. Inside a paragraph, a line breaks after a space or a
/// hyphen. A word that doesn't fit on a line of its own is split between glyphs. Spaces at the
/// end of a line don't count towards its width.
/// </remarks>
internal static class XLWrappedText
{
    /// <summary>
    /// Get the height of the text in pixels.
    /// </summary>
    /// <param name="glyphs">Glyphs of the text, as produced by <see cref="XLCellGlyphHelper"/>.</param>
    /// <param name="breaks">Break kind of each glyph in <paramref name="glyphs"/>.</param>
    /// <param name="availableWidthPx">Width of the cell in pixels, including padding and the grid line.</param>
    /// <param name="scaledMdw">Maximum digit width of the cell font, used for cell padding.</param>
    internal static double GetHeight(List<GlyphBox> glyphs, List<GlyphBreak> breaks, int availableWidthPx, double scaledMdw)
    {
        var layout = new Layout(glyphs, breaks, availableWidthPx, scaledMdw);
        var textHeight = 0d;
        var paragraphStart = 0;
        for (var i = 0; i <= glyphs.Count; i++)
        {
            if (i == glyphs.Count || glyphs[i].IsLineBreak)
            {
                textHeight += layout.GetParagraphHeight(paragraphStart, i);
                paragraphStart = i + 1;
            }
        }

        return textHeight;
    }

    private readonly struct Layout(List<GlyphBox> glyphs, List<GlyphBreak> breaks, int availableWidthPx, double scaledMdw)
    {
        private bool Fits(double lineWidthPx)
        {
            var textWidthPx = (int)Math.Ceiling(lineWidthPx);
            return XLColumn.GetCellWidthPx(textWidthPx, scaledMdw) <= availableWidthPx;
        }

        /// <summary>
        /// Height of glyphs <c>[start, end)</c>, a paragraph without hard line breaks.
        /// </summary>
        internal double GetParagraphHeight(int start, int end)
        {
            var height = 0d;
            var line = new Line();
            var i = start;
            while (i < end)
            {
                // A word runs up to a space, or up to and including a hyphen.
                var wordStart = i;
                var wordWidth = 0d;
                var wordHeight = 0d;
                while (i < end && breaks[i] != GlyphBreak.Space)
                {
                    wordWidth += glyphs[i].AdvanceWidth;
                    wordHeight = Math.Max(wordHeight, glyphs[i].LineHeight);
                    i++;
                    if (breaks[i - 1] == GlyphBreak.BreakAfter)
                        break;
                }

                var wordEnd = i;

                // The spaces after a word stay on its line, but only count towards the width
                // when another word follows on the same line.
                var spaceWidth = 0d;
                var spaceHeight = 0d;
                while (i < end && breaks[i] == GlyphBreak.Space)
                {
                    spaceWidth += glyphs[i].AdvanceWidth;
                    spaceHeight = Math.Max(spaceHeight, glyphs[i].LineHeight);
                    i++;
                }

                if (wordStart == wordEnd)
                {
                    // Spaces at the start of a paragraph are indentation, so they take up width.
                    line.Width = spaceWidth;
                    line.Height = spaceHeight;
                    line.HasContent = true;
                    continue;
                }

                if (line.HasContent && Fits(line.Width + line.PendingSpaceWidth + wordWidth))
                {
                    line.Width += line.PendingSpaceWidth + wordWidth;
                    line.Height = Math.Max(line.Height, wordHeight);
                }
                else
                {
                    if (line.HasContent)
                        height += line.Break();

                    if (Fits(wordWidth))
                    {
                        line.Width = wordWidth;
                        line.Height = wordHeight;
                    }
                    else
                    {
                        height += SplitWord(ref line, wordStart, wordEnd);
                    }

                    line.HasContent = true;
                }

                line.PendingSpaceWidth = spaceWidth;
                line.Height = Math.Max(line.Height, spaceHeight);
            }

            if (line.HasContent)
                height += line.Height;

            return height;
        }

        /// <summary>
        /// Lay out glyphs <c>[start, end)</c> of a word that doesn't fit on a line, starting on an
        /// empty line and breaking between glyphs. Returns the height of the lines it completed.
        /// </summary>
        private double SplitWord(ref Line line, int start, int end)
        {
            var height = 0d;
            for (var i = start; i < end; i++)
            {
                var glyph = glyphs[i];

                // Every line gets at least one glyph, even when the cell is narrower than it.
                if (line.Width > 0 && !Fits(line.Width + glyph.AdvanceWidth))
                    height += line.Break();

                line.Width += glyph.AdvanceWidth;
                line.Height = Math.Max(line.Height, glyph.LineHeight);
            }

            return height;
        }
    }

    private struct Line
    {
        internal bool HasContent;
        internal double Width;
        internal double Height;
        internal double PendingSpaceWidth;

        /// <summary>
        /// End the line and start an empty one. Returns the height of the ended line.
        /// </summary>
        internal double Break()
        {
            var height = Height;
            this = default;
            return height;
        }
    }
}
