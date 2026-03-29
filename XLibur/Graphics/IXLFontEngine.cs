using System;
using XLibur.Excel;

namespace XLibur.Graphics;

/// <summary>
/// An interface to abstract away font-related functionality, like measuring text width, height, and glyph metrics.
/// This allows swapping the underlying font library (e.g., SixLabors.Fonts) with an alternative implementation.
/// </summary>
public interface IXLFontEngine
{
    /// <summary>
    /// Get the height of a text with the font in pixels. Should be <c>EMHeight+descent</c>.
    /// </summary>
    double GetTextHeight(IXLFontBase font, double dpiY);

    /// <summary>
    /// Get the width of a text in pixels. Do not add any padding, there can be
    /// multiple spans of a text with different fonts in a line.
    /// </summary>
    double GetTextWidth(string text, IXLFontBase font, double dpiX);

    /// <summary>
    /// The width of the widest 0-9 digit in pixels.
    /// </summary>
    /// <remarks>OOXML measures width of a column in multiples of widest 0-9 digit character in a normal style font.</remarks>
    double GetMaxDigitWidth(IXLFontBase font, double dpiX);

    /// <summary>
    /// Get font descent in pixels (positive value).
    /// </summary>
    /// <remarks>Excel is using OS/2 WinAscent/WinDescent for TrueType fonts (e.g. Calibri), not a correct font ascent/descent.</remarks>
    double GetDescent(IXLFontBase font, double dpiY);

    /// <summary>
    /// Get a glyph-bounding box for a grapheme cluster.
    /// </summary>
    /// <remarks>
    /// In 99+%, grapheme cluster will be just a codepoint. Method uses grapheme instead, so it can be
    /// a future-proof signature and have less braking changes. Implementing method by adding widths of
    /// individual code points is acceptable.
    /// </remarks>
    /// <param name="graphemeCluster">
    /// A part of a string in code points (or runes in C# terminology, not UTF-16 code units) that together
    /// form a grapheme. Multiple Unicode codepoints can form a single glyph, e.g. family grapheme is a single
    /// glyph created from 6 codepoints (man, zero-width-join, woman, zero-width-join and a girl). A string
    /// can be split into a grapheme clusters through <see cref="System.Globalization.StringInfo.GetTextElementEnumerator(string)"/>.
    /// </param>
    /// <param name="font">Font used to determine the size of a glyph for the grapheme cluster.</param>
    /// <param name="dpi">
    /// A resolution used to determine the pixel size of a glyph. Font might be rendered differently at different resolutions.
    /// </param>
    /// <returns>Bounding box containing the glyph.</returns>
    GlyphBox GetGlyphBox(ReadOnlySpan<int> graphemeCluster, IXLFontBase font, Dpi dpi);
}
