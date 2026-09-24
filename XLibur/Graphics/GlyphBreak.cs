namespace XLibur.Graphics;

/// <summary>
/// How a glyph takes part in soft line wrapping. Kept alongside a list of <see cref="GlyphBox"/>
/// so that the public glyph box doesn't have to carry the code point it measures.
/// </summary>
internal enum GlyphBreak : byte
{
    /// <summary>A glyph that belongs to a word. A line can only break inside it when the word doesn't fit on a line of its own.</summary>
    None,

    /// <summary>A space. A line can break after it, and a space at the end of a line doesn't count towards its width.</summary>
    Space,

    /// <summary>A hyphen. A line can break after it, and the hyphen stays on the first line.</summary>
    BreakAfter,
}
