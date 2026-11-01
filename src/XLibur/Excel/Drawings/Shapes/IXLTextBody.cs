using System.Collections.Generic;

namespace XLibur.Excel.Drawings;

/// <summary>The text of a shape, <c>xdr:txBody</c>: a list of paragraphs, each a list of runs.</summary>
public interface IXLTextBody : IEnumerable<IXLTextParagraph>
{
    /// <summary>The number of paragraphs.</summary>
    int Count { get; }

    IXLTextParagraph this[int index] { get; }

    /// <summary>Adds a paragraph at the end, holding <paramref name="text"/> in one run, or no run when it is empty.</summary>
    IXLTextParagraph AddParagraph(string text = "");

    /// <summary>Removes every paragraph. This rebuilds the body; the formatting of its runs goes with them.</summary>
    void Clear();

    /// <summary>
    /// The whole text, paragraphs joined by <c>\n</c>. Setting it replaces every paragraph; each new
    /// paragraph takes the properties of the first paragraph the body had, and each new run the
    /// formatting of its first run. A control character other than a tab or <c>\n</c> throws
    /// <see cref="System.ArgumentException"/>.
    /// </summary>
    string Text { get; set; }

    /// <summary>Where the text sits between the top and bottom of the shape, <c>a:bodyPr/@anchor</c>.</summary>
    XLTextVerticalAlignment VerticalAlignment { get; set; }

    /// <summary>Whether lines wrap at the shape's width, <c>a:bodyPr/@wrap</c>.</summary>
    bool WrapText { get; set; }
}

/// <summary>A paragraph of shape text, <c>a:p</c>.</summary>
public interface IXLTextParagraph : IEnumerable<IXLTextRun>
{
    /// <summary>The horizontal alignment, <c>a:pPr/@algn</c>.</summary>
    XLTextAlignment Alignment { get; set; }

    /// <summary>Adds a run holding <paramref name="text"/> at the end, and returns it so its formatting can be chained.</summary>
    IXLTextRun AddText(string text);

    /// <summary>
    /// The text of every run, in order. A line break, <c>a:br</c>, reads as <c>\v</c>, which cannot be
    /// written back: start a new paragraph instead.
    /// </summary>
    string Text { get; }

    /// <summary>The number of runs, <c>a:r</c>.</summary>
    int Count { get; }

    IXLTextRun this[int index] { get; }
}

/// <summary>
/// The font of a run of shape text, <c>a:rPr</c>. Deliberately narrower than
/// <see cref="IXLFontBase"/>: each member maps onto <c>a:rPr</c>, which has no shadow flag, no
/// accounting underlines and no font family, character set or scheme. A <c>null</c> property was not
/// assigned, and the run keeps what the file or the shape's style says.
/// </summary>
public interface IXLShapeTextFont
{
    /// <summary><c>a:rPr/@b</c></summary>
    bool? Bold { get; set; }

    /// <summary><c>a:rPr/@i</c></summary>
    bool? Italic { get; set; }

    /// <summary><c>a:rPr/@u</c></summary>
    XLShapeTextUnderline? Underline { get; set; }

    /// <summary><c>a:rPr/@strike</c></summary>
    bool? Strikethrough { get; set; }

    /// <summary>The size in points, <c>a:rPr/@sz</c>.</summary>
    double? FontSize { get; set; }

    /// <summary><c>a:rPr/a:solidFill</c></summary>
    XLColor? FontColor { get; set; }

    /// <summary><c>a:rPr/a:latin/@typeface</c></summary>
    string? FontName { get; set; }
}

/// <summary>A run of shape text, <c>a:r</c>: text in one font.</summary>
public interface IXLTextRun : IXLShapeTextFont
{
    /// <summary>The text. Setting it rewrites the text only; the run's formatting is kept.</summary>
    string Text { get; set; }

    IXLTextRun SetBold(bool value = true);

    IXLTextRun SetItalic(bool value = true);

    IXLTextRun SetFontSize(double value);

    IXLTextRun SetFontColor(XLColor value);

    IXLTextRun SetFontName(string value);
}
