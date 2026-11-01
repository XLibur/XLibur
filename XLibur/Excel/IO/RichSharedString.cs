using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.RichText;

namespace XLibur.Excel.IO;

/// <summary>
/// A rich-text entry of the file's shared string table, and the rich text a load built from it for
/// each cell font it was shown in.
/// </summary>
/// <remarks>
/// <para>
/// A rich string's runs take whatever their <c>rPr</c> leaves unstated from the font of the cell
/// that shows them, so one string reads as one rich text per cell font. The first cell in a font
/// builds it through <see cref="WorksheetSheetDataReader.SetCellText"/>, as every cell used to;
/// each later cell in that font takes the built text with one write. Building it again for each
/// cell applied every run's font one property at a time, and each property interned the cell's
/// whole rich text in the shared string table again: some 26 kB a cell.
/// </para>
/// <para>
/// The built text is immutable and interned by value, so cells that take it share one entry of the
/// table, as cells that built equal texts did. A load reads sheets one after another, so this needs
/// no lock.
/// </para>
/// </remarks>
internal sealed class RichSharedString(RstType element)
{
    // Most strings are shown in one cell font, which takes no dictionary.
    private XLFontKey _firstFont;
    private XLImmutableRichText? _first;
    private Dictionary<XLFontKey, XLImmutableRichText>? _others;

    /// <summary>The <c>&lt;si&gt;</c> element, which the first cell in each font builds from.</summary>
    internal RstType Element { get; } = element;

    /// <summary>The rich text a cell in <paramref name="font"/> was given, if one was.</summary>
    internal bool TryGetBuilt(in XLFontKey font, out XLImmutableRichText richText)
    {
        if (_first is not null && _firstFont == font)
        {
            richText = _first;
            return true;
        }

        if (_others is not null && _others.TryGetValue(font, out var other))
        {
            richText = other;
            return true;
        }

        richText = null!;
        return false;
    }

    /// <summary>Records the rich text a cell in <paramref name="font"/> was built with.</summary>
    internal void AddBuilt(in XLFontKey font, XLImmutableRichText richText)
    {
        if (_first is null)
        {
            _firstFont = font;
            _first = richText;
            return;
        }

        (_others ??= [])[font] = richText;
    }
}
