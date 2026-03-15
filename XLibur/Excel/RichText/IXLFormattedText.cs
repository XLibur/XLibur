using System;
using System.Collections.Generic;

namespace XLibur.Excel;

public interface IXLFormattedText<T> : IEnumerable<IXLRichString>, IEquatable<IXLFormattedText<T>>, IXLWithRichString
{
    bool Bold { set; }
    bool Italic { set; }
    XLFontUnderlineValues Underline { set; }
    bool Strikethrough { set; }
    XLFontVerticalTextAlignmentValues VerticalAlignment { set; }
    bool Shadow { set; }
    double FontSize { set; }
    XLColor FontColor { set; }
    string FontName { set; }
    XLFontFamilyNumberingValues FontFamilyNumbering { set; }

    IXLFormattedText<T> SetBold(); IXLFormattedText<T> SetBold(bool value);
    IXLFormattedText<T> SetItalic(); IXLFormattedText<T> SetItalic(bool value);
    IXLFormattedText<T> SetUnderline(); IXLFormattedText<T> SetUnderline(XLFontUnderlineValues value);
    IXLFormattedText<T> SetStrikethrough(); IXLFormattedText<T> SetStrikethrough(bool value);
    IXLFormattedText<T> SetVerticalAlignment(XLFontVerticalTextAlignmentValues value);
    IXLFormattedText<T> SetShadow(); IXLFormattedText<T> SetShadow(bool value);
    IXLFormattedText<T> SetFontSize(double value);
    IXLFormattedText<T> SetFontColor(XLColor value);
    IXLFormattedText<T> SetFontName(string value);
    IXLFormattedText<T> SetFontFamilyNumbering(XLFontFamilyNumberingValues value);

    IXLRichString AddText(string text, IXLFontBase font);
    IXLFormattedText<T> ClearText();
    IXLFormattedText<T> ClearFont();
    IXLFormattedText<T> Substring(int index);
    IXLFormattedText<T> Substring(int index, int length);

    /// <summary>
    /// Replace the text and formatting of this text by texts and formatting from the <paramref name="original"/> text.
    /// </summary>
    /// <param name="original">Original to copy from.</param>
    /// <returns>This text.</returns>
    IXLFormattedText<T> CopyFrom(IXLFormattedText<T> original);

    /// <summary>
    /// How many rich strings is the formatted text composed of.
    /// </summary>
    int Count { get; }

    /// <summary>
    /// Length of the whole formatted text.
    /// </summary>
    int Length { get; }

    /// <summary>
    /// Get text of the whole formatted text.
    /// </summary>
    string Text { get; }

    /// <summary>
    /// Does this text has phonetics? Unlike accessing the <see cref="Phonetics"/> property, this method
    /// doesn't create a new instance on access.
    /// </summary>
    bool HasPhonetics { get; }

    /// <summary>
    /// Get or create phonetics for the text. Use <see cref="HasPhonetics"/> to check for existence to avoid unnecessary creation.
    /// </summary>
    IXLPhonetics Phonetics { get; }
}
