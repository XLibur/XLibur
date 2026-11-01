using System;
using System.Text;
using XLibur.Excel.RichText;
using XLibur.Extensions;

namespace XLibur.Excel;

internal sealed class XLHFText
{
    private readonly XLHFItem _hfItem;

    public XLHFText(XLRichString richText, XLHFItem hfItem)
    {
        RichText = richText;
        _hfItem = hfItem;
    }

    public XLRichString RichText { get; private set; }

    public string GetHFText(string prevText)
    {
        var wsFont = _hfItem.HeaderFooter.Worksheet.Style.Font;

        var isRichText = DiffersFromFont(wsFont);

        if (!isRichText)
            return RichText.Text;

        StringBuilder sb = new StringBuilder();

        AppendFontNameAndStyle(sb, wsFont);

        if (HasOwnFontSize(wsFont))
            sb.Append("&" + RichText.FontSize);

        AppendStrikethrough(sb, wsFont);
        AppendVerticalAlignment(sb, wsFont);
        AppendUnderline(sb, wsFont);
        AppendFontColor(sb, prevText, wsFont);

        sb.Append(RichText.Text);

        AppendUnderline(sb, wsFont);
        AppendVerticalAlignment(sb, wsFont);
        AppendStrikethrough(sb, wsFont);

        return sb.ToString();
    }

    /// <summary>
    /// Whether the text states any formatting of its own over the worksheet font, and so needs
    /// header/footer formatting codes.
    /// </summary>
    private bool DiffersFromFont(IXLFontBase wsFont)
    {
        return HasOwnFontName(wsFont)
               || RichText.Bold != wsFont.Bold
               || RichText.Italic != wsFont.Italic
               || RichText.Strikethrough != wsFont.Strikethrough
               || HasOwnFontSize(wsFont)
               || RichText.VerticalAlignment != wsFont.VerticalAlignment
               || RichText.Underline != wsFont.Underline
               || !RichText.FontColor.Equals(wsFont.FontColor);
    }

    private bool HasOwnFontName(IXLFontBase wsFont)
    {
        return RichText.FontName != null && RichText.FontName != wsFont.FontName;
    }

    private bool HasOwnFontSize(IXLFontBase wsFont)
    {
        return RichText.FontSize > 0 && Math.Abs(RichText.FontSize - wsFont.FontSize) > XLHelper.Epsilon;
    }

    private void AppendStrikethrough(StringBuilder sb, IXLFontBase wsFont)
    {
        if (RichText.Strikethrough && !wsFont.Strikethrough)
            sb.Append("&S");
    }

    private void AppendFontNameAndStyle(StringBuilder sb, IXLFontBase wsFont)
    {
        if (HasOwnFontName(wsFont))
            sb.Append("&\"" + RichText.FontName);
        else
            sb.Append("&\"-");

        if (RichText.Bold && RichText.Italic)
            sb.Append(",Bold Italic\"");
        else if (RichText.Bold)
            sb.Append(",Bold\"");
        else if (RichText.Italic)
            sb.Append(",Italic\"");
        else
            sb.Append(",Regular\"");
    }

    private void AppendVerticalAlignment(StringBuilder sb, IXLFontBase wsFont)
    {
        if (RichText.VerticalAlignment != wsFont.VerticalAlignment)
        {
            if (RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Subscript)
                sb.Append("&Y");
            else if (RichText.VerticalAlignment == XLFontVerticalTextAlignmentValues.Superscript)
                sb.Append("&X");
        }
    }

    private void AppendUnderline(StringBuilder sb, IXLFontBase wsFont)
    {
        if (RichText.Underline != wsFont.Underline)
        {
            if (RichText.Underline == XLFontUnderlineValues.Single)
                sb.Append("&U");
            else if (RichText.Underline == XLFontUnderlineValues.Double)
                sb.Append("&E");
        }
    }

    private void AppendFontColor(StringBuilder sb, string prevText, IXLFontBase wsFont)
    {
        var lastColorPosition = prevText.LastIndexOf("&K", StringComparison.Ordinal);

        var hasPrevColor = lastColorPosition >= 0 && prevText.Length >= lastColorPosition + 8;
        XLColor? prevColor = null;
        if (hasPrevColor)
        {
            try
            {
                prevColor = XLColor.FromHtml(string.Concat("#", prevText.AsSpan(lastColorPosition + 2, 6)));
            }
            catch (FormatException)
            {
                hasPrevColor = false;
            }
        }

        if (
            (hasPrevColor && !RichText.FontColor.Equals(prevColor))
            || (!hasPrevColor && !RichText.FontColor.Equals(wsFont.FontColor))
        )
            sb.Append("&K").Append(RichText.FontColor.Color.ToHex().AsSpan(2));
    }
}
