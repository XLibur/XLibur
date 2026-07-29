using System;
using System.Text;
using XLibur.Excel;

namespace XLibur.Report.Rewriting;

/// <summary>
/// A sheet-qualified A1 range reference, in the form a chart series holds one
/// (<c>Sales!$B$3:$B$3</c>).
/// </summary>
/// <remarks>
/// Chart references are plain strings in the model and plain strings in the file, so re-pointing a
/// series means taking one apart and putting it back together. Only the single-area form is
/// understood — a multi-area reference, a 3-D reference across sheets, or a whole-row or
/// whole-column one is left alone rather than guessed at, because a wrong reference is worse for a
/// report than a stale one.
/// </remarks>
internal readonly record struct SeriesReference(
    string? SheetName,
    int FirstRow,
    int FirstColumn,
    int LastRow,
    int LastColumn)
{
    /// <summary>Reads <paramref name="text"/>, reporting whether it is a form worth rewriting.</summary>
    public static bool TryParse(string? text, out SeriesReference reference)
    {
        reference = default;

        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var span = text!.Trim();

        // A multi-area reference is written as a parenthesised, comma-separated list.
        if (span.IndexOf(',') >= 0 || span.IndexOf('(') >= 0)
        {
            return false;
        }

        if (!TrySplitSheet(span, out var sheetName, out var area))
        {
            return false;
        }

        var colon = area.IndexOf(':');
        if (colon < 0)
        {
            if (!TryParseCell(area, out var row, out var column))
            {
                return false;
            }

            reference = new SeriesReference(sheetName, row, column, row, column);
            return true;
        }

        if (!TryParseCell(area[..colon], out var firstRow, out var firstColumn)
            || !TryParseCell(area[(colon + 1)..], out var lastRow, out var lastColumn))
        {
            return false;
        }

        reference = new SeriesReference(
            sheetName,
            Math.Min(firstRow, lastRow),
            Math.Min(firstColumn, lastColumn),
            Math.Max(firstRow, lastRow),
            Math.Max(firstColumn, lastColumn));

        return true;
    }

    /// <summary>Writes the reference back out, fully absolute, the way Excel stores one.</summary>
    public string ToText()
    {
        var text = new StringBuilder();

        if (SheetName is { Length: > 0 } sheet)
        {
            text.Append(QuoteSheetName(sheet)).Append('!');
        }

        text.Append('$').Append(XLHelper.GetColumnLetterFromNumber(FirstColumn)).Append('$').Append(FirstRow);

        if (FirstRow != LastRow || FirstColumn != LastColumn)
        {
            text.Append(":$").Append(XLHelper.GetColumnLetterFromNumber(LastColumn)).Append('$').Append(LastRow);
        }

        return text.ToString();
    }

    /// <summary>
    /// Separates the sheet name from the area, understanding the quoting Excel uses for a name that
    /// contains a space or a punctuation mark.
    /// </summary>
    private static bool TrySplitSheet(string text, out string? sheetName, out string area)
    {
        sheetName = null;
        area = text;

        if (text.Length == 0)
        {
            return false;
        }

        if (text[0] == '\'')
        {
            var name = new StringBuilder();
            var i = 1;

            while (i < text.Length)
            {
                if (text[i] == '\'')
                {
                    // A doubled quote is a literal one inside the name.
                    if (i + 1 < text.Length && text[i + 1] == '\'')
                    {
                        name.Append('\'');
                        i += 2;
                        continue;
                    }

                    break;
                }

                name.Append(text[i]);
                i++;
            }

            if (i >= text.Length - 1 || text[i] != '\'' || text[i + 1] != '!')
            {
                return false;
            }

            sheetName = name.ToString();
            area = text[(i + 2)..];
            return area.Length > 0;
        }

        var separator = text.IndexOf('!');
        if (separator < 0)
        {
            return true;
        }

        var unquoted = text[..separator];

        // A 3-D reference spans sheets; there is no single sheet to match it against.
        if (unquoted.IndexOf(':') >= 0)
        {
            return false;
        }

        sheetName = unquoted;
        area = text[(separator + 1)..];
        return area.Length > 0;
    }

    /// <summary>Parses <c>$B$3</c>, <c>B3</c> and the mixed forms into row and column numbers.</summary>
    private static bool TryParseCell(string text, out int row, out int column)
    {
        row = 0;
        column = 0;

        var i = 0;
        if (i < text.Length && text[i] == '$')
        {
            i++;
        }

        var letterStart = i;
        while (i < text.Length && char.IsAsciiLetter(text[i]))
        {
            i++;
        }

        // No letters, or no digits: a whole-row or whole-column reference, which has no fixed
        // extent to stretch.
        if (i == letterStart)
        {
            return false;
        }

        var letters = text[letterStart..i];

        if (i < text.Length && text[i] == '$')
        {
            i++;
        }

        var digitStart = i;
        while (i < text.Length && char.IsAsciiDigit(text[i]))
        {
            i++;
        }

        if (i != text.Length || i == digitStart)
        {
            return false;
        }

        if (!int.TryParse(text[digitStart..i], out row) || row < 1)
        {
            return false;
        }

        try
        {
            column = XLHelper.GetColumnNumberFromLetter(letters);
        }
        catch (ArgumentException)
        {
            return false;
        }

        return column >= 1;
    }

    private static string QuoteSheetName(string name)
    {
        var needsQuotes = false;

        for (var i = 0; i < name.Length; i++)
        {
            var c = name[i];
            if (!char.IsAsciiLetterOrDigit(c) && c != '_' && c != '.')
            {
                needsQuotes = true;
                break;
            }
        }

        needsQuotes |= name.Length > 0 && char.IsAsciiDigit(name[0]);

        return needsQuotes ? "'" + name.Replace("'", "''") + "'" : name;
    }
}
