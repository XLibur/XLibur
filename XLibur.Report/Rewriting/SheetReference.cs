using System;
using System.Text;
using XLibur.Excel;

namespace XLibur.Report.Rewriting;

/// <summary>
/// A sheet-qualified A1 reference to a cell or a rectangle (<c>Sales!$B$3:$B$3</c>).
/// </summary>
/// <remarks>
/// <para>
/// Chart references are plain strings in the model and plain strings in the file, so re-pointing a
/// series means taking one apart and putting it back together. Only the single-area form is
/// understood — a multi-area reference, a 3-D reference across sheets, or a whole-row or
/// whole-column one is left alone rather than guessed at, because a wrong reference is worse for a
/// report than a stale one.
/// </para>
/// <para>
/// The same reading serves anywhere a template writes a reference as text rather than as a range —
/// <c>&lt;&lt;Pivot dest="Summary!A3"&gt;&gt;</c> among them.
/// </para>
/// </remarks>
internal readonly record struct SheetReference(
    string? SheetName,
    int FirstRow,
    int FirstColumn,
    int LastRow,
    int LastColumn)
{
    /// <summary>Reads <paramref name="text"/>, reporting whether it is a form worth rewriting.</summary>
    public static bool TryParse(string? text, out SheetReference reference)
    {
        reference = default;

        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var span = text.Trim();

        // A multi-area reference is written as a parenthesised, comma-separated list.
        if (span.Contains(',') || span.Contains('('))
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

            reference = new SheetReference(sheetName, row, column, row, column);
            return true;
        }

        if (!TryParseCell(area[..colon], out var firstRow, out var firstColumn)
            || !TryParseCell(area[(colon + 1)..], out var lastRow, out var lastColumn))
        {
            return false;
        }

        reference = new SheetReference(
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
            return TrySplitQuotedSheet(text, ref sheetName, ref area);
        }

        var separator = text.IndexOf('!');
        if (separator < 0)
        {
            return true;
        }

        var unquoted = text[..separator];

        // A 3-D reference spans sheets; there is no single sheet to match it against.
        if (unquoted.Contains(':'))
        {
            return false;
        }

        sheetName = unquoted;
        area = text[(separator + 1)..];
        return area.Length > 0;
    }

    /// <summary>
    /// Splits a sheet name written in quotes (<c>'My Sheet'!A1</c>) from its area. Leaves
    /// <paramref name="sheetName"/> and <paramref name="area"/> untouched when the quoting is malformed.
    /// </summary>
    private static bool TrySplitQuotedSheet(string text, ref string? sheetName, ref string area)
    {
        var name = ReadQuotedName(text, out var i);

        if (i >= text.Length - 1 || text[i] != '\'' || text[i + 1] != '!')
        {
            return false;
        }

        sheetName = name;
        area = text[(i + 2)..];
        return area.Length > 0;
    }

    /// <summary>
    /// Reads a quoted sheet name starting after the opening quote, stopping at the closing quote
    /// (whose index is returned in <paramref name="end"/>) or the end of the text.
    /// </summary>
    private static string ReadQuotedName(string text, out int end)
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

        end = i;
        return name.ToString();
    }

    /// <summary>Parses <c>$B$3</c>, <c>B3</c> and the mixed forms into row and column numbers.</summary>
    private static bool TryParseCell(string text, out int row, out int column)
    {
        row = 0;
        column = 0;

        var i = SkipDollar(text, 0);

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

        i = SkipDollar(text, i);

        var digitStart = i;
        i = SkipDigits(text, i);

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

    /// <summary>Steps past an absolute-reference marker at <paramref name="i"/>, if there is one.</summary>
    private static int SkipDollar(string text, int i) => i < text.Length && text[i] == '$' ? i + 1 : i;

    private static int SkipDigits(string text, int i)
    {
        while (i < text.Length && char.IsAsciiDigit(text[i]))
        {
            i++;
        }

        return i;
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
