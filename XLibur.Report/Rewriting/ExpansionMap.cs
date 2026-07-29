using System;
using System.Collections.Generic;
using XLibur.Report.Ranges;

namespace XLibur.Report.Rewriting;

/// <summary>
/// Where a row ends up after an expansion, and which expansions happened on which sheet.
/// </summary>
/// <remarks>
/// Everything that refers to a range by address rather than by identity — a chart series, a pivot
/// cache source, a pivot table's own corner — has to be moved through the same coordinate change,
/// so the arithmetic lives here once rather than in each of them.
/// </remarks>
internal static class ExpansionMap
{
    /// <summary>Groups a ledger by the sheet each expansion happened on.</summary>
    /// <remarks>
    /// Keyed by name and compared case-insensitively, because that is how a reference names a sheet
    /// and how Excel compares the two.
    /// </remarks>
    public static Dictionary<string, List<ExpansionRecord>> BySheet(IReadOnlyList<ExpansionRecord> expansions)
    {
        var bySheet = new Dictionary<string, List<ExpansionRecord>>(StringComparer.OrdinalIgnoreCase);

        foreach (var expansion in expansions)
        {
            if (!bySheet.TryGetValue(expansion.Worksheet.Name, out var list))
            {
                list = new List<ExpansionRecord>();
                bySheet[expansion.Worksheet.Name] = list;
            }

            list.Add(expansion);
        }

        return bySheet;
    }

    /// <summary>
    /// Where the top of something anchored at <paramref name="row"/> ends up: unchanged above the
    /// template, moved by the delta below it, and keeping its offset from the top when inside it.
    /// </summary>
    public static int MapStart(int row, ExpansionRecord expansion)
    {
        var template = expansion.TemplateArea;
        var rendered = expansion.RenderedArea;

        if (row < template.FirstRow)
        {
            return row;
        }

        if (row > template.LastRow)
        {
            return row + expansion.RowDelta;
        }

        return Math.Min(
            rendered.FirstRow + (row - template.FirstRow),
            Math.Max(rendered.LastRow, rendered.FirstRow));
    }

    /// <summary>
    /// Where the bottom of something ending at <paramref name="row"/> ends up. A row anywhere inside
    /// the template goes to the bottom of what was generated, which is what turns a range covering
    /// the row a template repeats into one covering every copy of it.
    /// </summary>
    public static int MapEnd(int row, ExpansionRecord expansion)
    {
        var template = expansion.TemplateArea;

        if (row < template.FirstRow)
        {
            return row;
        }

        if (row > template.LastRow)
        {
            return row + expansion.RowDelta;
        }

        return expansion.RenderedArea.LastRow;
    }
}
