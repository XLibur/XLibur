using System;
using System.Collections.Generic;
using XLibur.Report.Ranges;

namespace XLibur.Report.Rewriting;

/// <summary>
/// Where a row or a column ends up after an expansion, and which expansions happened on which sheet.
/// </summary>
/// <remarks>
/// <para>
/// Everything that refers to a range by address rather than by identity — a chart series, a pivot
/// cache source, a pivot table's own corner — has to be moved through the same coordinate change, so
/// the arithmetic lives here once rather than in each of them.
/// </para>
/// <para>
/// An expansion moves one dimension and leaves the other alone: a range repeating downwards changes
/// rows, one repeating across changes columns. The row and column functions are therefore not
/// interchangeable — asking for the row mapping of a horizontal expansion returns the row unchanged,
/// which is the right answer and not a special case.
/// </para>
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
    /// Moves one area through one expansion, along whichever axis the expansion ran.
    /// </summary>
    /// <remarks>
    /// An area has to cross the template on the <em>other</em> axis to be stretched by it — a chart
    /// plotting a column beside a vertical range is not plotting that range — but it still moves if
    /// it sits past the template, because the insert that grew the range was a full-row or
    /// full-column one. Chart series references and pivot cache sources both follow an expansion
    /// through here, so the two cannot drift apart.
    /// </remarks>
    public static RangeArea MapArea(RangeArea area, ExpansionRecord expansion)
    {
        var template = expansion.TemplateArea;

        if (expansion.Axis.IsHorizontal)
        {
            if (area.LastRow < template.FirstRow || area.FirstRow > template.LastRow)
            {
                return new RangeArea(
                    area.FirstRow,
                    Shift(area.FirstColumn, template.LastColumn, expansion.SlotDelta),
                    area.LastRow,
                    Shift(area.LastColumn, template.LastColumn, expansion.SlotDelta));
            }

            var firstColumn = MapColumnStart(area.FirstColumn, expansion);
            var lastColumn = MapColumnEnd(area.LastColumn, expansion);

            return new RangeArea(area.FirstRow, firstColumn, area.LastRow, Math.Max(firstColumn, lastColumn));
        }

        if (area.LastColumn < template.FirstColumn || area.FirstColumn > template.LastColumn)
        {
            return new RangeArea(
                Shift(area.FirstRow, template.LastRow, expansion.SlotDelta),
                area.FirstColumn,
                Shift(area.LastRow, template.LastRow, expansion.SlotDelta),
                area.LastColumn);
        }

        var firstRow = MapRowStart(area.FirstRow, expansion);
        var lastRow = MapRowEnd(area.LastRow, expansion);

        return new RangeArea(firstRow, area.FirstColumn, Math.Max(firstRow, lastRow), area.LastColumn);
    }

    /// <summary>Where the top of something starting at <paramref name="row"/> ends up.</summary>
    public static int MapRowStart(int row, ExpansionRecord expansion) => expansion.Axis.IsHorizontal
        ? row
        : MapStart(row, expansion.TemplateArea.FirstRow, expansion.TemplateArea.LastRow,
            expansion.RenderedArea.FirstRow, expansion.RenderedArea.LastRow, expansion.SlotDelta);

    /// <summary>Where the bottom of something ending at <paramref name="row"/> ends up.</summary>
    public static int MapRowEnd(int row, ExpansionRecord expansion) => expansion.Axis.IsHorizontal
        ? row
        : MapEnd(row, expansion.TemplateArea.FirstRow, expansion.TemplateArea.LastRow,
            expansion.RenderedArea.LastRow, expansion.SlotDelta);

    /// <summary>Where the left of something starting at <paramref name="column"/> ends up.</summary>
    public static int MapColumnStart(int column, ExpansionRecord expansion) => expansion.Axis.IsHorizontal
        ? MapStart(column, expansion.TemplateArea.FirstColumn, expansion.TemplateArea.LastColumn,
            expansion.RenderedArea.FirstColumn, expansion.RenderedArea.LastColumn, expansion.SlotDelta)
        : column;

    /// <summary>Where the right of something ending at <paramref name="column"/> ends up.</summary>
    public static int MapColumnEnd(int column, ExpansionRecord expansion) => expansion.Axis.IsHorizontal
        ? MapEnd(column, expansion.TemplateArea.FirstColumn, expansion.TemplateArea.LastColumn,
            expansion.RenderedArea.LastColumn, expansion.SlotDelta)
        : column;

    /// <summary>
    /// Where one end of an area the expansion did not stretch ends up: unchanged where it sat at or
    /// before the template, moved by the delta where it sat past it.
    /// </summary>
    /// <remarks>
    /// Each end is moved on its own, because an area may straddle the template — starting above a
    /// vertical range and ending below it, in columns the range does not touch. Its start stays put
    /// and its tail follows the rows the insert pushed down; treating the area as a unit would leave
    /// the tail naming rows that have moved.
    /// </remarks>
    private static int Shift(int position, int templateLast, int delta) =>
        position > templateLast ? position + delta : position;

    /// <summary>
    /// Unchanged before the template, moved by the delta past it, and keeping its offset from the
    /// start when inside it.
    /// </summary>
    private static int MapStart(
        int position, int templateFirst, int templateLast, int renderedFirst, int renderedLast, int delta)
    {
        if (position < templateFirst)
        {
            return position;
        }

        if (position > templateLast)
        {
            return position + delta;
        }

        return Math.Min(renderedFirst + (position - templateFirst), Math.Max(renderedLast, renderedFirst));
    }

    /// <summary>
    /// As <see cref="MapStart"/>, except that anywhere inside the template ends at the end of what
    /// was generated — which is what turns a range covering the slot a template repeats into one
    /// covering every copy of it.
    /// </summary>
    private static int MapEnd(int position, int templateFirst, int templateLast, int renderedLast, int delta)
    {
        if (position < templateFirst)
        {
            return position;
        }

        if (position > templateLast)
        {
            return position + delta;
        }

        return renderedLast;
    }
}
