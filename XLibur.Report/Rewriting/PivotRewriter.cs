using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using XLibur.Excel.Exceptions;
using XLibur.Report.Ranges;

namespace XLibur.Report.Rewriting;

/// <summary>
/// Points the pivot tables in a generated workbook at what the report actually produced.
/// </summary>
/// <remarks>
/// <para>
/// Two things need doing, and only one of them was expected. A cache whose source is a plain cell
/// area does not grow with the rows expansion inserted, because the source is a sheet-plus-rectangle
/// value rather than a live range — so it is re-pointed. A cache sourced from a defined name or a
/// table does grow on its own, the name or table being tracked, but it still holds the records it
/// read at load: it needs the refresh, not the re-point. Losing that refresh is what upstream
/// ClosedXML.Report 0.2.12 regressed in its issue #399, and re-pointing without one is half of the
/// corrupt output in #200.
/// </para>
/// <para>
/// The unexpected one: a pivot table's own position is a plain rectangle too, so a pivot below a
/// bound range stays where the template put it while the rows it sat below multiply underneath it.
/// It is moved with them.
/// </para>
/// <para>
/// Every cache is also marked to refresh on open. XLibur recomputes the cached records here, but
/// Excel is the authority on its own pivot layout, and asking it to re-read on open is what makes a
/// generated pivot agree with the data beneath it however the fields are arranged.
/// </para>
/// <para>
/// Only what the <em>template</em> held is here to rewrite. A pivot that <c>&lt;&lt;Pivot&gt;&gt;</c>
/// asks for is built after this runs, precisely so that it is not re-pointed off the rows it was built
/// over — see <see cref="PivotBuilder"/>.
/// </para>
/// </remarks>
internal static class PivotRewriter
{
    /// <summary>
    /// Re-points, refreshes and repositions everything pivot-related that
    /// <paramref name="expansions"/> affected.
    /// </summary>
    public static void Rewrite(IXLWorkbook workbook, IReadOnlyList<ExpansionRecord> expansions, TemplateErrors errors)
    {
        if (expansions.Count == 0)
        {
            return;
        }

        var bySheet = ExpansionMap.BySheet(expansions);

        RewriteCaches(workbook, bySheet, errors);
        MovePivotTables(workbook, bySheet);
    }

    private static void RewriteCaches(
        IXLWorkbook workbook,
        Dictionary<string, List<ExpansionRecord>> bySheet,
        TemplateErrors errors)
    {
        foreach (var cache in workbook.PivotCaches)
        {
            var sourceRange = cache.SourceKind == XLPivotSourceKind.Range ? cache.SourceRange : null;
            string sheetName;

            switch (cache.SourceKind)
            {
                case XLPivotSourceKind.Range when sourceRange is not null:
                    sheetName = sourceRange.Worksheet.Name;
                    break;

                case XLPivotSourceKind.Range:
                    // The area names a sheet that is not in the workbook, so nothing was generated
                    // on it either. Left alone, and not worth reporting: the template was already
                    // like this before the report ran.
                    continue;

                case XLPivotSourceKind.Name when cache.SourceWorksheet is { } named:
                    sheetName = named.Name;
                    break;

                case XLPivotSourceKind.Name:
                    // A named source that resolves to nothing: most likely a name the report deleted
                    // along with a range that turned out to be empty. The pivot is in the output with
                    // no data behind it, which the report's author needs telling about.
                    errors.Add(SourceUnreadable(cache.SourceName, null));
                    continue;

                default:
                    // A source XLibur cannot read — an external workbook, a consolidation, a
                    // scenario, a connection — is left exactly as the template had it, silently.
                    // There is nothing wrong with it, and nothing this can do for it.
                    continue;
            }

            if (!bySheet.TryGetValue(sheetName, out var expansions))
            {
                // Nothing was generated on the sheet this cache reads, so its records are still
                // right and its source still points where the template meant.
                continue;
            }

            if (sourceRange is not null)
            {
                cache.SetSourceRange(RePoint(sourceRange, expansions));
            }

            Refresh(cache, sheetName, errors);
        }
    }

    /// <summary>
    /// Moves and stretches an area source through the expansions on its sheet, in the order they
    /// ran.
    /// </summary>
    private static IXLRange RePoint(IXLRange source, List<ExpansionRecord> expansions)
    {
        var worksheet = source.Worksheet;
        var topRow = source.RangeAddress.FirstAddress.RowNumber;
        var leftColumn = source.RangeAddress.FirstAddress.ColumnNumber;
        var bottomRow = source.RangeAddress.LastAddress.RowNumber;
        var rightColumn = source.RangeAddress.LastAddress.ColumnNumber;

        foreach (var expansion in expansions)
        {
            var template = expansion.TemplateArea;

            // The area has to cross the template on the other axis to be stretched by the expansion,
            // but a full-row or full-column insert moves everything past it either way.
            if (expansion.Axis.IsHorizontal)
            {
                var crosses = bottomRow >= template.FirstRow && topRow <= template.LastRow;

                if (!crosses)
                {
                    leftColumn = Shift(leftColumn, template.LastColumn, expansion.SlotDelta);
                    rightColumn = Shift(rightColumn, template.LastColumn, expansion.SlotDelta);

                    continue;
                }

                leftColumn = ExpansionMap.MapColumnStart(leftColumn, expansion);
                rightColumn = Math.Max(leftColumn, ExpansionMap.MapColumnEnd(rightColumn, expansion));
                continue;
            }

            var overlaps = rightColumn >= template.FirstColumn && leftColumn <= template.LastColumn;

            if (!overlaps)
            {
                topRow = Shift(topRow, template.LastRow, expansion.SlotDelta);
                bottomRow = Shift(bottomRow, template.LastRow, expansion.SlotDelta);

                continue;
            }

            topRow = ExpansionMap.MapRowStart(topRow, expansion);
            bottomRow = Math.Max(topRow, ExpansionMap.MapRowEnd(bottomRow, expansion));
        }

        return worksheet.Range(topRow, leftColumn, bottomRow, rightColumn);
    }

    /// <summary>
    /// Where one edge of a source the expansion did not stretch ends up: unchanged where it sat at or
    /// before the template, moved by the delta where it sat past it. Each edge moves on its own, so a
    /// source straddling the template keeps its start and follows the rows its tail sat above.
    /// </summary>
    private static int Shift(int position, int templateLast, int delta) =>
        position > templateLast ? position + delta : position;

    /// <summary>
    /// Re-reads the cache's records, and asks Excel to do the same when it opens the file.
    /// </summary>
    private static void Refresh(IXLPivotCache cache, string sheetName, TemplateErrors errors)
    {
        try
        {
            cache.Refresh();
        }
        catch (Exception ex) when (ex is InvalidReferenceException or ArgumentException)
        {
            errors.Add(SourceUnreadable(null, sheetName, ex));
        }

        cache.SetRefreshDataOnOpen();
    }

    /// <summary>
    /// The one thing that can go wrong here, said the same way however it was discovered: a pivot in
    /// the generated report whose source cannot be read. Reported rather than thrown, because a
    /// finished report with one stale pivot is worth more than no report at all.
    /// </summary>
    private static TemplateError SourceUnreadable(string? sourceName, string? sheetName, Exception? exception = null)
    {
        var what = sourceName is { Length: > 0 } name ? $"'{name}'" : "the range it is built on";

        return new TemplateError(
            $"A pivot table's source data could not be read after generation ({what}), so its cache "
            + "still holds what the template was saved with. Check that the range or name it is "
            + "built on survives generation.",
            sheetName,
            exception: exception);
    }

    /// <summary>
    /// Moves a pivot table that sat below rows the report generated, so that it is not written over
    /// by them.
    /// </summary>
    private static void MovePivotTables(IXLWorkbook workbook, Dictionary<string, List<ExpansionRecord>> bySheet)
    {
        foreach (var worksheet in workbook.Worksheets)
        {
            if (!bySheet.TryGetValue(worksheet.Name, out var expansions))
            {
                continue;
            }

            // Materialised: moving a pivot table rewrites the sheet's pivot areas.
            foreach (var pivot in worksheet.PivotTables.ToList())
            {
                var target = pivot.TargetCell;
                var row = target.Address.RowNumber;
                var column = target.Address.ColumnNumber;

                foreach (var expansion in expansions)
                {
                    row = ExpansionMap.MapRowStart(row, expansion);
                    column = ExpansionMap.MapColumnStart(column, expansion);
                }

                if (row != target.Address.RowNumber || column != target.Address.ColumnNumber)
                {
                    pivot.TargetCell = worksheet.Cell(row, column);
                }
            }
        }
    }
}
