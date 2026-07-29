using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using XLibur.Report.Ranges;

namespace XLibur.Report.Rewriting;

/// <summary>
/// Re-points the things that refer to a range by address, once expansion has moved it.
/// </summary>
/// <remarks>
/// <para>
/// A cell formula, a defined name, a conditional format and a picture anchor are all shifted by the
/// core library when rows are inserted, because each is held as a live range. A chart series
/// reference is not: it is a plain string in the model and a plain string in the file, and nothing
/// shifts it. So a template chart plotting the one row a range repeats keeps plotting that one row
/// however many the report generated — which is exactly the complaint behind ClosedXML.Report's
/// issues #123 and #351, where series references "silently go stale after expansion".
/// </para>
/// <para>
/// The rewriter closes that gap, and only that gap: pictures were measured to move on their own
/// (see <c>DrawingMechanicsCharacterizationTests</c>), so there is deliberately no picture code
/// here.
/// </para>
/// </remarks>
internal static class ReferenceRewriter
{
    /// <summary>
    /// Re-points every chart series in <paramref name="workbook"/> that refers to a range
    /// <paramref name="expansions"/> grew or moved.
    /// </summary>
    public static void Rewrite(IXLWorkbook workbook, IReadOnlyList<ExpansionRecord> expansions)
    {
        if (expansions.Count == 0)
        {
            return;
        }

        // Grouped by the sheet the data is on, not the sheet the chart is on: a chart usually sits
        // beside its data but is under no obligation to.
        var bySheet = ExpansionMap.BySheet(expansions);

        foreach (var worksheet in workbook.Worksheets)
        {
            foreach (var chart in worksheet.Charts)
            {
                foreach (var series in chart.Series.Concat(chart.SecondarySeries))
                {
                    RewriteSeries(series, worksheet.Name, bySheet);
                }
            }
        }
    }

    private static void RewriteSeries(
        IXLChartSeries series,
        string chartSheetName,
        Dictionary<string, List<ExpansionRecord>> bySheet)
    {
        if (TryRewrite(series.ValueReferences, chartSheetName, bySheet, out var values))
        {
            series.ValueReferences = values;
        }

        if (TryRewrite(series.CategoryReferences, chartSheetName, bySheet, out var categories))
        {
            series.CategoryReferences = categories;
        }
    }

    /// <summary>
    /// Applies every expansion on the referenced sheet, in the order they ran, and reports whether
    /// the reference ended up anywhere different.
    /// </summary>
    /// <remarks>
    /// In order, because each expansion recorded where things were when <em>it</em> ran: the second
    /// range on a sheet was already sitting lower than the template put it. Replaying them in the
    /// same order walks the reference through the same coordinate changes the sheet went through.
    /// </remarks>
    private static bool TryRewrite(
        string? reference,
        string chartSheetName,
        Dictionary<string, List<ExpansionRecord>> bySheet,
        out string rewritten)
    {
        rewritten = string.Empty;

        if (!SeriesReference.TryParse(reference, out var parsed))
        {
            return false;
        }

        // An unqualified reference means the sheet the chart is on.
        var sheetName = parsed.SheetName ?? chartSheetName;
        if (!bySheet.TryGetValue(sheetName, out var expansions))
        {
            return false;
        }

        var current = parsed;
        foreach (var expansion in expansions)
        {
            current = Apply(current, expansion);
        }

        if (current == parsed)
        {
            return false;
        }

        rewritten = (current with { SheetName = parsed.SheetName }).ToText();
        return true;
    }

    /// <summary>
    /// Moves one reference through one expansion.
    /// </summary>
    /// <remarks>
    /// A reference sharing no column with the template cannot be stretched by the expansion — a
    /// chart plotting a column beside the range is not plotting the range — but it still moves if it
    /// sits below one, because the insert that grew the range was a full-row one.
    /// </remarks>
    private static SeriesReference Apply(SeriesReference reference, ExpansionRecord expansion)
    {
        var template = expansion.TemplateArea;

        if (reference.LastColumn < template.FirstColumn || reference.FirstColumn > template.LastColumn)
        {
            return reference.FirstRow <= template.LastRow
                ? reference
                : reference with
                {
                    FirstRow = reference.FirstRow + expansion.RowDelta,
                    LastRow = reference.LastRow + expansion.RowDelta,
                };
        }

        var firstRow = ExpansionMap.MapStart(reference.FirstRow, expansion);
        var lastRow = ExpansionMap.MapEnd(reference.LastRow, expansion);

        return reference with
        {
            FirstRow = firstRow,
            LastRow = Math.Max(firstRow, lastRow),
        };
    }
}
