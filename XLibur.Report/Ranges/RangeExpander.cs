using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using XLibur.Excel.ConditionalFormats;
using XLibur.Excel.Coordinates;
using XLibur.Report.Expressions;

namespace XLibur.Report.Ranges;

/// <summary>
/// Repeats a bound range's template rows once per data item.
/// </summary>
/// <remarks>
/// Expansion inserts sheet rows and copies the template block into them, rather than rendering
/// onto a scratch sheet and splicing the result back. Insert-and-copy hands formula adjustment,
/// merge tracking, conditional-format extension, row sizing and defined-name shifting to the core
/// library, all of which it already does correctly and quickly. It also means a conditional format
/// over the template rows is <em>stretched</em> over the generated ones, rather than being copied
/// once per generated cell — the behaviour ClosedXML.Report is criticised for in its issue #216.
/// <para>
/// The last row of a multi-row range is its options row: it carries tags and summary formulas, and
/// is not repeated. A single-row range has no options row, so the whole row repeats.
/// </para>
/// </remarks>
internal sealed class RangeExpander
{
    private readonly CellEvaluator _evaluator;

    public RangeExpander(CellEvaluator evaluator) => _evaluator = evaluator;

    /// <summary>
    /// Expands <paramref name="bound"/> in place and returns what it did, or <c>null</c> if the
    /// range had already been removed from the sheet.
    /// </summary>
    public ExpansionRecord? Expand(BoundRange bound, ExpressionScope globalScope)
    {
        var definedName = bound.DefinedName;

        // Re-read the address: expanding a range higher up the sheet has already moved this one.
        var range = definedName.Ranges.FirstOrDefault();
        if (range is null)
        {
            return null;
        }

        var sheet = range.Worksheet;
        var area = RangeArea.From(range);
        var hasOptionsRow = area.RowCount > 1;
        var dataLastRow = hasOptionsRow ? area.LastRow - 1 : area.LastRow;
        var dataRowCount = dataLastRow - area.FirstRow + 1;
        var items = bound.Items;

        if (items.Count == 0)
        {
            return Clear(definedName, sheet, area, dataLastRow, hasOptionsRow);
        }

        // Captured before anything moves: the rules are re-pointed afterwards from the template's
        // original geometry.
        var templateFormats = CaptureConditionalFormats(sheet, area.FirstRow, dataLastRow);
        var formatsBeforeExpansion = new HashSet<IXLConditionalFormat>(
            sheet.ConditionalFormats,
            ReferenceEqualityComparer.Instance);

        var extraRows = (items.Count - 1) * dataRowCount;
        if (extraRows > 0)
        {
            sheet.Row(dataLastRow).InsertRowsBelow(extraRows);
        }

        // Every copy is taken from the still-unevaluated template block, so each item's rows start
        // from the same source; evaluation happens afterwards, once per block.
        var template = sheet.Range(area.FirstRow, area.FirstColumn, dataLastRow, area.LastColumn);
        for (var i = 1; i < items.Count; i++)
        {
            template.CopyTo(sheet.Cell(area.FirstRow + (i * dataRowCount), area.FirstColumn));
        }

        for (var i = 0; i < items.Count; i++)
        {
            var blockFirstRow = area.FirstRow + (i * dataRowCount);
            var scope = CreateItemScope(globalScope, items, i);
            EvaluateBlock(sheet, blockFirstRow, blockFirstRow + dataRowCount - 1, area, scope);
        }

        var renderedLastRow = area.FirstRow + (items.Count * dataRowCount) - 1;

        RestoreConditionalFormats(sheet, formatsBeforeExpansion, templateFormats, dataRowCount, items.Count);

        if (hasOptionsRow)
        {
            RemoveOptionsRowIfEmpty(sheet, renderedLastRow + 1, area);
        }

        var rendered = new RangeArea(area.FirstRow, area.FirstColumn, renderedLastRow, area.LastColumn);
        definedName.SetRefersTo(sheet.Range(rendered.FirstRow, rendered.FirstColumn, rendered.LastRow, rendered.LastColumn));

        return new ExpansionRecord(sheet, area, rendered, rendered.RowCount - area.RowCount);
    }

    /// <summary>
    /// Removes a range bound to no data. An options row that still holds content is kept — it may
    /// carry a total that should show as zero rather than vanish.
    /// </summary>
    private static ExpansionRecord? Clear(
        IXLDefinedName definedName,
        IXLWorksheet sheet,
        RangeArea area,
        int dataLastRow,
        bool hasOptionsRow)
    {
        var optionsRowEmpty = !hasOptionsRow || IsRowEmpty(sheet, area.LastRow, area);
        var lastRowToDelete = optionsRowEmpty ? area.LastRow : dataLastRow;

        sheet.Rows(area.FirstRow, lastRowToDelete).Delete();

        if (optionsRowEmpty)
        {
            definedName.Delete();
        }

        var rendered = new RangeArea(area.FirstRow, area.FirstColumn, area.FirstRow - 1, area.LastColumn);
        return new ExpansionRecord(sheet, area, rendered, -(lastRowToDelete - area.FirstRow + 1));
    }

    private static ExpressionScope CreateItemScope(ExpressionScope globalScope, IReadOnlyList<object?> items, int index) =>
        globalScope.CreateChild(new[]
        {
            new KeyValuePair<string, object?>("item", items[index]),
            new KeyValuePair<string, object?>("index", index),
            new KeyValuePair<string, object?>("items", items),
        });

    private void EvaluateBlock(IXLWorksheet sheet, int firstRow, int lastRow, RangeArea area, ExpressionScope scope)
    {
        var block = sheet.Range(firstRow, area.FirstColumn, lastRow, area.LastColumn);

        // Materialised: evaluating a cell writes to it, and writing while enumerating the used-cell
        // set is not safe.
        foreach (var cell in block.CellsUsed(XLCellsUsedOptions.Contents).ToList())
        {
            _evaluator.Evaluate(cell, scope);
        }
    }

    private static void RemoveOptionsRowIfEmpty(IXLWorksheet sheet, int rowNumber, RangeArea area)
    {
        if (IsRowEmpty(sheet, rowNumber, area))
        {
            sheet.Row(rowNumber).Delete();
        }
    }

    /// <summary>
    /// Records the conditional formatting rules that live entirely inside the template block,
    /// together with where they sit, before expansion moves anything.
    /// </summary>
    private static List<CapturedConditionalFormat> CaptureConditionalFormats(
        IXLWorksheet sheet,
        int dataFirstRow,
        int dataLastRow)
    {
        var captured = new List<CapturedConditionalFormat>();

        foreach (var format in sheet.ConditionalFormats)
        {
            var areas = format.Ranges
                .Select(RangeArea.From)
                .Where(a => a.FirstRow >= dataFirstRow && a.LastRow <= dataLastRow)
                .ToList();

            if (areas.Count > 0)
            {
                captured.Add(new CapturedConditionalFormat(format, areas));
            }
        }

        return captured;
    }

    /// <summary>
    /// Puts conditional formatting back the way a report author expects: one rule covering the
    /// generated rows, rather than a copy of the rule per generated row.
    /// </summary>
    /// <remarks>
    /// Copying the template block carries its conditional formats with it, which is how
    /// ClosedXML.Report ends up with as many rules as it generated cells — the complaint in its
    /// issue #216, where three rules over three rows become nine and, as the reporter puts it,
    /// "kills the generation time". The copies are discarded here and the original rule is widened
    /// over every block instead, so the generated workbook holds exactly the rules the template
    /// declared.
    /// </remarks>
    private static void RestoreConditionalFormats(
        IXLWorksheet sheet,
        HashSet<IXLConditionalFormat> formatsBeforeExpansion,
        List<CapturedConditionalFormat> captured,
        int dataRowCount,
        int itemCount)
    {
        if (captured.Count == 0)
        {
            return;
        }

        sheet.ConditionalFormats.Remove(format => !formatsBeforeExpansion.Contains(format));

        foreach (var (format, areas) in captured)
        {
            var widened = areas
                .SelectMany(original => Widen(sheet, original, dataRowCount, itemCount))
                .Cast<IXLRange>()
                .ToList();

            // IXLConditionalFormat.Ranges is a fresh projection of the rule's internal area list,
            // so mutating it does nothing; SetAreas is the supported way to rewrite coverage.
            ((XLConditionalFormat)format).SetAreas(XLAreaList.FromRanges(widened));
        }
    }

    /// <summary>
    /// Projects a rule's original rectangle over every generated block, as one range when the
    /// blocks abut and as one range per block when they do not.
    /// </summary>
    private static IEnumerable<IXLRange> Widen(IXLWorksheet sheet, RangeArea original, int dataRowCount, int itemCount)
    {
        if (original.RowCount == dataRowCount)
        {
            yield return sheet.Range(
                original.FirstRow,
                original.FirstColumn,
                original.FirstRow + (itemCount * dataRowCount) - 1,
                original.LastColumn);
            yield break;
        }

        for (var i = 0; i < itemCount; i++)
        {
            var offset = i * dataRowCount;
            yield return sheet.Range(
                original.FirstRow + offset,
                original.FirstColumn,
                original.LastRow + offset,
                original.LastColumn);
        }
    }

    private static bool IsRowEmpty(IXLWorksheet sheet, int rowNumber, RangeArea area) =>
        !sheet.Range(rowNumber, area.FirstColumn, rowNumber, area.LastColumn)
            .CellsUsed(XLCellsUsedOptions.Contents)
            .Any();
}

/// <summary>A conditional formatting rule and the rectangles it covered before expansion.</summary>
internal readonly record struct CapturedConditionalFormat(IXLConditionalFormat Format, List<RangeArea> Areas);

/// <summary>
/// What one expansion did to a sheet: where the template was, where the generated block ended up,
/// and how far everything below it moved.
/// </summary>
/// <remarks>
/// Collected so that references pointing into or below the template — chart series, picture
/// anchors, pivot cache sources — can be re-pointed once the sheet has settled.
/// </remarks>
internal sealed record ExpansionRecord(
    IXLWorksheet Worksheet,
    RangeArea TemplateArea,
    RangeArea RenderedArea,
    int RowDelta);
