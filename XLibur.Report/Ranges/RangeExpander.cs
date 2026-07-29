using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using XLibur.Excel.ConditionalFormats;
using XLibur.Excel.Coordinates;
using XLibur.Report.Expressions;
using XLibur.Report.Tags;

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
    private readonly IExpressionEngine _engine;
    private readonly TemplateErrors _errors;

    public RangeExpander(CellEvaluator evaluator, IExpressionEngine engine, TemplateErrors errors)
    {
        _evaluator = evaluator;
        _engine = engine;
        _errors = errors;
    }

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

        // Read before anything else, and in this order: reading the tags strips their text, so a
        // column holding an expression beside a tag is only recognised as that expression once the
        // tag has been taken out of it.
        var tags = ReadTags(sheet, area, dataLastRow, hasOptionsRow);
        var columnExpressions = ReadColumnExpressions(sheet, area, dataLastRow);
        var groupOptions = GroupOptions.Read(tags);

        // The generated rows do not exist yet, so this context points at the template block; a tag
        // acting at this stage is reordering data, not reading the sheet.
        var templateContext = new ProcessingContext(
            sheet,
            sheet.Range(area.FirstRow, area.FirstColumn, dataLastRow, area.LastColumn),
            null,
            bound.Items,
            _engine,
            globalScope,
            _errors,
            columnExpressions,
            groupOptions.GrandTotalDisabled);

        var items = ApplyItemTransforms(tags, bound.Items, templateContext);

        if (items.Count == 0)
        {
            return Clear(definedName, sheet, area, dataLastRow, hasOptionsRow);
        }

        // Grouping reorders too, and has to do it last: its ordering is stable, so whatever
        // <<Sort>> decided survives as the order within each group.
        var grouping = GroupRenderer.Prepare(tags, items, groupOptions, templateContext);
        if (grouping is not null)
        {
            items = grouping.Items;
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
            var scope = ItemScopes.For(globalScope, items, i);
            EvaluateBlock(sheet, blockFirstRow, blockFirstRow + dataRowCount - 1, area, scope);
        }

        var renderedLastRow = area.FirstRow + (items.Count * dataRowCount) - 1;

        RestoreConditionalFormats(sheet, formatsBeforeExpansion, templateFormats, dataRowCount, items.Count);

        if (grouping is not null)
        {
            // Grouping runs before the options row is written, so that the grand total spans the
            // group totals as well — SUBTOTAL ignores nested SUBTOTALs, which is why it is the
            // formula the summary tags leave.
            var groupContext = new ProcessingContext(
                sheet,
                sheet.Range(area.FirstRow, area.FirstColumn, renderedLastRow, area.LastColumn),
                null,
                items,
                _engine,
                globalScope,
                _errors,
                columnExpressions,
                groupOptions.GrandTotalDisabled);

            renderedLastRow = grouping.Render(
                sheet,
                area,
                area.FirstRow,
                dataRowCount,
                hasOptionsRow,
                tags,
                groupContext);
        }

        if (hasOptionsRow)
        {
            var optionsRowNumber = renderedLastRow + 1;

            // Tags run before the options row is considered for removal: a total writes into it,
            // and a row holding a total is not an empty row.
            ExecuteTags(
                tags,
                sheet,
                area,
                renderedLastRow,
                optionsRowNumber,
                items,
                globalScope,
                columnExpressions,
                groupOptions.GrandTotalDisabled);

            RemoveOptionsRowIfEmpty(sheet, optionsRowNumber, area);
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
    /// Reads every tag the range declares — those in the options row, and those in the rows it
    /// repeats — ordered so that a tag which has to see what another did runs after it.
    /// </summary>
    /// <remarks>
    /// A tag in a repeated row is describing a row rather than the range, which is what
    /// <see cref="OptionTag.InRepeatedRow"/> tells it. Read before anything else happens, because a
    /// repeated row's tag text must not survive into the copies.
    /// </remarks>
    private List<OptionTag> ReadTags(IXLWorksheet sheet, RangeArea area, int dataLastRow, bool hasOptionsRow)
    {
        var tags = new List<OptionTag>();

        for (var row = area.FirstRow; row <= dataLastRow; row++)
        {
            ReadTagsFromRow(sheet, row, area, inRepeatedRow: true, tags);
        }

        if (hasOptionsRow)
        {
            ReadTagsFromRow(sheet, area.LastRow, area, inRepeatedRow: false, tags);
        }

        return tags.OrderBy(tag => TagsRegister.PriorityOf(tag.Token.Name)).ToList();
    }

    /// <summary>
    /// Reads the tags out of one row, clearing their text so it does not reach the report.
    /// </summary>
    private void ReadTagsFromRow(
        IXLWorksheet sheet,
        int rowNumber,
        RangeArea area,
        bool inRepeatedRow,
        List<OptionTag> tags)
    {
        for (var column = area.FirstColumn; column <= area.LastColumn; column++)
        {
            var cell = sheet.Cell(rowNumber, column);
            if (!cell.Value.IsText)
            {
                continue;
            }

            var text = cell.Value.GetText();
            if (!TagParser.Contains(text))
            {
                continue;
            }

            foreach (var token in TagParser.Parse(text))
            {
                if (TagsRegister.TryCreate(token, column, out var tag))
                {
                    tag.InRepeatedRow = inRepeatedRow;
                    tags.Add(tag);
                }
                else
                {
                    _errors.Add(new TemplateError(
                        $"<<{token.Name}>> is not a tag this library knows.",
                        sheet.Name,
                        cell.Address.ToString()));
                }
            }

            var remaining = TagParser.Strip(text);
            if (remaining.Length == 0)
            {
                // Cleared rather than set to an empty string, so an options row that held nothing
                // but tags still counts as empty and is removed.
                cell.Clear(XLClearOptions.Contents);
            }
            else
            {
                cell.Value = remaining;
            }
        }
    }

    /// <summary>
    /// Records what expression each column holds, so a column-placed tag can tell what the column
    /// means without the template having to say it twice.
    /// </summary>
    private static Dictionary<int, string> ReadColumnExpressions(IXLWorksheet sheet, RangeArea area, int dataLastRow)
    {
        var expressions = new Dictionary<int, string>();

        for (var row = area.FirstRow; row <= dataLastRow; row++)
        {
            for (var column = area.FirstColumn; column <= area.LastColumn; column++)
            {
                if (expressions.ContainsKey(column))
                {
                    continue;
                }

                var value = sheet.Cell(row, column).Value;
                if (value.IsText && ExpressionText.TryGetSingleExpression(value.GetText(), out var expression))
                {
                    expressions[column] = expression;
                }
            }
        }

        return expressions;
    }

    private static IReadOnlyList<object?> ApplyItemTransforms(
        List<OptionTag> tags,
        IReadOnlyList<object?> items,
        ProcessingContext context)
    {
        foreach (var tag in tags)
        {
            items = tag.TransformItems(items, context);
        }

        return items;
    }

    private void ExecuteTags(
        List<OptionTag> tags,
        IXLWorksheet sheet,
        RangeArea area,
        int renderedLastRow,
        int optionsRowNumber,
        IReadOnlyList<object?> items,
        ExpressionScope globalScope,
        Dictionary<int, string> columnExpressions,
        bool grandTotalsDisabled)
    {
        if (tags.Count == 0)
        {
            return;
        }

        var context = new ProcessingContext(
            sheet,
            sheet.Range(area.FirstRow, area.FirstColumn, renderedLastRow, area.LastColumn),
            sheet.Range(optionsRowNumber, area.FirstColumn, optionsRowNumber, area.LastColumn),
            items,
            _engine,
            globalScope,
            _errors,
            columnExpressions,
            grandTotalsDisabled);

        foreach (var tag in tags)
        {
            tag.Execute(context);
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
