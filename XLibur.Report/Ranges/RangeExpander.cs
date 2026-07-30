using System.Collections.Generic;
using System.Linq;
using XLibur.Excel;
using XLibur.Excel.ConditionalFormats;
using XLibur.Excel.Coordinates;
using XLibur.Report.Expressions;
using XLibur.Report.Tags;

namespace XLibur.Report.Ranges;

/// <summary>
/// Repeats a bound range's template slots once per data item.
/// </summary>
/// <remarks>
/// <para>
/// Expansion inserts sheet rows — or columns, for a range that repeats across — and copies the
/// template block into them, rather than rendering onto a scratch sheet and splicing the result
/// back. Insert-and-copy hands formula adjustment, merge tracking, conditional-format extension,
/// row sizing and defined-name shifting to the core library, all of which it already does correctly
/// and quickly. It also means a conditional format over the template slots is <em>stretched</em>
/// over the generated ones, rather than being copied once per generated cell — the behaviour
/// ClosedXML.Report is criticised for in its issue #216.
/// </para>
/// <para>
/// The last slot of a range spanning more than one is its options slot: it carries the tags and the
/// summary formulas, and is not repeated. A range one slot deep has no options slot, so the whole of
/// it repeats.
/// </para>
/// <para>
/// Everything here is written against <see cref="RangeAxis"/> rather than against rows, so the
/// vertical and horizontal cases are the same code read two ways. See that type for the
/// slot-and-line vocabulary.
/// </para>
/// </remarks>
internal sealed class RangeExpander
{
    private readonly CellEvaluator _evaluator;
    private readonly IExpressionEngine _engine;
    private readonly TemplateErrors _errors;
    private readonly List<Rewriting.PivotRequest> _pendingPivots;

    public RangeExpander(
        CellEvaluator evaluator,
        IExpressionEngine engine,
        TemplateErrors errors,
        List<Rewriting.PivotRequest> pendingPivots)
    {
        _evaluator = evaluator;
        _engine = engine;
        _errors = errors;
        _pendingPivots = pendingPivots;
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
        var axis = DetectAxis(sheet, area);

        var firstSlot = axis.FirstSlot(area);
        var hasOptionsSlot = axis.SlotCount(area) > 1;
        var dataLastSlot = hasOptionsSlot ? axis.LastSlot(area) - 1 : axis.LastSlot(area);
        var slotsPerItem = dataLastSlot - firstSlot + 1;

        // Read before anything else, and in this order: reading the tags strips their text, so a
        // line holding an expression beside a tag is only recognised as that expression once the
        // tag has been taken out of it.
        var tags = ReadTags(sheet, area, axis, dataLastSlot, hasOptionsSlot);
        var lineExpressions = ReadLineExpressions(sheet, area, axis, dataLastSlot);
        var groupOptions = GroupOptions.Read(tags);

        // The generated slots do not exist yet, so this context points at the template block; a tag
        // acting at this stage is reordering data, not reading the sheet.
        var templateContext = Context(
            sheet, axis, area, firstSlot, dataLastSlot, null, bound.Items, globalScope, lineExpressions, tags,
            groupOptions);

        var items = ApplyItemTransforms(tags, bound.Items, templateContext);

        if (items.Count == 0)
        {
            return Clear(definedName, sheet, area, axis, dataLastSlot, hasOptionsSlot);
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
        var templateFormats = CaptureConditionalFormats(sheet, axis, area, firstSlot, dataLastSlot);
        var formatsBeforeExpansion = new HashSet<IXLConditionalFormat>(
            sheet.ConditionalFormats,
            ReferenceEqualityComparer.Instance);

        var extraSlots = (items.Count - 1) * slotsPerItem;
        if (extraSlots > 0)
        {
            axis.InsertSlotsAfter(sheet, dataLastSlot, extraSlots);
        }

        FillWithCopiesOfTemplate(sheet, axis, area, firstSlot, slotsPerItem, items.Count);

        for (var i = 0; i < items.Count; i++)
        {
            var blockFirstSlot = firstSlot + (i * slotsPerItem);
            var scope = ItemScopes.For(globalScope, items, i);
            EvaluateBlock(sheet, axis.Slots(area, blockFirstSlot, blockFirstSlot + slotsPerItem - 1), scope);
        }

        var renderedLastSlot = firstSlot + (items.Count * slotsPerItem) - 1;

        RestoreConditionalFormats(sheet, axis, formatsBeforeExpansion, templateFormats, slotsPerItem, items.Count);

        if (grouping is not null)
        {
            // Grouping runs before the options row is written, so that the grand total spans the
            // group totals as well — SUBTOTAL ignores nested SUBTOTALs, which is why it is the
            // formula the summary tags leave.
            var groupContext = Context(
                sheet, axis, area, firstSlot, renderedLastSlot, null, items, globalScope, lineExpressions, tags,
                groupOptions);

            renderedLastSlot = grouping.Render(
                sheet,
                area,
                firstSlot,
                slotsPerItem,
                hasOptionsSlot,
                tags,
                groupContext);
        }

        // How many slots the range ends up occupying, which is not the same as how many it produced:
        // an options slot that a total was written into stays, and everything past the range moves
        // one further because of it.
        var occupiedSlots = renderedLastSlot - firstSlot + 1;

        if (hasOptionsSlot)
        {
            var optionsSlot = renderedLastSlot + 1;
            var optionsRange = RangeAxis.Range(sheet, axis.Slots(area, optionsSlot, optionsSlot));

            // Tags run before the options slot is considered for removal: a total writes into it,
            // and a slot holding a total is not an empty one.
            var executeContext = Context(
                sheet, axis, area, firstSlot, renderedLastSlot, optionsRange, items, globalScope, lineExpressions,
                tags, groupOptions);

            foreach (var tag in tags)
            {
                tag.Execute(executeContext);
            }

            if (!RemoveOptionsSlotIfEmpty(sheet, axis, area, optionsSlot))
            {
                occupiedSlots++;
            }
        }

        var rendered = axis.Slots(area, firstSlot, renderedLastSlot);
        definedName.SetRefersTo(RangeAxis.Range(sheet, rendered));

        return new ExpansionRecord(sheet, area, rendered, axis, occupiedSlots - axis.SlotCount(area));
    }

    /// <summary>
    /// Which way the range repeats. Looked for in the last column, the one place that means the same
    /// thing before the answer is known: a vertical range's options row ends there and a horizontal
    /// range's options column <em>is</em> there.
    /// </summary>
    private static RangeAxis DetectAxis(IXLWorksheet sheet, RangeArea area)
    {
        for (var row = area.FirstRow; row <= area.LastRow; row++)
        {
            var value = sheet.Cell(row, area.LastColumn).Value;
            if (!value.IsText)
            {
                continue;
            }

            foreach (var token in TagParser.Parse(value.GetText()))
            {
                if (string.Equals(token.Name, "Horizontal", System.StringComparison.OrdinalIgnoreCase))
                {
                    return RangeAxis.Horizontal;
                }
            }
        }

        return RangeAxis.Vertical;
    }

    private ProcessingContext Context(
        IXLWorksheet sheet,
        RangeAxis axis,
        RangeArea area,
        int firstSlot,
        int lastSlot,
        IXLRange? optionsSlot,
        IReadOnlyList<object?> items,
        ExpressionScope globalScope,
        Dictionary<int, string> lineExpressions,
        List<OptionTag> tags,
        GroupOptions groupOptions) =>
        new(
            sheet,
            RangeAxis.Range(sheet, axis.Slots(area, firstSlot, lastSlot)),
            optionsSlot,
            items,
            _engine,
            globalScope,
            _errors,
            lineExpressions,
            axis,
            tags,
            _pendingPivots,
            groupOptions.GrandTotalDisabled);

    /// <summary>
    /// Removes a range bound to no data. An options slot that still holds content is kept — it may
    /// carry a total that should show as zero rather than vanish.
    /// </summary>
    private static ExpansionRecord? Clear(
        IXLDefinedName definedName,
        IXLWorksheet sheet,
        RangeArea area,
        RangeAxis axis,
        int dataLastSlot,
        bool hasOptionsSlot)
    {
        var firstSlot = axis.FirstSlot(area);
        var lastSlot = axis.LastSlot(area);
        var optionsEmpty = !hasOptionsSlot || IsSlotEmpty(sheet, axis, area, lastSlot);
        var lastToDelete = optionsEmpty ? lastSlot : dataLastSlot;

        axis.DeleteSlots(sheet, firstSlot, lastToDelete);

        if (optionsEmpty)
        {
            definedName.Delete();
        }

        // An empty rendered area: its last slot is before its first, which is what everything
        // downstream reads as "nothing was produced".
        var rendered = axis.Slots(area, firstSlot, firstSlot - 1);

        return new ExpansionRecord(sheet, area, rendered, axis, -(lastToDelete - firstSlot + 1));
    }

    /// <summary>
    /// Fills the inserted slots with copies of the template block, by repeatedly doubling what has
    /// already been written rather than copying the template once per item.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The two are interchangeable because every copy is taken from the <em>unevaluated</em> template:
    /// after any number of copies the filled region is n identical blocks, so copying the first k of
    /// them onto the end is the same as copying the template k times. Evaluation happens afterwards,
    /// once per item, and is unaffected.
    /// </para>
    /// <para>
    /// It matters because <c>CopyTo</c> costs more the larger the target worksheet is — measurably, and
    /// independently of how much is being copied: <c>ExpansionPhaseProbe</c> shows the per-call cost
    /// tripling over 30,000 rows with no report code involved. One call per item therefore made
    /// generation super-linear in the row count, which was the one thing spec 12's benchmark criterion
    /// asked it not to be. Doubling makes it ⌈log₂ n⌉ calls — sixteen for fifty thousand rows instead of
    /// fifty thousand — and the same total number of cells copied.
    /// </para>
    /// </remarks>
    private static void FillWithCopiesOfTemplate(
        IXLWorksheet sheet,
        RangeAxis axis,
        RangeArea area,
        int firstSlot,
        int slotsPerItem,
        int itemCount)
    {
        var line = axis.FirstLine(area);
        var filled = 1;

        while (filled < itemCount)
        {
            // The last round copies only as many blocks as are still wanted, so the region never
            // overshoots the slots that were inserted for it.
            var copies = System.Math.Min(filled, itemCount - filled);

            var source = RangeAxis.Range(
                sheet,
                axis.Slots(area, firstSlot, firstSlot + (copies * slotsPerItem) - 1));

            source.CopyTo(axis.Cell(sheet, firstSlot + (filled * slotsPerItem), line));

            filled += copies;
        }
    }

    private void EvaluateBlock(IXLWorksheet sheet, RangeArea block, ExpressionScope scope)
    {
        // Materialised: evaluating a cell writes to it, and writing while enumerating the used-cell
        // set is not safe.
        foreach (var cell in RangeAxis.Range(sheet, block).CellsUsed(XLCellsUsedOptions.Contents).ToList())
        {
            _evaluator.Evaluate(cell, scope);
        }
    }

    /// <summary>
    /// Drops the options slot when nothing was left in it, reporting whether it did — the caller has
    /// to know, because a surviving options slot moves everything past the range one further.
    /// </summary>
    private static bool RemoveOptionsSlotIfEmpty(IXLWorksheet sheet, RangeAxis axis, RangeArea area, int slot)
    {
        if (!IsSlotEmpty(sheet, axis, area, slot))
        {
            return false;
        }

        axis.DeleteSlots(sheet, slot, slot);
        return true;
    }

    /// <summary>
    /// Reads every tag the range declares — those in the options slot, and those in the slots it
    /// repeats — ordered so that a tag which has to see what another did runs after it.
    /// </summary>
    /// <remarks>
    /// A tag in a repeated slot is describing one item rather than the range, which is what
    /// <see cref="OptionTag.InRepeatedRow"/> tells it. Read before anything else happens, because a
    /// repeated slot's tag text must not survive into the copies.
    /// </remarks>
    private List<OptionTag> ReadTags(
        IXLWorksheet sheet,
        RangeArea area,
        RangeAxis axis,
        int dataLastSlot,
        bool hasOptionsSlot)
    {
        var tags = new List<OptionTag>();

        for (var slot = axis.FirstSlot(area); slot <= dataLastSlot; slot++)
        {
            ReadTagsFromSlot(sheet, area, axis, slot, inRepeatedSlot: true, tags);
        }

        if (hasOptionsSlot)
        {
            ReadTagsFromSlot(sheet, area, axis, axis.LastSlot(area), inRepeatedSlot: false, tags);
        }

        return tags.OrderBy(tag => TagsRegister.PriorityOf(tag.Token.Name)).ToList();
    }

    /// <summary>
    /// Reads the tags out of one slot, clearing their text so it does not reach the report.
    /// </summary>
    private void ReadTagsFromSlot(
        IXLWorksheet sheet,
        RangeArea area,
        RangeAxis axis,
        int slot,
        bool inRepeatedSlot,
        List<OptionTag> tags)
    {
        for (var line = axis.FirstLine(area); line <= axis.LastLine(area); line++)
        {
            var cell = axis.Cell(sheet, slot, line);
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
                if (TagsRegister.TryCreate(token, cell.Address.RowNumber, cell.Address.ColumnNumber, line, out var tag))
                {
                    tag.InRepeatedRow = inRepeatedSlot;
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
                // Cleared rather than set to an empty string, so an options slot that held nothing
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
    /// Records what expression each line holds, so a tag sitting in a line can tell what that line
    /// means without the template having to say it twice.
    /// </summary>
    private static Dictionary<int, string> ReadLineExpressions(
        IXLWorksheet sheet,
        RangeArea area,
        RangeAxis axis,
        int dataLastSlot)
    {
        var expressions = new Dictionary<int, string>();

        for (var slot = axis.FirstSlot(area); slot <= dataLastSlot; slot++)
        {
            for (var line = axis.FirstLine(area); line <= axis.LastLine(area); line++)
            {
                if (expressions.ContainsKey(line))
                {
                    continue;
                }

                var value = axis.Cell(sheet, slot, line).Value;
                if (value.IsText && ExpressionText.TryGetSingleExpression(value.GetText(), out var expression))
                {
                    expressions[line] = expression;
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

    /// <summary>
    /// Records the conditional formatting rules that live entirely inside the template block,
    /// together with where they sit, before expansion moves anything.
    /// </summary>
    private static List<CapturedConditionalFormat> CaptureConditionalFormats(
        IXLWorksheet sheet,
        RangeAxis axis,
        RangeArea area,
        int dataFirstSlot,
        int dataLastSlot)
    {
        var captured = new List<CapturedConditionalFormat>();

        foreach (var format in sheet.ConditionalFormats)
        {
            var areas = format.Ranges
                .Select(RangeArea.From)
                .Where(a => axis.FirstSlot(a) >= dataFirstSlot && axis.LastSlot(a) <= dataLastSlot)
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
        RangeAxis axis,
        HashSet<IXLConditionalFormat> formatsBeforeExpansion,
        List<CapturedConditionalFormat> captured,
        int slotsPerItem,
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
                .SelectMany(original => Widen(sheet, axis, original, slotsPerItem, itemCount))
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
    private static IEnumerable<IXLRange> Widen(
        IXLWorksheet sheet,
        RangeAxis axis,
        RangeArea original,
        int slotsPerItem,
        int itemCount)
    {
        var firstSlot = axis.FirstSlot(original);

        if (axis.SlotCount(original) == slotsPerItem)
        {
            yield return RangeAxis.Range(
                sheet,
                axis.Slots(original, firstSlot, firstSlot + (itemCount * slotsPerItem) - 1));
            yield break;
        }

        for (var i = 0; i < itemCount; i++)
        {
            var offset = i * slotsPerItem;
            yield return RangeAxis.Range(
                sheet,
                axis.Slots(original, firstSlot + offset, axis.LastSlot(original) + offset));
        }
    }

    private static bool IsSlotEmpty(IXLWorksheet sheet, RangeAxis axis, RangeArea area, int slot) =>
        !RangeAxis.Range(sheet, axis.Slots(area, slot, slot))
            .CellsUsed(XLCellsUsedOptions.Contents)
            .Any();
}

/// <summary>A conditional formatting rule and the rectangles it covered before expansion.</summary>
internal readonly record struct CapturedConditionalFormat(IXLConditionalFormat Format, List<RangeArea> Areas);

/// <summary>
/// What one expansion did to a sheet: where the template was, where the generated block ended up,
/// and how far everything past it moved.
/// </summary>
/// <remarks>
/// Collected so that references pointing into or past the template — chart series, pivot cache
/// sources, pivot table positions — can be re-pointed once the sheet has settled. The axis says
/// which way the range ran, and so which dimension <see cref="SlotDelta"/> moved: a range repeating
/// downwards moves what is below it, one repeating across moves what is to its right.
/// </remarks>
internal sealed record ExpansionRecord(
    IXLWorksheet Worksheet,
    RangeArea TemplateArea,
    RangeArea RenderedArea,
    RangeAxis Axis,
    int SlotDelta);
