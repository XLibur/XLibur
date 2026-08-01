using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using XLibur.Excel;
using XLibur.Report.Expressions;
using XLibur.Report.Tags;

namespace XLibur.Report.Ranges;

/// <summary>
/// Groups the rows a bound range generated: orders them so that each group is contiguous, inserts a
/// subtotal row per group, and gives the block an Excel outline.
/// </summary>
/// <remarks>
/// <para>
/// The work splits in two because the two halves happen at different times.
/// <see cref="Prepare"/> runs before a single row exists: it evaluates each level's key for each
/// item and reorders the items, which has to happen while reordering still costs nothing.
/// <see cref="Render"/> runs once the rows have been written and evaluated, and turns the flat
/// block into a grouped one.
/// </para>
/// <para>
/// Rendering inserts sheet rows rather than rewriting the block, for the same reason expansion
/// does: the core library already shifts content, formulas, defined names and conditional formats
/// correctly, and a second implementation of that here would only be a worse one. The insertions
/// run bottom-up so that a row number worked out from the layout is still valid when its turn
/// comes.
/// </para>
/// </remarks>
internal sealed class GroupRenderer
{
    private readonly List<GroupLevel> _levels;
    private readonly IReadOnlyList<object?> _items;

    /// <summary>Each level's key for each item, as <c>[level][item]</c>.</summary>
    private readonly object?[][] _keys;

    private GroupRenderer(List<GroupLevel> levels, IReadOnlyList<object?> items, object?[][] keys)
    {
        _levels = levels;
        _items = items;
        _keys = keys;
    }

    /// <summary>
    /// The items in the order they should be written — grouped so that each group is contiguous.
    /// </summary>
    public IReadOnlyList<object?> Items => _items;

    /// <summary>
    /// Reads the range's group levels and orders its items by them, or returns <c>null</c> when the
    /// template asked for no grouping or asked for something that could not be evaluated.
    /// </summary>
    public static GroupRenderer? Prepare(
        IReadOnlyList<OptionTag> tags,
        IReadOnlyList<object?> items,
        GroupOptions options,
        ProcessingContext context)
    {
        var groups = tags.OfType<GroupTag>().ToList();
        if (groups.Count == 0)
        {
            return null;
        }

        if (context.IsHorizontal)
        {
            // Excel outlines columns as well as rows, but a subtotal *column* labelled with a group
            // key is not a thing report readers ask for, and pretending otherwise would mean
            // guessing at half the layout. Said plainly rather than half-done.
            context.Errors.Add(new TemplateError(
                "<<Group>> is not supported in a range that repeats across. Group a range that repeats downwards, "
                + "or summarise with <<Sum>> and friends instead.",
                context.Worksheet.Name));
            return null;
        }

        var levels = BuildLevels(groups, options, context);
        if (levels.Count == 0)
        {
            return null;
        }

        var keys = EvaluateKeys(levels, items, context);
        if (keys is null)
        {
            // A key that will not evaluate would put rows in arbitrary groups. The error is
            // recorded and the range renders ungrouped, which is the recoverable outcome.
            return null;
        }

        var (ordered, orderedKeys) = Order(levels, items, keys, context.Engine.Culture);
        return new GroupRenderer(levels, ordered, orderedKeys);
    }

    /// <summary>
    /// Groups the block that was generated for <see cref="Items"/> and returns its new last row.
    /// </summary>
    /// <param name="sheet">The sheet holding the generated block.</param>
    /// <param name="area">The range's columns, and the rows it occupied as a template.</param>
    /// <param name="dataFirstRow">The first generated row.</param>
    /// <param name="dataRowCount">How many rows one item occupies.</param>
    /// <param name="hasOptionsRow">Whether an options row follows the block.</param>
    /// <param name="tags">The range's tags, of which the summaries are repeated per group.</param>
    /// <param name="context">Where a summary reports a problem, and how it resolves its column.</param>
    public int Render(
        IXLWorksheet sheet,
        RangeArea area,
        int dataFirstRow,
        int dataRowCount,
        bool hasOptionsRow,
        IReadOnlyList<OptionTag> tags,
        ProcessingContext context)
    {
        var itemCount = _items.Count;
        var runs = BuildRuns(itemCount);
        var runOfItem = IndexRuns(runs, itemCount);
        var subtotalled = runs.Where(run => !_levels[run.Level].NoSubtotals).ToList();

        var (headers, footers) = PlaceSubtotalRows(subtotalled, itemCount);
        var itemFirstRow = AssignRows(headers, footers, runOfItem, dataFirstRow, dataRowCount, itemCount, out var lastRow);

        InsertSubtotalRows(sheet, headers, footers, dataFirstRow, dataRowCount, itemCount);

        var optionsRowNumber = hasOptionsRow ? lastRow + 1 : (int?)null;
        WriteSubtotalRows(sheet, area, subtotalled, tags, context, optionsRowNumber);
        Outline(sheet, runs, subtotalled, itemFirstRow, dataRowCount);
        MergeLabels(sheet, runs);
        AddPageBreaks(sheet, runs, lastRow);

        return lastRow;
    }

    private static List<GroupLevel> BuildLevels(List<GroupTag> groups, GroupOptions options, ProcessingContext context)
    {
        var levels = new List<GroupLevel>(groups.Count);

        foreach (var tag in groups)
        {
            var expression = tag.KeyExpression;

            if (expression.Length == 0 && !context.LineExpressions.TryGetValue(tag.Line, out expression!))
            {
                context.Errors.Add(new TemplateError(
                    "<<Group>> has nothing to group by: the column it is under holds no expression, and no 'by' was given.",
                    context.Worksheet.Name));
                continue;
            }

            levels.Add(new GroupLevel(
                tag.Column,
                expression,
                tag.Descending,
                tag.KeepOrder,
                tag.MergeLabels || options.MergeLabels,
                tag.SummaryAbove || options.SummaryAbove,
                tag.PageBreaks || options.PageBreaks,
                tag.Collapse || options.Collapse,
                tag.NoSubtotals || options.NoSubtotals,
                tag.TotalLabel));
        }

        return levels;
    }

    private static object?[][]? EvaluateKeys(
        List<GroupLevel> levels,
        IReadOnlyList<object?> items,
        ProcessingContext context)
    {
        var keys = new object?[levels.Count][];

        for (var level = 0; level < levels.Count; level++)
        {
            var lane = new object?[items.Count];

            for (var i = 0; i < items.Count; i++)
            {
                try
                {
                    lane[i] = context.Engine.Evaluate(levels[level].Expression, ItemScopes.For(context.Scope, items, i));
                }
                catch (ExpressionEvaluationException ex)
                {
                    context.Errors.Add(new TemplateError(ex.Message, context.Worksheet.Name, exception: ex));
                    return null;
                }
            }

            keys[level] = lane;
        }

        return keys;
    }

    /// <summary>
    /// Orders the items so that every group is contiguous, outermost level first.
    /// </summary>
    /// <remarks>
    /// One ordering over all the levels at once, rather than one per level in turn: chaining stable
    /// sorts would make the <em>last</em> level sorted the primary key, which is the opposite of
    /// how the levels nest. Because the ordering is stable, whatever <c>&lt;&lt;Sort&gt;&gt;</c>
    /// already decided survives as the order within each group.
    /// </remarks>
#pragma warning disable S3776 // Building one ordering across all levels; the first level seeds it and the rest chain
    private static (IReadOnlyList<object?> Items, object?[][] Keys) Order(
        List<GroupLevel> levels,
        IReadOnlyList<object?> items,
        object?[][] keys,
        CultureInfo culture)
    {
        var comparer = new SortKeyComparer(culture);
        IOrderedEnumerable<int>? ordered = null;

        for (var level = 0; level < levels.Count; level++)
        {
            if (levels[level].KeepOrder)
            {
                continue;
            }

            var lane = keys[level];
            var descending = levels[level].Descending;
            Func<int, object?> key = i => lane[i];

            if (ordered is null)
            {
                ordered = descending
                    ? Enumerable.Range(0, items.Count).OrderByDescending(key, comparer)
                    : Enumerable.Range(0, items.Count).OrderBy(key, comparer);
            }
            else
            {
                ordered = descending
                    ? ordered.ThenByDescending(key, comparer)
                    : ordered.ThenBy(key, comparer);
            }
        }

        if (ordered is null)
        {
            return (items, keys);
        }

        var order = ordered.ToArray();
        var orderedItems = new object?[items.Count];
        var orderedKeys = new object?[keys.Length][];

        for (var level = 0; level < keys.Length; level++)
        {
            orderedKeys[level] = new object?[items.Count];
        }

        for (var position = 0; position < order.Length; position++)
        {
            orderedItems[position] = items[order[position]];

            for (var level = 0; level < keys.Length; level++)
            {
                orderedKeys[level][position] = keys[level][order[position]];
            }
        }

        return (orderedItems, orderedKeys);
    }
#pragma warning restore S3776

    /// <summary>
    /// Finds each level's runs of consecutive items sharing a key — a run being one group.
    /// </summary>
    /// <remarks>
    /// An inner level's run also ends where an outer key changes, so that two groups that happen to
    /// share an inner key across an outer boundary stay two groups.
    /// </remarks>
    private List<Run> BuildRuns(int itemCount)
    {
        var runs = new List<Run>();

        for (var level = 0; level < _levels.Count; level++)
        {
            var start = 0;

            for (var i = 1; i <= itemCount; i++)
            {
                if (i < itemCount && SameThroughLevel(level, i - 1, i))
                {
                    continue;
                }

                runs.Add(new Run(level, start, i - 1, _keys[level][start]));
                start = i;
            }
        }

        return runs;
    }

    private bool SameThroughLevel(int level, int left, int right)
    {
        for (var l = 0; l <= level; l++)
        {
            if (!Equals(_keys[l][left], _keys[l][right]))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>Which run covers each item, at each level.</summary>
    private Run[][] IndexRuns(List<Run> runs, int itemCount)
    {
        var index = new Run[_levels.Count][];

        for (var level = 0; level < _levels.Count; level++)
        {
            index[level] = new Run[itemCount];
        }

        foreach (var run in runs)
        {
            for (var i = run.FirstItem; i <= run.LastItem; i++)
            {
                index[run.Level][i] = run;
            }
        }

        return index;
    }

    /// <summary>
    /// Decides which item each subtotal row attaches to, and in what order several of them stack.
    /// </summary>
    /// <remarks>
    /// Below a group, the innermost total comes first and the outermost last, so the totals read
    /// outwards from the rows they cover. Above a group, that reverses.
    /// </remarks>
    private (List<Run>?[] Headers, List<Run>?[] Footers) PlaceSubtotalRows(List<Run> subtotalled, int itemCount)
    {
        var headers = new List<Run>?[itemCount];
        var footers = new List<Run>?[itemCount];

        foreach (var run in subtotalled)
        {
            if (_levels[run.Level].SummaryAbove)
            {
                (headers[run.FirstItem] ??= new List<Run>()).Add(run);
            }
            else
            {
                (footers[run.LastItem] ??= new List<Run>()).Add(run);
            }
        }

        // Runs arrive outermost level first, which is the order headers want; footers want the
        // reverse.
        foreach (var list in footers)
        {
            list?.Reverse();
        }

        return (headers, footers);
    }

    /// <summary>
    /// Walks the layout the block is about to have and records where everything lands.
    /// </summary>
    private int[] AssignRows(
        List<Run>?[] headers,
        List<Run>?[] footers,
        Run[][] runOfItem,
        int dataFirstRow,
        int dataRowCount,
        int itemCount,
        out int lastRow)
    {
        var itemFirstRow = new int[itemCount];
        var row = dataFirstRow;

        for (var i = 0; i < itemCount; i++)
        {
            if (headers[i] is { } above)
            {
                foreach (var run in above)
                {
                    run.SubtotalRow = row;
                    ExtendAncestors(runOfItem, run, i, row);
                    row++;
                }
            }

            itemFirstRow[i] = row;

            for (var level = 0; level < _levels.Count; level++)
            {
                Extend(runOfItem[level][i], row);
                Extend(runOfItem[level][i], row + dataRowCount - 1);
            }

            row += dataRowCount;

            if (footers[i] is { } below)
            {
                foreach (var run in below)
                {
                    run.SubtotalRow = row;
                    ExtendAncestors(runOfItem, run, i, row);
                    row++;
                }
            }
        }

        lastRow = row - 1;
        return itemFirstRow;
    }

    /// <summary>
    /// Grows the runs that enclose <paramref name="run"/> to cover <paramref name="row"/>. A run
    /// never covers its own subtotal row, so that the row is not summarised by the formula it holds.
    /// </summary>
    private static void ExtendAncestors(Run[][] runOfItem, Run run, int item, int row)
    {
        for (var level = 0; level < run.Level; level++)
        {
            Extend(runOfItem[level][item], row);
        }
    }

    private static void Extend(Run? run, int row)
    {
        if (run is null)
        {
            return;
        }

        if (row < run.ContentFirstRow)
        {
            run.ContentFirstRow = row;
        }

        if (row > run.ContentLastRow)
        {
            run.ContentLastRow = row;
        }
    }

    /// <summary>
    /// Makes room for the subtotal rows, working up the sheet so that the row numbers worked out
    /// for the rows above are still the row numbers they will have.
    /// </summary>
    private static void InsertSubtotalRows(
        IXLWorksheet sheet,
        List<Run>?[] headers,
        List<Run>?[] footers,
        int dataFirstRow,
        int dataRowCount,
        int itemCount)
    {
        var insertBefore = new int[itemCount + 1];

        for (var i = 0; i < itemCount; i++)
        {
            insertBefore[i] += headers[i]?.Count ?? 0;
            insertBefore[i + 1] += footers[i]?.Count ?? 0;
        }

        for (var i = itemCount; i >= 0; i--)
        {
            var count = insertBefore[i];
            if (count == 0)
            {
                continue;
            }

            if (i < itemCount)
            {
                sheet.Row(dataFirstRow + (i * dataRowCount)).InsertRowsAbove(count);
            }
            else
            {
                sheet.Row(dataFirstRow + (itemCount * dataRowCount) - 1).InsertRowsBelow(count);
            }
        }
    }

    /// <summary>
    /// Fills each subtotal row: the options row's styling, then its summaries over the group's rows,
    /// then the group's label.
    /// </summary>
    /// <remarks>
    /// Taking the options row's styling is what makes a subtotal look like one. A template author
    /// styles the grand total once — bold, a rule above it — and every group total matches it,
    /// which is both what a reader expects and the only styling the template has a way to express.
    /// The label goes in the grouped column unless a summary already claims it.
    /// </remarks>
    private void WriteSubtotalRows(
        IXLWorksheet sheet,
        RangeArea area,
        List<Run> subtotalled,
        IReadOnlyList<OptionTag> tags,
        ProcessingContext context,
        int? optionsRowNumber)
    {
        if (subtotalled.Count == 0)
        {
            return;
        }

        var summaries = tags.Where(tag => tag is IRangeSummaryTag).ToList();
        var summaryColumns = new HashSet<int>(summaries.Select(tag => tag.Column));

        foreach (var run in subtotalled)
        {
            if (optionsRowNumber is { } styleRow)
            {
                for (var column = area.FirstColumn; column <= area.LastColumn; column++)
                {
                    sheet.Cell(run.SubtotalRow, column).Style = sheet.Cell(styleRow, column).Style;
                }
            }

            foreach (var tag in summaries)
            {
                ((IRangeSummaryTag)tag).TryWriteSummary(
                    sheet.Cell(run.SubtotalRow, tag.Column),
                    run.ContentFirstRow,
                    run.ContentLastRow,
                    context);
            }

            var level = _levels[run.Level];
            if (!summaryColumns.Contains(level.Column))
            {
                sheet.Cell(run.SubtotalRow, level.Column).Value =
                    level.TotalLabel.Replace("{0}", KeyText.For(run.Key, context.Engine.Culture));
            }
        }
    }

    /// <summary>
    /// Gives the block its outline: data rows at the innermost level, each subtotal row one level
    /// out from the rows it covers, so that collapsing a level in Excel leaves its totals showing.
    /// </summary>
#pragma warning disable S3776 // Outline levels, collapsing and summary position are three independent passes
    private void Outline(IXLWorksheet sheet, List<Run> runs, List<Run> subtotalled, int[] itemFirstRow, int dataRowCount)
    {
        // Excel allows eight outline levels. A template with more still generates; it just stops
        // nesting deeper.
        var depth = Math.Min(_levels.Count, 8);

        foreach (var firstRow in itemFirstRow)
        {
            for (var row = firstRow; row < firstRow + dataRowCount; row++)
            {
                sheet.Row(row).OutlineLevel = depth;
            }
        }

        foreach (var run in subtotalled)
        {
            sheet.Row(run.SubtotalRow).OutlineLevel = Math.Min(run.Level, 8);
        }

        foreach (var run in runs)
        {
            if (!_levels[run.Level].Collapse)
            {
                continue;
            }

            for (var row = run.ContentFirstRow; row <= run.ContentLastRow; row++)
            {
                if (sheet.Row(row).OutlineLevel > run.Level)
                {
                    sheet.Row(row).Hide();
                }
            }

            if (run.SubtotalRow > 0)
            {
                sheet.Row(run.SubtotalRow).Group(Math.Min(run.Level, 8), collapse: true);
            }
        }

        if (subtotalled.Count > 0)
        {
            sheet.Outline.SummaryVLocation = subtotalled.TrueForAll(run => _levels[run.Level].SummaryAbove)
                ? XLOutlineSummaryVLocation.Top
                : XLOutlineSummaryVLocation.Bottom;
        }
    }
#pragma warning restore S3776

    /// <summary>
    /// Merges a group's repeated label cells into one, which is the point of grouping a column the
    /// report also displays.
    /// </summary>
    private void MergeLabels(IXLWorksheet sheet, List<Run> runs)
    {
        foreach (var run in runs)
        {
            var level = _levels[run.Level];

            if (!level.MergeLabels || run.ContentLastRow <= run.ContentFirstRow)
            {
                continue;
            }

            var label = sheet.Range(run.ContentFirstRow, level.Column, run.ContentLastRow, level.Column);
            label.Merge();

            // A merged cell is bottom-aligned by default, which reads as if the label belonged to
            // the last row of the group rather than to all of them.
            label.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }
    }

    private void AddPageBreaks(IXLWorksheet sheet, List<Run> runs, int lastRow)
    {
        foreach (var run in runs)
        {
            if (!_levels[run.Level].PageBreaks)
            {
                continue;
            }

            var breakRow = Math.Max(run.ContentLastRow, run.SubtotalRow);

            // No break after the last group: it would leave a page holding nothing but the totals.
            if (breakRow < lastRow)
            {
                sheet.PageSetup.AddHorizontalPageBreak(breakRow);
            }
        }
    }

    /// <summary>One <c>&lt;&lt;Group&gt;&gt;</c> tag, resolved against the range-wide options.</summary>
    private sealed record GroupLevel(
        int Column,
        string Expression,
        bool Descending,
        bool KeepOrder,
        bool MergeLabels,
        bool SummaryAbove,
        bool PageBreaks,
        bool Collapse,
        bool NoSubtotals,
        string TotalLabel);

    /// <summary>
    /// One group: the items sharing a key at one level, and where they ended up on the sheet.
    /// </summary>
    private sealed class Run
    {
        public Run(int level, int firstItem, int lastItem, object? key)
        {
            Level = level;
            FirstItem = firstItem;
            LastItem = lastItem;
            Key = key;
        }

        public int Level { get; }

        public int FirstItem { get; }

        public int LastItem { get; }

        public object? Key { get; }

        /// <summary>The group's subtotal row, or zero if the level has none.</summary>
        public int SubtotalRow { get; set; }

        /// <summary>
        /// The first row the group covers, excluding its own subtotal row. Rows a nested group
        /// added are included, because they are part of what this group totals.
        /// </summary>
        public int ContentFirstRow { get; set; } = int.MaxValue;

        /// <summary>The last row the group covers, excluding its own subtotal row.</summary>
        public int ContentLastRow { get; set; }
    }
}
