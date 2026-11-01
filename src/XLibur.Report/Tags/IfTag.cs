using System;
using System.Collections.Generic;
using XLibur.Report.Expressions;

namespace XLibur.Report.Tags;

/// <summary>
/// Leaves rows out of a report. Written as <c>&lt;&lt;If test="…"&gt;&gt;</c>.
/// </summary>
/// <remarks>
/// <para>
/// Where it is written decides what it drops. In one of the range's <strong>repeated rows</strong> it
/// asks the question of each item and keeps the ones that answer yes:
/// <c>&lt;&lt;If test="item.Quantity &gt; 0"&gt;&gt;</c> beside <c>{{ item.Product }}</c> leaves the
/// empty lines out. In the <strong>options row</strong> it asks once, of the range as a whole, and a
/// no renders the range as an empty collection does — the rows gone, the headings and any
/// options-row total behaving exactly as they would over no data.
/// </para>
/// <para>
/// The test runs before any row is written, so what survives it is what everything else sees: sorting
/// orders the survivors, grouping groups them, and a total totals them. Nothing has to know that
/// anything was dropped.
/// </para>
/// <para>
/// Only <c>null</c> and <c>false</c> are false. <strong>Zero is true</strong>, so
/// <c>test="item.Quantity"</c> keeps every row including the empty ones; write the comparison out.
/// </para>
/// </remarks>
public sealed class IfTag : OptionTag
{
    /// <inheritdoc />
    public override IReadOnlyList<object?> TransformItems(IReadOnlyList<object?> items, ProcessingContext context)
    {
        var test = Token.Value("test");

        if (test.Length == 0)
        {
            context.Errors.Add(new TemplateError(
                "<<If>> needs something to test: <<If test=\"item.Quantity > 0\">>.",
                context.Worksheet.Name));
            return items;
        }

        if (!InRepeatedRow)
        {
            // Asked once, of the range. `items` goes into scope so the question can be about the
            // collection — <<If test="items.size > 3">> — and not only about the workbook variables.
            return context.IsTrue(test, context.Scope.CreateChild("items", items))
                ? items
                : Array.Empty<object?>();
        }

        var kept = new List<object?>(items.Count);

        for (var i = 0; i < items.Count; i++)
        {
            if (context.IsTrue(test, ItemScopes.For(context.Scope, items, i)))
            {
                kept.Add(items[i]);
            }
        }

        return kept;
    }
}
