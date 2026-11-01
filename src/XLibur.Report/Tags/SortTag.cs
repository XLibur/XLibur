using System.Collections.Generic;
using System.Linq;
using XLibur.Report.Expressions;

namespace XLibur.Report.Tags;

/// <summary>
/// Orders a range's rows. Written in the options row under the column to sort by —
/// <c>&lt;&lt;Sort&gt;&gt;</c>, or <c>&lt;&lt;Sort desc&gt;&gt;</c> to reverse it.
/// </summary>
/// <remarks>
/// The sort key is the template expression of the column the tag sits under, so a column already
/// showing <c>{{ item.SoldOn }}</c> needs no second mention of it. Give <c>by</c> to sort by
/// something the range does not display: <c>&lt;&lt;Sort by="item.Customer.Name"&gt;&gt;</c>.
/// </remarks>
public class SortTag : OptionTag
{
    /// <summary>Whether rows are ordered largest first.</summary>
    protected virtual bool Descending => Token.Flag("desc");

    /// <inheritdoc />
    public override IReadOnlyList<object?> TransformItems(IReadOnlyList<object?> items, ProcessingContext context)
    {
        var expression = Token.Value("by");

        if (expression.Length == 0 && !context.LineExpressions.TryGetValue(Line, out expression!))
        {
            context.Errors.Add(new TemplateError(
                $"<<{Token.Name}>> has nothing to sort by: the line it sits in holds no expression, and no 'by' was given.",
                context.Worksheet.Name));
            return items;
        }

        var keyed = new List<(object? Item, object? Key)>(items.Count);

        for (var i = 0; i < items.Count; i++)
        {
            try
            {
                keyed.Add((items[i], context.Engine.Evaluate(expression, ItemScopes.For(context.Scope, items, i))));
            }
            catch (ExpressionEvaluationException ex)
            {
                context.Errors.Add(new TemplateError(ex.Message, context.Worksheet.Name, exception: ex));
                return items;
            }
        }

        var comparer = new SortKeyComparer(context.Engine.Culture);
        var sorted = Descending
            ? keyed.OrderByDescending(pair => pair.Key, comparer)
            : keyed.OrderBy(pair => pair.Key, comparer);

        return sorted.Select(pair => pair.Item).ToList();
    }
}

/// <summary>
/// Orders a range's rows largest first. Equivalent to <c>&lt;&lt;Sort desc&gt;&gt;</c>.
/// </summary>
public sealed class DescendingSortTag : SortTag
{
    /// <inheritdoc />
    protected override bool Descending => true;
}
