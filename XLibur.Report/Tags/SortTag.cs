using System;
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

        if (expression.Length == 0 && !context.ColumnExpressions.TryGetValue(Column, out expression!))
        {
            context.Errors.Add(new TemplateError(
                $"<<{Token.Name}>> has nothing to sort by: the column it is under holds no expression, and no 'by' was given.",
                context.Worksheet.Name));
            return items;
        }

        var keyed = new List<(object? Item, object? Key)>(items.Count);

        for (var i = 0; i < items.Count; i++)
        {
            try
            {
                keyed.Add((items[i], context.Engine.Evaluate(expression, ItemScope(context, items, i))));
            }
            catch (ExpressionEvaluationException ex)
            {
                context.Errors.Add(new TemplateError(ex.Message, context.Worksheet.Name, exception: ex));
                return items;
            }
        }

        var comparer = SortKeyComparer.Instance;
        var sorted = Descending
            ? keyed.OrderByDescending(pair => pair.Key, comparer)
            : keyed.OrderBy(pair => pair.Key, comparer);

        return sorted.Select(pair => pair.Item).ToList();
    }

    private static ExpressionScope ItemScope(ProcessingContext context, IReadOnlyList<object?> items, int index) =>
        context.Scope.CreateChild(new[]
        {
            new KeyValuePair<string, object?>("item", items[index]),
            new KeyValuePair<string, object?>("index", index),
            new KeyValuePair<string, object?>("items", items),
        });

    /// <summary>
    /// Orders values of whatever type the expression produced, keeping blanks together at the end.
    /// </summary>
    private sealed class SortKeyComparer : IComparer<object?>
    {
        public static readonly SortKeyComparer Instance = new();

        public int Compare(object? x, object? y)
        {
            if (x is null)
            {
                return y is null ? 0 : 1;
            }

            if (y is null)
            {
                return -1;
            }

            if (x is IComparable comparable && x.GetType() == y.GetType())
            {
                return comparable.CompareTo(y);
            }

            // Different types: compare numerically when both are numbers, textually otherwise.
            if (IsNumeric(x) && IsNumeric(y))
            {
                return Convert.ToDouble(x, System.Globalization.CultureInfo.InvariantCulture)
                    .CompareTo(Convert.ToDouble(y, System.Globalization.CultureInfo.InvariantCulture));
            }

            return string.Compare(x.ToString(), y.ToString(), StringComparison.CurrentCulture);
        }

        private static bool IsNumeric(object value) =>
            value is sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal;
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
