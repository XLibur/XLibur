using System.Collections.Generic;

namespace XLibur.Report.Expressions;

/// <summary>
/// Builds the scope one row of a bound range is evaluated in.
/// </summary>
/// <remarks>
/// Three things go into scope for a row: <c>item</c>, its zero-based <c>index</c>, and the whole
/// <c>items</c> collection so a row can refer to its neighbours or to a total. Expansion, sorting
/// and grouping all evaluate expressions per row and all have to agree on those names, so they
/// share this one factory rather than each spelling the trio out.
/// </remarks>
internal static class ItemScopes
{
    public static ExpressionScope For(ExpressionScope parent, IReadOnlyList<object?> items, int index) =>
        parent.CreateChild(new[]
        {
            new KeyValuePair<string, object?>("item", items[index]),
            new KeyValuePair<string, object?>("index", index),
            new KeyValuePair<string, object?>("items", items),
        });
}
