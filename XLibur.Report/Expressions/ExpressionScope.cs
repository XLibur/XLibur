using System;
using System.Collections.Generic;

namespace XLibur.Report.Expressions;

/// <summary>
/// The named values visible to one expression evaluation, as a chain of nested scopes.
/// </summary>
/// <remarks>
/// A template evaluates expressions at several nesting levels: workbook-wide variables at the
/// outside, then a bound range's collection, then the current row's item. Rather than flattening
/// those into one dictionary per cell, each level is its own scope linked to its parent, so a row
/// scope is cheap to create and discard. Inner scopes shadow outer ones.
/// <para>
/// Names are matched with <see cref="StringComparer.Ordinal"/>, matching C# member semantics —
/// <c>item.Price</c> and <c>item.price</c> are different names.
/// </para>
/// </remarks>
public sealed class ExpressionScope
{
    /// <summary>A scope with no values and no parent.</summary>
    public static readonly ExpressionScope Empty = new(Array.Empty<KeyValuePair<string, object?>>());

    private readonly Dictionary<string, object?> _values;

    /// <summary>
    /// Creates a scope holding <paramref name="values"/>, optionally nested inside
    /// <paramref name="parent"/>.
    /// </summary>
    public ExpressionScope(IEnumerable<KeyValuePair<string, object?>> values, ExpressionScope? parent = null)
    {
        if (values is null)
        {
            throw new ArgumentNullException(nameof(values));
        }

        _values = new Dictionary<string, object?>(StringComparer.Ordinal);
        foreach (var pair in values)
        {
            _values[pair.Key] = pair.Value;
        }

        Parent = parent;
    }

    /// <summary>The enclosing scope, or <c>null</c> for the outermost one.</summary>
    public ExpressionScope? Parent { get; }

    /// <summary>The values declared by this scope, excluding anything inherited from <see cref="Parent"/>.</summary>
    public IReadOnlyDictionary<string, object?> Values => _values;

    /// <summary>Creates a scope nested inside this one.</summary>
    public ExpressionScope CreateChild(IEnumerable<KeyValuePair<string, object?>> values) => new(values, this);

    /// <summary>Creates a scope nested inside this one, holding a single value.</summary>
    public ExpressionScope CreateChild(string name, object? value) =>
        CreateChild(new[] { new KeyValuePair<string, object?>(name, value) });

    /// <summary>
    /// Looks a name up through the scope chain, innermost first.
    /// </summary>
    public bool TryGetValue(string name, out object? value)
    {
        for (var scope = this; scope is not null; scope = scope.Parent)
        {
            if (scope._values.TryGetValue(name, out value))
            {
                return true;
            }
        }

        value = null;
        return false;
    }

    /// <summary>
    /// The scope chain ordered outermost first, which is the order engines need in order to let
    /// inner scopes shadow outer ones.
    /// </summary>
    public IReadOnlyList<ExpressionScope> FromOutermost()
    {
        var chain = new List<ExpressionScope>();
        for (var scope = this; scope is not null; scope = scope.Parent)
        {
            chain.Add(scope);
        }

        chain.Reverse();
        return chain;
    }
}
