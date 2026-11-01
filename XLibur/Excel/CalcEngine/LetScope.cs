using System;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// The names the <c>LET</c> calls around the node being visited have bound, innermost first: one
/// link per name, pointing at the names bound before it (#847, spec 08 phase 1).
/// </summary>
/// <typeparam name="TValue">
/// What a name stands for: its value for <see cref="CalculationVisitor"/>, the areas it refers to for
/// <see cref="DependenciesVisitor"/>.
/// </typeparam>
/// <remarks>
/// <para>
/// A link is never changed, so a <c>LET</c> leaves its scope by putting back the link it started
/// from, and an inner binding of a name hides an outer one only while the inner <c>LET</c> is
/// visited. A name is looked up in the scope before the defined names, as in Excel.
/// </para>
/// <para>
/// Names compare in any case, without the <c>_xlpm.</c> prefix a file writes them with, so a formula
/// loaded from a file and the same formula typed through the API bind alike.
/// </para>
/// </remarks>
internal sealed class LetScope<TValue>
{
    private const string ParameterPrefix = "_xlpm.";

    private readonly string _name;
    private readonly TValue _value;
    private readonly LetScope<TValue>? _outer;

    internal LetScope(string name, TValue value, LetScope<TValue>? outer)
    {
        _name = BareName(name);
        _value = value;
        _outer = outer;
    }

    /// <summary>Finds the value the innermost binding of <paramref name="name"/> gives it.</summary>
    internal bool TryGetValue(string name, out TValue value)
    {
        var bare = BareName(name);
        for (var scope = this; scope is not null; scope = scope._outer)
        {
            if (string.Equals(scope._name, bare, StringComparison.OrdinalIgnoreCase))
            {
                value = scope._value;
                return true;
            }
        }

        value = default!;
        return false;
    }

    private static string BareName(string name) =>
        name.StartsWith(ParameterPrefix, StringComparison.OrdinalIgnoreCase) ? name[ParameterPrefix.Length..] : name;
}
