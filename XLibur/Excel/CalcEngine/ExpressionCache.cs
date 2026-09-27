using System.Runtime.CompilerServices;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// Caches expressions based on their string representation.
/// This saves parsing time.
/// </summary>
/// <remarks>
/// <para>Uses weak references to avoid accumulating unused expressions.</para>
/// <para>
/// Only for text that belongs to no cell: <c>Evaluate</c> calls and defined names. A cell's formula
/// keeps its own tree (<see cref="XLCellFormula.GetAst"/>), because its text lives as long as the
/// formula does, and a weak entry keyed by it kept every tree alive anyway (#686).
/// </para>
/// </remarks>
internal sealed class ExpressionCache
{
    private readonly ConditionalWeakTable<string, Formula> _cache;
    private readonly XLCalcEngine _ce;

    public ExpressionCache(XLCalcEngine ce)
    {
        _ce = ce;
        _cache = new ConditionalWeakTable<string, Formula>();
    }

    // gets the parsed version of a string expression
    public Formula this[string expression]
    {
        get
        {
            if (!_cache.TryGetValue(expression, out var formula))
            {
                formula = _ce.Parse(expression);
                _cache.Add(expression, formula);
            }
            return formula;
        }
    }
}
