using System;
using System.Collections.Generic;

namespace XLibur.Excel.CalcEngine;

/// <summary>Which parameters of a function allow ranges? That is important for implicit intersection.</summary>
internal enum AllowRange
{
    /// <summary>None of the parameters allow ranges.</summary>
    None,

    /// <summary>All parameters allow ranges.</summary>
    All,

    /// <summary>All parameters except marked ones allow ranges.</summary>
    Except,

    /// <summary>Only marked parameters allow ranges.</summary>
    Only,
}

internal sealed class FunctionRegistry
{
    private readonly Dictionary<string, FunctionDefinition> _func = new(StringComparer.InvariantCultureIgnoreCase);

    /// <summary>
    /// Names of every registered function. Used by XLibur.Report to surface the function library
    /// inside template expressions without maintaining its own list of what exists.
    /// </summary>
    public IEnumerable<string> Names => _func.Keys;

    /// <summary>
    /// The name of a function as its node keeps it. For a registered function that is the
    /// registry's own string, so a parse does not copy the name out of the formula text (#686). The
    /// case can differ from the text; every reader of the name ignores case.
    /// </summary>
    /// <remarks>
    /// Finding a string key by a span needs <c>GetAlternateLookup</c>, which .NET 9 added. On
    /// .NET 8 the name is always copied.
    /// </remarks>
    public string GetName(ReadOnlySpan<char> name)
    {
#if NET9_0_OR_GREATER
        if (_func.GetAlternateLookup<ReadOnlySpan<char>>().TryGetValue(name, out var registeredName, out _))
            return registeredName;
#endif
        return name.ToString();
    }

    public bool TryGetFunc(string name, out FunctionDefinition? func)
    {
        return _func.TryGetValue(name, out func);
    }

    public bool TryGetFunc(string name, out int paramMin, out int paramMax)
    {
        if (_func.TryGetValue(name, out var func))
        {
            paramMin = func.MinParams;
            paramMax = func.MaxParams;
            return true;
        }

        paramMin = -1;
        paramMax = -1;
        return false;
    }

    /// <summary>
    /// Add a function to the registry.
    /// </summary>
    /// <param name="functionName">Name of function in formulas.</param>
    /// <param name="minParams">Minimum number of parameters.</param>
    /// <param name="maxParams">Maximum number of parameters.</param>
    /// <param name="fn">A delegate of a function that will be called when the function is supposed to be evaluated.</param>
    /// <param name="flags">Flags that indicate some additional info about a function.</param>
    /// <param name="allowRanges">Which parameters allow ranges to be argument? Useful for array formulas.</param>
    /// <param name="markedParams">Index of parameter that is marked, start from 0</param>
    public void RegisterFunction(string functionName, int minParams, int maxParams, CalcEngineFunction fn,
        FunctionFlags flags, AllowRange allowRanges = AllowRange.None, params int[] markedParams)
    {
        _func.Add(functionName, new FunctionDefinition(minParams, maxParams, fn, flags, allowRanges, markedParams));
    }
}
