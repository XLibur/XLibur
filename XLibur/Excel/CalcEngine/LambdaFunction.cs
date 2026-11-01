using System.Collections.Generic;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// The value of a <c>LAMBDA</c>: a function that a formula can call, bind to a <c>LET</c> name, keep in
/// a defined name, or pass to a function that takes one, such as <c>MAP</c> (#847, spec 08). One of
/// Excel's functions written by its name as a value, <c>BYROW(D1:E2,SUM)</c>, is one too.
/// </summary>
/// <remarks>
/// A lambda is lexically scoped. It keeps the names bound where it was written, so
/// <c>LET(y,10,f,LAMBDA(x,x+y),LET(y,100,f(1)))</c> is 11, as in Excel, whatever the caller binds.
/// </remarks>
internal sealed class LambdaFunction
{
    internal LambdaFunction(IReadOnlyList<string> parameters, int requiredParameterCount, ValueNode body,
        LetScope<AnyValue>? scope)
    {
        Parameters = parameters;
        RequiredParameterCount = requiredParameterCount;
        Body = body;
        Scope = scope;
    }

    /// <summary>One of Excel's functions as a value. Excel writes it <c>_xleta.SUM</c>.</summary>
    internal LambdaFunction(FunctionDefinition function)
    {
        Parameters = [];
        Function = function;
    }

    /// <summary>The names of the parameters, as the formula writes them.</summary>
    internal IReadOnlyList<string> Parameters { get; }

    /// <summary>
    /// How many parameters a call must pass. The others are optional, typed <c>[y]</c> and written
    /// <c>_xlop.y</c>, and a call may leave them out, which <c>ISOMITTED</c> tells.
    /// </summary>
    internal int RequiredParameterCount { get; }

    /// <summary>The expression a call evaluates, with the parameters bound to the arguments.</summary>
    internal ValueNode? Body { get; }

    /// <summary>The names bound where the <c>LAMBDA</c> was written, which its body sees.</summary>
    internal LetScope<AnyValue>? Scope { get; }

    /// <summary>The function, for one of Excel's functions written by its name as a value.</summary>
    internal FunctionDefinition? Function { get; }
}
