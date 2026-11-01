using System.Collections.Generic;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// The value of a <c>LAMBDA</c>: a function that a formula can call, bind to a <c>LET</c> name, or keep
/// in a defined name (#847, spec 08 phase 2).
/// </summary>
/// <remarks>
/// A lambda is lexically scoped. It keeps the names bound where it was written, so
/// <c>LET(y,10,f,LAMBDA(x,x+y),LET(y,100,f(1)))</c> is 11, as in Excel, whatever the caller binds.
/// </remarks>
internal sealed class LambdaFunction
{
    internal LambdaFunction(IReadOnlyList<string> parameters, ValueNode body, LetScope<AnyValue>? scope)
    {
        Parameters = parameters;
        Body = body;
        Scope = scope;
    }

    /// <summary>The names of the parameters, as the formula writes them.</summary>
    internal IReadOnlyList<string> Parameters { get; }

    /// <summary>The expression a call evaluates, with the parameters bound to the arguments.</summary>
    internal ValueNode Body { get; }

    /// <summary>The names bound where the <c>LAMBDA</c> was written, which its body sees.</summary>
    internal LetScope<AnyValue>? Scope { get; }
}
