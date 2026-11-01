using System;
using System.Buffers;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Threading;
using XLibur.Parser;
using XLibur.Excel.CalcEngine.Exceptions;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Tables;

namespace XLibur.Excel.CalcEngine;

internal sealed class CalculationVisitor : IFormulaVisitor<CalcContext, AnyValue>
{
    /// <summary>
    /// The most calls of a <c>LAMBDA</c> that can be evaluated one inside another. A deeper call is
    /// <c>#NUM!</c>, as in Excel (<see cref="LambdaDepthException"/>).
    /// </summary>
    internal const int MaxLambdaDepth = 1024;

    /// <summary>
    /// The stack of a thread a deep call goes on on (<see cref="OnFreshStack"/>): enough for every call
    /// <see cref="MaxLambdaDepth"/> allows, with room to spare.
    /// </summary>
    private const int FreshStackSize = 16 * 1024 * 1024;

    private readonly FunctionRegistry _functions;
    private readonly ArrayPool<AnyValue> _argsPool;
    private readonly FunctionDefinition? _if;

    public CalculationVisitor(FunctionRegistry functions)
    {
        _functions = functions;
        _argsPool = ArrayPool<AnyValue>.Create(XLConstants.MaxFunctionArguments, 100);
        functions.TryGetFunc("IF", out _if);
    }

    public AnyValue Visit(CalcContext context, ScalarNode node)
    {
        return node.Value.ToAnyValue();
    }

    public AnyValue Visit(CalcContext context, ArrayNode node)
    {
        return node.Value;
    }

    public AnyValue Visit(CalcContext context, UnaryNode node)
    {
        var arg = node.Expression.Accept(context, this);

        return node.Operation switch
        {
            UnaryOp.Add => arg.UnaryPlus(),
            UnaryOp.Subtract => arg.UnaryMinus(context),
            UnaryOp.Percentage => arg.UnaryPercent(context),
            UnaryOp.SpillRange => EvaluateSpillRange(context, arg),
            UnaryOp.ImplicitIntersection => EvaluateImplicitIntersection(context, arg),
            _ => throw new NotSupportedException($"Unknown operator {node.Operation}.")
        };
    }

    public AnyValue Visit(CalcContext context, BinaryNode node)
    {
        var leftArg = node.LeftExpression.Accept(context, this);
        var rightArg = node.RightExpression.Accept(context, this);

        return node.Operation switch
        {
            BinaryOp.Range => AnyValue.ReferenceRange(leftArg, rightArg, context),
            BinaryOp.Union => AnyValue.ReferenceUnion(leftArg, rightArg),
            BinaryOp.Intersection => AnyValue.ReferenceIntersection(leftArg, rightArg, context),
            BinaryOp.Concat => AnyValue.Concat(leftArg, rightArg, context),
            BinaryOp.Add => AnyValue.BinaryPlus(leftArg, rightArg, context),
            BinaryOp.Sub => AnyValue.BinaryMinus(leftArg, rightArg, context),
            BinaryOp.Mult => AnyValue.BinaryMult(leftArg, rightArg, context),
            BinaryOp.Div => AnyValue.BinaryDiv(leftArg, rightArg, context),
            BinaryOp.Exp => AnyValue.BinaryExp(leftArg, rightArg, context),
            BinaryOp.Lt => AnyValue.IsLessThan(leftArg, rightArg, context),
            BinaryOp.Lte => AnyValue.IsLessThanOrEqual(leftArg, rightArg, context),
            BinaryOp.Eq => AnyValue.IsEqual(leftArg, rightArg, context),
            BinaryOp.Neq => AnyValue.IsNotEqual(leftArg, rightArg, context),
            BinaryOp.Gte => AnyValue.IsGreaterThanOrEqual(leftArg, rightArg, context),
            BinaryOp.Gt => AnyValue.IsGreaterThan(leftArg, rightArg, context),
            _ => throw new NotSupportedException($"Unknown operator {node.Operation}.")
        };
    }

    public AnyValue Visit(CalcContext context, FunctionNode node)
    {
        // LET and LAMBDA are special forms: their arguments are names, and values evaluated in a
        // scope of those names, which no registered function can be, as a function's arguments are
        // all evaluated before it runs.
        if (node.Prefix is null && string.Equals(node.Name, "LET", StringComparison.OrdinalIgnoreCase))
            return EvaluateLet(context, node.Parameters);

        if (node.Prefix is null && string.Equals(node.Name, "LAMBDA", StringComparison.OrdinalIgnoreCase))
            return DefineLambda(context, node.Parameters);

        if (!_functions.TryGetFunc(node.Name, out var fn))
            return CallName(context, node);

        var parameters = node.Parameters;
        var pool = _argsPool.Rent(parameters.Count);
        var args = new Span<AnyValue>(pool, 0, parameters.Count);

        // D38. An argument is evaluated in the context its function supplies, not the one the
        // formula started in, so operand intersection stops at the argument boundary. Excel is
        // finer-grained than this — it intersects inside SUM and MIN as well, and stops only at a
        // genuinely array-typed parameter such as SUMPRODUCT's — but XLibur has no data telling
        // SUM apart from SUMPRODUCT, and the suite's SUMPRODUCT/AVERAGE expectations are written
        // for array semantics. See specs/53 for the evidence and what closing the gap would cost.
        // Whatever the boundary, it must be restored rather than assumed false: a function call can
        // appear inside an operand of a top-level operator (`=A1+SUM(B1:B3)*C1:C3`), and the
        // operator after the call still has to intersect.
        var outerIntersectOperands = context.IntersectOperands;
        try
        {
            context.IntersectOperands = false;
            if (EvaluateArguments(context, fn!, parameters, args) is { } pickedFunction)
                return pickedFunction;

            // The flag stays off for the call itself, not just for the arguments. A few functions
            // apply an operator to their own arguments from inside the body — SWITCH compares with
            // AnyValue.IsEqual — and an operator reached that way is not a top-level operator of
            // the formula, whatever the enclosing context was. The finally restores it.
            return !context.IsArrayCalculation
                ? fn!.CallFunction(context, args)
                : fn!.CallAsArray(context, args);
        }
        finally
        {
            context.IntersectOperands = outerIntersectOperands;
            _argsPool.Return(pool);
        }
    }

    public AnyValue Visit(CalcContext context, ReferenceNode node)
    {
        return node.GetReference(context);
    }

    public AnyValue Visit(CalcContext context, NameNode node)
    {
        // A name a LET around this node has bound comes before a defined name of the same spelling.
        if (node.Prefix is null && context.LetScope is { } scope && scope.TryGetValue(node.Name, out var bound))
            return bound;

        return node.GetValue(context);
    }

    public AnyValue Visit(CalcContext context, NotSupportedNode node)
        => throw new UnsupportedFeatureException($"Evaluation of {node.FeatureName} is not implemented.");

    public AnyValue Visit(CalcContext context, CallNode node)
        => Call(context, node.Target.Accept(context, this), node.Arguments);

    public AnyValue Visit(CalcContext context, StructuredReferenceNode node)
    {
        if (!StructuredReferenceResolver.TryResolve(context, node, out var worksheet, out var range, out var error))
            return error;

        // The table's own sheet, not the formula's — a table name is workbook scoped and can be
        // referenced from anywhere in the workbook.
        return new Reference(XLRangeAddress.FromSheetRange(worksheet, range));
    }

    public AnyValue Visit(CalcContext context, PrefixNode node)
        => throw new InvalidOperationException("Node should never be visited.");

    public AnyValue Visit(CalcContext context, FileNode node)
        => throw new InvalidOperationException("Node should never be visited.");

    /// <summary>
    /// <c>LET(name1, value1, name2, value2, …, body)</c>: evaluates each value in turn, in a scope
    /// that holds the names bound before it, then the body in a scope that holds them all.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A value is bound as it evaluates: a range stays a reference, so <c>LET(r,A1:A3,INDEX(r,2))</c>
    /// is 2, and an error is the value of its name, which reaches the result only where the name is
    /// used: <c>LET(x,1/0,5)</c> is 5, as in Excel. A computed array is bound through
    /// <see cref="MemoizedArray"/>, so a name read many times computes its elements once.
    /// </para>
    /// <para>
    /// Excel refuses a <c>LET</c> without a body, or with something other than a name where a name
    /// goes, as it is typed. A file from another producer can still hold one, and it is
    /// <c>#VALUE!</c>.
    /// </para>
    /// </remarks>
    private AnyValue EvaluateLet(CalcContext context, IReadOnlyList<ValueNode> parameters)
    {
        if (parameters.Count < 3 || parameters.Count % 2 == 0)
            return XLError.IncompatibleValue;

        for (var i = 0; i < parameters.Count - 1; i += 2)
        {
            if (parameters[i] is not NameNode { Prefix: null })
                return XLError.IncompatibleValue;
        }

        var outer = context.LetScope;
        try
        {
            for (var i = 0; i < parameters.Count - 1; i += 2)
            {
                // A computed array keeps each element once it is read, as a name can be read many
                // times; a reference stays a reference.
                var value = parameters[i + 1].Accept(context, this);
                if (value.TryPickArray(out var array))
                    value = MemoizedArray.Of(array!);

                context.LetScope = new LetScope<AnyValue>(((NameNode)parameters[i]).Name, value, context.LetScope);
            }

            return parameters[^1].Accept(context, this);
        }
        finally
        {
            context.LetScope = outer;
        }
    }

    /// <summary>
    /// A call of a name that is no function: a <c>LET</c> name or a parameter bound to a <c>LAMBDA</c>,
    /// or a defined name that refers to one. One of Excel's functions that XLibur lacks is <c>#NAME?</c>.
    /// </summary>
    private AnyValue CallName(CalcContext context, FunctionNode node)
    {
        return FormulaParser.IsCallOfName(_functions, node.Name)
            ? Call(context, ResolveCalledName(context, node), node.Parameters)
            : XLError.NameNotRecognized;
    }

    /// <summary>
    /// Evaluates the arguments of a call of <paramref name="function"/> into <paramref name="args"/>.
    /// </summary>
    /// <remarks>
    /// <c>IF</c> evaluates only the branch its condition picks, as Excel does, so a <c>LAMBDA</c> that
    /// calls itself inside <c>IF</c> stops at its base case. A branch left out is blank, which <c>IF</c>
    /// never reads.
    /// </remarks>
    /// <returns>
    /// The function the branch <c>IF</c> picks holds, which is the result of the <c>IF</c>, as the array
    /// path, reading each argument as values, would make it <c>#VALUE!</c>; otherwise null.
    /// </returns>
    private AnyValue? EvaluateArguments(CalcContext context, FunctionDefinition function,
        IReadOnlyList<ValueNode> parameters, Span<AnyValue> args)
    {
        var isIf = ReferenceEquals(function, _if);
        var unused = 0;
        for (var i = 0; i < parameters.Count; ++i)
        {
            if ((unused & (1 << i)) != 0)
            {
                args[i] = AnyValue.Blank;
                continue;
            }

            args[i] = parameters[i].Accept(context, this);
            if (!isIf)
                continue;

            if (i == 0)
                unused = UnusedIfBranches(context, args[0]);
            else if (unused != 0 && args[i].IsLambda)
                return args[i];
        }

        return null;
    }

    /// <summary>
    /// The branches of an <c>IF</c> its condition leaves unused, as bits by argument index: the false
    /// branch when the condition is true, the true branch when it is false, and both when it is an
    /// error, which is the result of the <c>IF</c>.
    /// </summary>
    /// <remarks>
    /// The condition is read as <c>IF</c> reads it when it is called: in a legacy formula intersected
    /// and then reduced to one value, the first element of an array or the one cell of a reference. In
    /// an array formula a condition with many values picks a branch for each of them, and both are
    /// needed.
    /// </remarks>
    private static int UnusedIfBranches(CalcContext context, in AnyValue condition)
    {
        const int both = (1 << 1) | (1 << 2);
        ScalarValue scalar;
        if (context.IsArrayCalculation)
        {
            if (!condition.TryPickSingleOrMultiValue(out scalar, out _, context))
                return 0;
        }
        else if (!condition.ImplicitIntersection(context).TryPickScalar(out scalar, out var collection))
        {
            if (collection.TryPickT0(out var array, out var reference))
                scalar = array[0, 0];
            else if (!reference.TryGetSingleCellValue(out scalar, context))
                return both;
        }

        if (!scalar.TryCoerceLogicalOrBlankOrNumberOrText(out var isTrue, out _))
            return both;

        return isTrue ? 1 << 2 : 1 << 1;
    }

    /// <summary>
    /// <c>LAMBDA(parameter1, …, body)</c>: a function value, which keeps the names bound where it is
    /// written. The body is not evaluated until the function is called.
    /// </summary>
    /// <remarks>
    /// Excel refuses a <c>LAMBDA</c> without a body, with something other than a name where a parameter
    /// goes, or with a parameter named twice, as it is typed. A file from another producer can still
    /// hold one, and it is <c>#VALUE!</c>.
    /// </remarks>
    private static AnyValue DefineLambda(CalcContext context, IReadOnlyList<ValueNode> parameters)
    {
        if (parameters.Count == 0)
            return XLError.IncompatibleValue;

        var names = new string[parameters.Count - 1];
        for (var i = 0; i < names.Length; ++i)
        {
            if (parameters[i] is not NameNode { Prefix: null } parameter)
                return XLError.IncompatibleValue;

            for (var j = 0; j < i; ++j)
            {
                if (LetScope<AnyValue>.IsSameName(names[j], parameter.Name))
                    return XLError.IncompatibleValue;
            }

            names[i] = parameter.Name;
        }

        return AnyValue.From(new LambdaFunction(names, parameters[^1], context.LetScope));
    }

    /// <summary>
    /// The value of a name written as a function, <c>f(3)</c>, that is no function: what a <c>LET</c>
    /// or a call around it has bound it to, or else the value of the defined name, or <c>#NAME?</c>.
    /// </summary>
    private static AnyValue ResolveCalledName(CalcContext context, FunctionNode node)
    {
        if (node.Prefix is null && context.LetScope is { } scope && scope.TryGetValue(node.Name, out var bound))
            return bound;

        // XLWorkbook.EvaluateExpr evaluates without a workbook, which has no names to find.
        if (!context.HasWorkbook)
            return XLError.NameNotRecognized;

        return new NameNode(node.Prefix, node.Name).GetValue(context);
    }

    /// <summary>
    /// Calls the value of <paramref name="target"/> with <paramref name="arguments"/>. An error is the
    /// result, and any value other than a function is <c>#VALUE!</c>.
    /// </summary>
    private AnyValue Call(CalcContext context, AnyValue target, IReadOnlyList<ValueNode> arguments)
    {
        if (target.TryPickLambda(out var lambda))
            return Invoke(context, lambda, arguments);

        return target.TryPickError(out var error) ? error : XLError.IncompatibleValue;
    }

    /// <summary>
    /// Evaluates the body of <paramref name="lambda"/> with each parameter bound to its argument, in
    /// the scope the lambda was written in, not the one it is called from.
    /// </summary>
    /// <remarks>
    /// Each argument is evaluated where the call is, before the body, and bound as a <c>LET</c> binds a
    /// value: a range stays a reference, a computed array is memoized, and an error is the value of its
    /// parameter, so <c>LAMBDA(x,5)(1/0)</c> is 5, as in Excel. A call with more or fewer arguments
    /// than parameters is <c>#VALUE!</c>. An argument left empty, <c>(1,)</c>, is blank.
    /// </remarks>
    /// <exception cref="LambdaDepthException">The call is too deep.</exception>
    private AnyValue Invoke(CalcContext context, LambdaFunction lambda, IReadOnlyList<ValueNode> arguments)
    {
        if (arguments.Count != lambda.Parameters.Count)
            return XLError.IncompatibleValue;

        if (context.LambdaDepth >= MaxLambdaDepth)
            throw new LambdaDepthException();

        // A call takes a few kilobytes of stack, so a thread with the usual 1 MB runs out at a depth of
        // about 400. The call then goes on on a thread with a stack of its own, while this one waits,
        // so every thread reaches the same depth.
        if (!RuntimeHelpers.TryEnsureSufficientExecutionStack())
            return OnFreshStack((Visitor: this, Context: context, Lambda: lambda, Arguments: arguments),
                static s => s.Visitor.Invoke(s.Context, s.Lambda, s.Arguments));

        var scope = lambda.Scope;
        for (var i = 0; i < arguments.Count; ++i)
        {
            var value = arguments[i].Accept(context, this);
            if (value.TryPickArray(out var array))
                value = MemoizedArray.Of(array!);

            scope = new LetScope<AnyValue>(lambda.Parameters[i], value, scope);
        }

        var callerScope = context.LetScope;
        context.LetScope = scope;
        context.LambdaDepth++;
        try
        {
            return lambda.Body.Accept(context, this);
        }
        finally
        {
            context.LambdaDepth--;
            context.LetScope = callerScope;
        }
    }

    /// <summary>
    /// Runs <paramref name="evaluate"/> on a new thread with a stack of <see cref="FreshStackSize"/>, and
    /// waits for it. Only one thread evaluates at a time, so nothing is shared between two that run.
    /// </summary>
    /// <exception cref="LambdaDepthException">The platform has no threads, as WebAssembly in a browser.</exception>
    private static AnyValue OnFreshStack<TState>(TState state, Func<TState, AnyValue> evaluate)
    {
        if (OperatingSystem.IsBrowser())
            throw new LambdaDepthException();

        var culture = CultureInfo.CurrentCulture;
        var uiCulture = CultureInfo.CurrentUICulture;
        var result = default(AnyValue);
        ExceptionDispatchInfo? failure = null;
        var thread = new Thread(() =>
        {
            CultureInfo.CurrentCulture = culture;
            CultureInfo.CurrentUICulture = uiCulture;
            try
            {
                result = evaluate(state);
            }
            catch (Exception e)
            {
                failure = ExceptionDispatchInfo.Capture(e);
            }
        }, FreshStackSize);
        thread.Start();
        thread.Join();
        failure?.Throw();
        return result;
    }

    /// <summary>
    /// Evaluates the <c>@</c> implicit intersection operator (e.g. <c>@A1:A10</c>): a range gives the
    /// cell on the formula's row or column, an array its top-left element, and anything else itself.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A range that the formula's row and column both miss, or a reference of more than one area, is
    /// <c>#VALUE!</c>. The result of a range is a reference rather than its value, so the operand of
    /// <c>:</c> in <c>D3:@A1:A2</c> and the argument of <c>ROW(@A1:A10)</c> are still references.
    /// </para>
    /// <para>
    /// This is the operator Excel 2019 and later show in a formula. It differs from the implicit
    /// intersection that legacy Excel applies silently, <see cref="AnyValue.ImplicitIntersection"/>,
    /// only for an array: the silent one applies where a scalar is expected and leaves an array
    /// alone, while <c>@</c> applies wherever it is written, so it reduces an array too.
    /// </para>
    /// </remarks>
    private static AnyValue EvaluateImplicitIntersection(CalcContext context, AnyValue operand)
    {
        if (operand.TryPickArray(out var array))
            return array![0, 0].ToAnyValue();

        return operand.ImplicitIntersection(context);
    }

    /// <summary>
    /// Evaluates the <c>#</c> spill-range operator (e.g. <c>A1#</c>): resolves the operand to a
    /// spill anchor and returns a <see cref="Reference"/> to that dynamic array's current
    /// footprint. Returns <c>#REF!</c> when the operand cell is not a spill anchor.
    /// </summary>
    private static AnyValue EvaluateSpillRange(CalcContext context, AnyValue operand)
    {
        // The operand of `#` must resolve to a single-cell reference: the spill anchor. A
        // multi-cell area (e.g. A1:B3#) is not a valid anchor, so it is a #REF!.
        if (!operand.TryPickArea(out var anchorArea, out var error))
            return error;

        if (anchorArea.FirstAddress.RowNumber != anchorArea.LastAddress.RowNumber ||
            anchorArea.FirstAddress.ColumnNumber != anchorArea.LastAddress.ColumnNumber)
            return XLError.CellReference;

        var sheet = anchorArea.Sheet as XLWorksheet ?? context.Worksheet;
        var anchorRow = anchorArea.FirstAddress.RowNumber;
        var anchorColumn = anchorArea.FirstAddress.ColumnNumber;

        // Force the anchor to be current before reading its footprint: for a dirty anchor this
        // throws GettingDataException so the calc chain evaluates the anchor (spilling it and
        // updating its Range) before this formula. The returned value itself is unused.
        _ = context.GetCellValue(sheet, anchorRow, anchorColumn);

        var formula = sheet.Internals.CellsCollection.FormulaSlice.Get(new Point(anchorRow, anchorColumn));
        if (formula is null || !formula.IsDynamicArray)
            return XLError.CellReference; // #REF! — the cell is not a spill anchor.

        var footprint = formula.Range;
        if (footprint == default)
            return XLError.CellReference; // Anchor exists but hasn't produced a footprint yet.

        var rangeAddress = new XLRangeAddress(
            new XLAddress(sheet, footprint.TopRow, footprint.LeftColumn, true, true),
            new XLAddress(sheet, footprint.BottomRow, footprint.RightColumn, true, true));
        return new Reference(rangeAddress);
    }
}
