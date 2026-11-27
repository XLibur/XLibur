using System;
using XLibur.Excel.CalcEngine.Functions;
using static XLibur.Excel.CalcEngine.Functions.SignatureAdapter;

namespace XLibur.Excel.CalcEngine;

internal static class Logical
{
    public static void Register(FunctionRegistry ce)
    {
        ce.RegisterFunction("AND", 1, int.MaxValue, And, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("FALSE", 0, 0, Adapt(False), FunctionFlags.Scalar);
        ce.RegisterFunction("IF", 2, 3, AdaptLastOptional(If, false), FunctionFlags.Range, AllowRange.Only, 1, 2);
        ce.RegisterFunction("IFERROR", 2, 2, IfError, FunctionFlags.TakesFunctionValue);
        ce.RegisterFunction("IFNA", 2, 2, IfNa, FunctionFlags.TakesFunctionValue); // Returns the value you specify if the expression resolves to #N/A
        ce.RegisterFunction("IFS", 2, 254, Ifs, FunctionFlags.Scalar); // Returns the value for the first TRUE condition
        ce.RegisterFunction("NOT", 1, 1, AdaptCoerced(Not), FunctionFlags.Scalar);
        ce.RegisterFunction("OR", 1, int.MaxValue, Or, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("SWITCH", 3, 254, Switch, FunctionFlags.Scalar); // Matches an expression against a list of values
        ce.RegisterFunction("TRUE", 0, 0, Adapt(True), FunctionFlags.Scalar);
    }

    private static AnyValue And(CalcContext ctx, Span<AnyValue> args)
    {
        var aggResult = args.Aggregate<bool?>(
            ctx,
            null,
            XLError.IncompatibleValue,
            static (acc, val) => val is null ? acc : (acc ?? true) && val.Value,
            ToLogicalOrSkip,
            IsLogicalNumberOrError);

        return ToLogicalResult(aggResult);
    }

    // Collection filter for AND and OR. Elements of an array or reference are not converted from
    // text (text is skipped) and blanks in references are ignored. Errors must pass the filter so
    // the convert step returns them: in Excel an error element wins over any logical result.
    private static bool IsLogicalNumberOrError(ScalarValue v) => v.IsLogical || v.IsNumber || v.IsError;

    /// <summary>
    /// A value of AND and OR, or null for a text argument other than <c>TRUE</c> and <c>FALSE</c>,
    /// which is skipped as text in a range is: <c>AND("x",TRUE)</c> is TRUE, and <c>OR("5")</c>,
    /// with no value left, is <c>#VALUE!</c>.
    /// </summary>
    private static OneOf<bool?, XLError> ToLogicalOrSkip(ScalarValue v, CalcContext _)
    {
        if (v.IsError)
            return v.GetError();

        return v.TryCoerceLogicalOrBlankOrNumberOrText(out var logical, out var _) ? logical : null;
    }

    private static AnyValue ToLogicalResult(OneOf<bool?, XLError> aggResult)
    {
        if (!aggResult.TryPickT0(out var value, out var error))
            return error;

        return value is { } logical ? logical : XLError.IncompatibleValue;
    }

    // FALSE() and TRUE() are zero-argument Excel functions. They are registered through Adapt as
    // method groups, so they have to stay methods; a constant cannot be converted to the delegate.
#pragma warning disable S3400
    private static ScalarValue False()
    {
        return false;
    }
#pragma warning restore S3400

    private static AnyValue If(ScalarValue condition, AnyValue valueIfTrue, AnyValue valueIfFalse)
    {
        if (!condition.TryCoerceLogicalOrBlankOrNumberOrText(out var value, out var error))
            return error;

        return value ? valueIfTrue : valueIfFalse;
    }

    private static AnyValue IfError(CalcContext ctx, Span<AnyValue> args)
        => IfErrorMatching(ctx, args, static _ => true);

    private static AnyValue IfNa(CalcContext ctx, Span<AnyValue> args)
        => IfErrorMatching(ctx, args, static error => error == XLError.NoValueAvailable);

    /// <summary>
    /// <c>IFERROR(value, value_if_error)</c>, and <c>IFNA</c>, for the errors <paramref name="isCaught"/>
    /// is true for. Each argument is read as one value, except a function value, a <c>LAMBDA</c>, which is
    /// no error and is the result as it is, so <c>LET(f,IFERROR(LAMBDA(x,x*2),1),f(4))</c> is 8 in Excel
    /// (#868).
    /// </summary>
    private static AnyValue IfErrorMatching(CalcContext ctx, Span<AnyValue> args, Func<XLError, bool> isCaught)
    {
        var value = args[0];
        if (value.IsLambda)
            return value;

        var scalar = value.ReduceToScalar(ctx);
        if (!scalar.TryPickError(out var error) || !isCaught(error))
            return scalar.ToAnyValue();

        var alternative = args[1];
        return alternative.IsLambda ? alternative : alternative.ReduceToScalar(ctx).ToAnyValue();
    }

    private static AnyValue Ifs(CalcContext ctx, Span<AnyValue> args)
    {
        // IFS(condition1, value1, [condition2, value2], ...). Return the value paired with the first
        // condition that is TRUE. Iterate whole pairs only, so a dangling trailing argument (odd
        // count) simply can't match — matching Excel, which then returns #N/A.
        for (var i = 0; i + 1 < args.Length; i += 2)
        {
            if (!args[i].TryPickScalar(out var conditionScalar, out _))
                return XLError.IncompatibleValue;

            if (!conditionScalar.TryCoerceLogicalOrBlankOrNumberOrText(out var condition, out var error))
                return error;

            if (condition)
                return args[i + 1];
        }

        // No condition evaluated to TRUE.
        return XLError.NoValueAvailable;
    }

    private static AnyValue Switch(CalcContext ctx, Span<AnyValue> args)
    {
        // SWITCH(expression, value1, result1, [value2, result2], ..., [default]).
        ref var expression = ref args[0];
        if (expression.TryPickError(out var expressionError))
            return expressionError;

        int i;
        for (i = 1; i + 1 < args.Length; i += 2)
        {
            // Compare with Excel '=' semantics (case-insensitive text, numeric coercion, ...).
            var comparison = AnyValue.IsEqual(expression, args[i], ctx);
            if (comparison.TryPickError(out var comparisonError))
                return comparisonError;

            if (comparison.TryPickScalar(out var scalar, out _) &&
                scalar.TryPickLogical(out var isMatch) && isMatch)
                return args[i + 1];
        }

        // An odd trailing argument (no result to pair with) is the default value.
        return i < args.Length ? args[i] : XLError.NoValueAvailable;
    }

    private static AnyValue Not(bool value)
    {
        return !value;
    }

    private static AnyValue Or(CalcContext ctx, Span<AnyValue> args)
    {
        var aggResult = args.Aggregate<bool?>(
            ctx,
            null,
            XLError.IncompatibleValue,
            static (acc, val) => val is null ? acc : (acc ?? false) || val.Value,
            ToLogicalOrSkip,
            IsLogicalNumberOrError);

        return ToLogicalResult(aggResult);
    }

    // See the note on False(): registered as a method group, so it cannot become a constant.
#pragma warning disable S3400
    private static ScalarValue True()
    {
        return true;
    }
#pragma warning restore S3400
}
