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
        ce.RegisterFunction("IFERROR", 2, 2, Adapt(IfError), FunctionFlags.Scalar);
        ce.RegisterFunction("IFS", 2, 254, Ifs, FunctionFlags.Scalar); // Returns the value for the first TRUE condition
        ce.RegisterFunction("NOT", 1, 1, AdaptCoerced(Not), FunctionFlags.Scalar);
        ce.RegisterFunction("OR", 1, int.MaxValue, Or, FunctionFlags.Range, AllowRange.All);
        ce.RegisterFunction("SWITCH", 3, 254, Switch, FunctionFlags.Scalar); // Matches an expression against a list of values
        ce.RegisterFunction("TRUE", 0, 0, Adapt(True), FunctionFlags.Scalar);
    }

    private static AnyValue And(CalcContext ctx, Span<AnyValue> args)
    {
        var aggResult = args.Aggregate(
            ctx,
            true,
            XLError.IncompatibleValue,
            static (acc, val) => acc && val,
            static (v, _) =>
            {
                // Skip values that can't be converted, but aren't errors, like "text"
                if (v.IsError)
                    return v.GetError();
                if (!v.TryCoerceLogicalOrBlankOrNumberOrText(out var logical, out var _))
                    return true;
                return logical;
            },
            static v => v.IsLogical || v.IsNumber); // No text conversion for element of collection, blanks are ignored in references

        if (!aggResult.TryPickT0(out var value, out var error))
            return error;

        return value;
    }

    private static ScalarValue False()
    {
        return false;
    }

    private static AnyValue If(ScalarValue condition, AnyValue valueIfTrue, AnyValue valueIfFalse)
    {
        if (!condition.TryCoerceLogicalOrBlankOrNumberOrText(out var value, out var error))
            return error;

        return value ? valueIfTrue : valueIfFalse;
    }

    private static AnyValue IfError(ScalarValue potentialError, ScalarValue alternative)
    {
        if (!potentialError.IsError)
            return potentialError.ToAnyValue();

        return alternative.ToAnyValue();
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
        var aggResult = args.Aggregate(
            ctx,
            false,
            XLError.IncompatibleValue,
            static (acc, val) => acc || val,
            static (v, _) =>
            {
                // Skip values that can't be converted, but aren't errors, like "text"
                if (v.IsError)
                    return v.GetError();
                if (!v.TryCoerceLogicalOrBlankOrNumberOrText(out var logical, out var _))
                    return false;
                return logical;
            },
            static v => v.IsLogical || v.IsNumber); // No text conversion for element of collection, blanks are ignored in references

        if (!aggResult.TryPickT0(out var value, out var error))
            return error;

        return value;
    }

    private static ScalarValue True()
    {
        return true;
    }
}
