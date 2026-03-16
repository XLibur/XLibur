using System;
using System.Collections.Generic;

namespace XLibur.Excel.CalcEngine.Functions;

/// <summary>
/// A collection of adapter functions from a more a generic formula function to more specific ones.
/// </summary>
internal static class SignatureAdapter
{
    #region Signature adapters
    // Each method converts a more specific signature of a function into a generic formula function type.
    // We have many functions with same signature and the adapters should be reusable. Convert parameters
    // through value converters below. We can hopefully generate them at a later date, so try to keep them similar.

    public static CalcEngineFunction Adapt(Func<ScalarValue> f)
    {
        return (_, _) => f().ToAnyValue();
    }

    public static CalcEngineFunction AdaptCoerced(Func<bool, AnyValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = CoerceToLogical(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            return f(arg0);
        };
    }

    public static CalcEngineFunction Adapt(Func<double, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            return f(arg0).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, double, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            return f(ctx, arg0).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, ScalarValue, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            return f(ctx, arg0).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, double, double, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            return f(ctx, arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, double, double, double, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, double, double, string, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2Converted = ToText(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<double, double, double, bool, AnyValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

            var arg3Converted = CoerceToLogical(args[3], ctx);
            if (!arg3Converted.TryPickT0(out var arg3, out var err3))
                return err3;

            return f(arg0, arg1, arg2, arg3);
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, string, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            return f(ctx, arg0).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<string, string, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToText(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            return f(arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<string, double, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            return f(arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, string, double, double, string, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

            var arg3Converted = ToText(args[3], ctx);
            if (!arg3Converted.TryPickT0(out var arg3, out var err3))
                return err3;

            return f(ctx, arg0, arg1, arg2, arg3).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, string, double, double, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, AnyValue, double, AnyValue> f)
    {
        return (ctx, args) =>
        {
            var arg0 = args[0];

            var arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            return f(ctx, arg0, arg1);
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, string, ScalarValue?, AnyValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1 = default(ScalarValue?);
            if (args.Length > 1)
            {
                var arg1Converted = ToScalarValue(args[1], ctx);
                if (!arg1Converted.TryPickT0(out var arg1Value, out var err1))
                    return err1;

                arg1 = arg1Value;
            }


            return f(ctx, arg0, arg1);
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, List<string>, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var texts = new List<string>(args.Length);
            foreach (var arg in args)
            {
                var argConverted = ToText(arg, ctx);
                if (!argConverted.TryPickT0(out var text, out var error))
                    return error;

                texts.Add(text);
            }

            return f(ctx, texts).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, AnyValue, AnyValue> f)
    {
        return (ctx, args) => f(ctx, args[0]);
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, ScalarValue, AnyValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            return f(ctx, arg0);
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, ScalarValue, string, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToText(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            return f(ctx, arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<ScalarValue, ScalarValue, AnyValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToScalarValue(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            return f(arg0, arg1);
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, AnyValue, ScalarValue, AnyValue> f)
    {
        return (ctx, args) =>
        {
            var arg0 = args[0];

            var arg1Converted = ToScalarValue(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            return f(ctx, arg0, arg1);
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, List<Array>, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arrays = new List<Array>();
            foreach (var arg in args)
            {
                if (arg.TryPickSingleOrMultiValue(out var scalar, out var array, ctx))
                    array = new ScalarArray(scalar, 1, 1);

                arrays.Add(array!);
            }

            return f(ctx, arrays).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, string, bool, List<AnyValue>, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = CoerceToLogical(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var remainingArgs = new List<AnyValue>();
            foreach (var arg in args[2..])
                remainingArgs.Add(arg);

            return f(ctx, arg0, arg1, remainingArgs).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<ScalarValue, AnyValue, AnyValue, AnyValue> f, AnyValue lastDefault)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1 = args[1];
            var arg2 = args.Length > 2 ? args[2] : lastDefault;
            return f(arg0, arg1, arg2);
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, double, double, ScalarValue> f, double lastDefault)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args.Length > 1 ? args[1] : lastDefault, ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            return f(ctx, arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, double, double, double, ScalarValue> f, double lastDefault)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2Converted = ToNumber(args.Length > 2 ? args[2] : lastDefault, ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, double, double, bool, ScalarValue> f, bool lastDefault)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2Converted = CoerceToLogical(args.Length > 2 ? args[2] : lastDefault, ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, string, double, ScalarValue> f, double lastDefault)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args.Length > 1 ? args[1] : lastDefault, ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            return f(ctx, arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<double, double, double, ScalarValue> f, double lastDefault)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2Converted = ToNumber(args.Length > 2 ? args[2] : lastDefault, ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

#pragma warning disable S2234            
            return f(arg0, arg1, arg2).ToAnyValue();
#pragma warning restore S2234            
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, double, AnyValue[], AnyValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var argsLoop = args[1..].ToArray();
            return f(ctx, arg0, argsLoop);
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, string, string, OneOf<double, Blank>, AnyValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToText(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            OneOf<double, Blank> arg2Optional = Blank.Value;
            if (args.Length > 2)
            {
                var arg2Converted = ToNumber(args[2], ctx);
                if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                    return err2;

                arg2Optional = arg2;
            }

            return f(ctx, arg0, arg1, arg2Optional);
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, ScalarValue, ScalarValue, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = args.Length > 1 ? ToScalarValue(args[1], ctx) : ScalarValue.Blank;
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            return f(ctx, arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, ScalarValue, ScalarValue, AnyValue, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToScalarValue(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2 = args.Length > 2 ? args[2] : AnyValue.Blank;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, AnyValue, ScalarValue, AnyValue, AnyValue> f)
    {
        return (ctx, args) =>
        {
            var arg0 = args[0];

            var arg1Converted = ToScalarValue(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2 = args.Length > 2 ? args[2] : AnyValue.Blank;

            return f(ctx, arg0, arg1, arg2);
        };
    }

    /// <summary>
    /// An adapter for <c>{SUM,AVERAGE}IFS</c> functions.
    /// </summary>
    public static CalcEngineFunction AdaptIfs(Func<CalcContext, AnyValue, List<(AnyValue Range, ScalarValue Criteria)>, AnyValue> f)
    {
        return (ctx, args) =>
        {
            var tallyRange = args[0];
            if (!ToCriteria(ctx, args[1..]).TryPickT0(out var criteria, out var error))
                return error;

            return f(ctx, tallyRange, criteria);
        };
    }

    /// <summary>
    /// An adapter for <c>COUNTIFS</c> function.
    /// </summary>
    public static CalcEngineFunction AdaptIfs(Func<CalcContext, List<(AnyValue Range, ScalarValue Criteria)>, AnyValue> f)
    {
        return (ctx, args) =>
        {
            if (!ToCriteria(ctx, args).TryPickT0(out var criteria, out var error))
                return error;

            return f(ctx, criteria);
        };
    }

    public static CalcEngineFunction AdaptIndex(Func<CalcContext, AnyValue, List<int>, AnyValue> f)
    {
        return (ctx, args) =>
        {
            var arg0 = args[0];
            var numbers = new List<int>(args.Length - 1);
            for (var i = 1; i < args.Length; ++i)
            {
                if (!ToNumber(args[i], ctx).TryPickT0(out var number, out var error))
                    return error;

                numbers.Add((int)number);
            }

            return f(ctx, arg0, numbers);
        };
    }

    public static CalcEngineFunction AdaptMatch(Func<CalcContext, ScalarValue, AnyValue, int, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1 = args[1];
            var arg2Converted = args.Length > 2 ? ToNumber(args[2], ctx) : 1;
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

            return f(ctx, arg0, arg1, (int)arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptSeriesSum(Func<CalcContext, double, double, double, Array, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            if (!ToNonLogicalNumber(args[0], ctx).TryPickT0(out var arg0, out var err0))
                return err0;

            if (!ToNonLogicalNumber(args[1], ctx).TryPickT0(out var arg1, out var err1))
                return err1;

            if (!ToNonLogicalNumber(args[2], ctx).TryPickT0(out var arg2, out var err2))
                return err2;

            if (!ToSeriesSumCoefficients(args[3], ctx).TryPickT0(out var arg3, out var err3))
                return err3;

            return f(ctx, arg0, arg1, arg2, arg3).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptNumberValue(Func<CalcContext, string, string, string, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var decimalSeparator = ctx.Culture.NumberFormat.NumberDecimalSeparator;
            var arg1Converted = ToText(args.Length > 1 ? args[1] : decimalSeparator, ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var groupSeparator = ctx.Culture.NumberFormat.NumberGroupSeparator;
            var arg2Converted = ToText(args.Length > 2 ? args[2] : groupSeparator, ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptSubstitute(Func<CalcContext, string, string, string, double?, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToText(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2Converted = ToText(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

            double? arg3 = null;
            if (args.Length > 3)
            {
                // Excel doesn't accept logical, be more permissive.
                var arg3Converted = ToNumber(args[3], ctx);
                if (!arg3Converted.TryPickT0(out var arg3Number, out var err3))
                    return err3;

                arg3 = arg3Number;
            }

            return f(ctx, arg0, arg1, arg2, arg3).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptMultinomial(Func<CalcContext, List<IEnumerable<ScalarValue>>, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            // This can skip blank values, because blank doesn't increase nominator
            // and doesn't change denominator due to 0! = 1
            var scalarCollections = new List<IEnumerable<ScalarValue>>(args.Length);
            foreach (var arg in args)
                scalarCollections.Add(GetNonBlankScalars(arg, ctx));

            return f(ctx, scalarCollections).ToAnyValue();
        };
    }

    /// <summary>
    /// Adapt a function that accepts areas as arguments (e.g. SUMPRODUCT). The key benefit is
    /// that all <c>ReferenceArray</c> allocation is done once for a function. The method
    /// shouldn't be used for functions that accept 3D references (e.g. SUMSQ). It is still
    /// necessary to check all errors in the <paramref name="f"/>, adapt method doesn't do that
    /// on its own (potential performance problem). The signature uses an array instead of
    /// IReadOnlyList interface for performance reasons (can't JIT access props through interface).
    /// </summary>
    public static CalcEngineFunction Adapt(Func<CalcContext, Array[], AnyValue> f)
    {
        return (ctx, args) =>
        {
            var areas = new Array[args.Length];
            for (var i = 0; i < args.Length; ++i)
            {
                areas[i] = args[i].TryPickSingleOrMultiValue(out var scalar, out var array, ctx)
                    ? new ScalarArray(scalar, 1, 1)
                    : array!;
            }

            return f(ctx, areas);
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, ScalarValue, AnyValue, double, bool, AnyValue> f, bool defaultValue0)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1 = args[1];

            var arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

            var arg3Converted = args.Length >= 4 ? CoerceToLogical(args[3], ctx) : defaultValue0;
            if (!arg3Converted.TryPickT0(out var arg3, out var err3))
                return err3;

            return f(ctx, arg0, arg1, arg2, arg3);
        };
    }

    public static CalcEngineFunction AdaptLastTwoOptional(Func<double, double, double, ScalarValue> f, double defaultValue1, double defaultValue2)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = args.Length > 1 ? ToNumber(args[1], ctx) : defaultValue1;
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2Converted = args.Length > 2 ? ToNumber(args[2], ctx) : defaultValue2;
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;
#pragma warning disable S2234
            return f(arg0, arg1, arg2).ToAnyValue();
#pragma warning restore S2234
        };
    }

    public static CalcEngineFunction AdaptLastTwoOptional(Func<CalcContext, double, double, bool, ScalarValue> f, double defaultValue1, bool defaultValue2)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = args.Length > 1 ? ToNumber(args[1], ctx) : defaultValue1;
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            // AnyValue to bool has different semantic than AnyValue to number, e.g. "0" is not valid for bool coercion
            var arg2Converted = args.Length > 2 ? args[2] : defaultValue2;
            if (!CoerceToLogical(arg2Converted, ctx).TryPickT0(out var arg2, out var err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastTwoOptional(Func<double, double, double, double, double, AnyValue> f, double defaultValue0, double defaultValue1)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

            if (!ToOptionalNumber(args, 3, defaultValue0, ctx).TryPickT0(out var arg3, out var err3))
                return err3;

            if (!ToOptionalNumber(args, 4, defaultValue1, ctx).TryPickT0(out var arg4, out var err4))
                return err4;
#pragma warning disable S2234
            return f(arg0, arg1, arg2, arg3, arg4);
#pragma warning restore S2234                        
        };
    }

    public static CalcEngineFunction AdaptLastTwoOptional(Func<double, double, double, double, double, double, AnyValue> f, double defaultValue0, double defaultValue1)
    {
        return (ctx, args) =>
        {
            var arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out var err0))
                return err0;

            var arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out var err1))
                return err1;

            var arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out var err2))
                return err2;

            var arg3Converted = ToNumber(args[3], ctx);
            if (!arg3Converted.TryPickT0(out var arg3, out var err3))
                return err3;

            if (!ToOptionalNumber(args, 4, defaultValue0, ctx).TryPickT0(out var arg4, out var err4))
                return err4;

            if (!ToOptionalNumber(args, 5, defaultValue1, ctx).TryPickT0(out var arg5, out var err5))
                return err5;
#pragma warning disable S2234
            return f(arg0, arg1, arg2, arg3, arg4, arg5);
#pragma warning restore S2234                                    
        };
    }

    #endregion

    #region Value converters
    // Each method is named ToSomething and it converts an argument into a desired type (e.g. for ToSomething it should be type Something).
    // Return value is always OneOf<Something, Error>, if there is an error, return it as an error.

    private static OneOf<bool, XLError> CoerceToLogical(in AnyValue value, CalcContext ctx)
    {
        if (!ToScalarValue(in value, ctx).TryPickT0(out var scalar, out var scalarError))
            return scalarError;

        // LibreOffice does accept text, tries to parse it as a number and coerces the number
        // to bool. Excel does not accept number in text argument.
        if (!scalar.TryCoerceLogicalOrBlankOrNumberOrText(out var logical, out var coercionError))
            return coercionError;

        return logical;
    }

    private static OneOf<double, XLError> ToNumber(in AnyValue value, CalcContext ctx)
    {
        if (value.TryPickScalar(out var scalar, out var collection))
            return scalar.ToNumber(ctx.Culture);

        // When user specifies array as an argument in an array formula for a scalar function, use [0,0]
        if (collection.TryPickT0(out var array, out var reference))
            return array[0, 0].ToNumber(ctx.Culture);

        if (reference.TryGetSingleCellValue(out var scalarValue, ctx))
            return scalarValue.ToNumber(ctx.Culture);

        throw new NotImplementedException("Array formulas not implemented.");
    }

    private static OneOf<string, XLError> ToText(in AnyValue value, CalcContext ctx)
    {
        if (value.TryPickScalar(out var scalar, out var collection))
            return scalar.ToText(ctx.Culture);

        if (collection.TryPickT0(out _, out var reference))
            throw new NotImplementedException("Array formulas not implemented.");

        if (reference.TryGetSingleCellValue(out var scalarValue, ctx))
            return scalarValue.ToText(ctx.Culture);

        throw new NotImplementedException("Array formulas not implemented.");
    }

    private static OneOf<ScalarValue, XLError> ToScalarValue(in AnyValue value, CalcContext ctx)
    {
        if (value.TryPickScalar(out var scalar, out var collection))
            return scalar;

        if (collection.TryPickT0(out var array, out var reference))
            return array[0, 0];

        if (reference.TryGetSingleCellValue(out var referenceScalar, ctx))
            return referenceScalar;

        return OneOf<ScalarValue, XLError>.FromT1(XLError.IncompatibleValue);
    }

    private static OneOf<double, XLError> ToNonLogicalNumber(in AnyValue value, CalcContext ctx)
    {
        if (value.IsLogical)
            return XLError.IncompatibleValue;

        return ToNumber(value, ctx);
    }

    private static OneOf<Array, XLError> ToSeriesSumCoefficients(in AnyValue value, CalcContext ctx)
    {
        if (value.TryPickSingleOrMultiValue(out var scalar, out var array, ctx))
        {
            if (scalar.IsLogical)
                return XLError.IncompatibleValue;

            if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var error))
                return error;

            return new ScalarArray(number, 1, 1);
        }

        return array!;
    }

    private static OneOf<double, XLError> ToOptionalNumber(Span<AnyValue> args, int index, double defaultValue, CalcContext ctx)
    {
        if (args.Length > index)
            return ToNumber(args[index], ctx);

        return defaultValue;
    }

    private static IEnumerable<ScalarValue> GetNonBlankScalars(AnyValue value, CalcContext ctx)
    {
        if (value.TryPickScalar(out var scalar, out var collection))
        {
            if (!scalar.IsBlank)
                yield return scalar;

            yield break;
        }

        IEnumerable<ScalarValue> source = collection.TryPickT0(out var array, out var reference)
            ? array
            : ctx.GetNonBlankValues(reference);

        foreach (var element in source)
        {
            if (!element.IsBlank)
                yield return element;
        }
    }

    private static OneOf<List<(AnyValue Range, ScalarValue Criteria)>, XLError> ToCriteria(CalcContext ctx, ReadOnlySpan<AnyValue> args)
    {
        var allCriteria = new List<(AnyValue Range, ScalarValue Criteria)>();
        var pairCount = (args.Length + 1) / 2;
        for (var i = 0; i < pairCount; ++i)
        {
            var rangeArgIndex = 2 * i;
            var range = args[rangeArgIndex];

            // Excel grammar requires even number of arguments. We can't
            // do that, so use blank for missing pair value.
            var criteriaArgIndex = rangeArgIndex + 1;
            var criteriaArgConverted = criteriaArgIndex < args.Length
                ? ToScalarValue(args[criteriaArgIndex], ctx)
                : ScalarValue.Blank;
            if (!criteriaArgConverted.TryPickT0(out var criteria, out var criteriaError))
                return criteriaError;

            allCriteria.Add((range, criteria));
        }

        return allCriteria;
    }
    #endregion
}
