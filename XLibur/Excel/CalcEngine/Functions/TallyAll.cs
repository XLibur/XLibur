using System;
using System.Collections.Generic;

namespace XLibur.Excel.CalcEngine.Functions;

/// <summary>
/// A tally function for *A functions (e.g. AverageA, MinA, MaxA). The behavior is buggy in Excel,
/// because they doesn't count logical values in array, but do count them in reference ¯\_(ツ)_/¯.
/// </summary>
internal sealed class TallyAll : ITally
{
    private readonly bool _ignoreArrayText;
    private readonly bool _includeErrors;
    private readonly Func<CalcContext, Reference, IEnumerable<ScalarValue>> _getNonBlankValues;

    /// <summary>
    /// <list type="bullet">
    ///   <item>Scalar values are converted to number, conversion might lead to errors.</item>
    ///   <item>Array values includes numbers, ignore logical and text.</item>
    ///   <item>Reference values include logical, number and text is considered a zero.</item>
    /// </list>
    /// Errors are propagated.
    /// </summary>
    internal static readonly ITally Default = new TallyAll(ignoreArrayText: true);

    /// <summary>
    /// <list type="bullet">
    ///   <item>Scalar values are converted to number, conversion might lead to errors.</item>
    ///   <item>Array values includes numbers, text is considered a zero and logical values are ignored.</item>
    ///   <item>Reference values include logical, number and text is considered a zero.</item>
    /// </list>
    /// Errors are propagated.
    /// </summary>
    internal static readonly ITally WithArrayText = new TallyAll(ignoreArrayText: false);

    /// <summary>
    /// <list type="bullet">
    ///   <item>Scalar values are converted to number, conversion might lead to errors.</item>
    ///   <item>Array values includes numbers, text is considered a zero and logical values are ignored.</item>
    ///   <item>Reference values include logical, number and text is considered a zero.</item>
    /// </list>
    /// Errors are considered zero and are <strong>not</strong> propagated.
    /// </summary>
    internal static readonly ITally IncludeErrors = new TallyAll(includeErrors: true);

    /// <summary>
    /// Tally algorithm for <c>SUBTOTAL</c> functions 1..11.
    /// </summary>
    internal static readonly ITally Subtotal10 = new TallyAll(getNonBlankValues: static (ctx, reference) => ctx.GetFilteredNonBlankValues(reference, "SUBTOTAL"));

    /// <summary>
    /// Tally algorithm for <c>SUBTOTAL</c> functions 101..111.
    /// </summary>
    internal static readonly ITally Subtotal100 = new TallyAll(getNonBlankValues: static (ctx, reference) => ctx.GetFilteredNonBlankValues(reference, "SUBTOTAL", skipHiddenRows: true));

    private TallyAll(bool ignoreArrayText = true, bool includeErrors = false, Func<CalcContext, Reference, IEnumerable<ScalarValue>>? getNonBlankValues = null)
    {
        _ignoreArrayText = ignoreArrayText;
        _includeErrors = includeErrors;
        _getNonBlankValues = getNonBlankValues ?? (static (ctx, reference) => ctx.GetNonBlankValues(reference));
    }

    public OneOf<T, XLError> Tally<T>(CalcContext ctx, Span<AnyValue> args, T initialState)
        where T : ITallyState<T>
    {
        var state = initialState;
        foreach (var arg in args)
        {
            if (arg.TryPickScalar(out var scalar, out var collection))
            {
                // Scalars are converted to number.
                if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out var error))
                {
                    if (!_includeErrors)
                        return error;

                    number = 0;
                }

                // All scalars are counted
                state = state.Tally(number);
            }
            else
            {
                var result = TallyCollection(collection, ctx, state);
                if (!result.TryPickT0(out state, out var error))
                    return error;
            }
        }

        return state;
    }

    private OneOf<T, XLError> TallyCollection<T>(OneOf<Array, Reference> collection, CalcContext ctx, T state)
        where T : ITallyState<T>
    {
        bool isArray;
        IEnumerable<ScalarValue> valuesIterator;
        if (collection.TryPickT0(out var array, out var reference))
        {
            valuesIterator = array;
            isArray = true;
        }
        else
        {
            valuesIterator = _getNonBlankValues(ctx, reference);
            isArray = false;
        }

        foreach (var value in valuesIterator)
        {
            var result = TallyValue(value, isArray, ref state);
            if (!result.TryPickT0(out _, out var error))
                return error;
        }

        return state;
    }

    private OneOf<bool, XLError> TallyValue<T>(ScalarValue value, bool isArray, ref T state)
        where T : ITallyState<T>
    {
        if (!isArray && value.TryPickLogical(out var logical))
        {
            state = state.Tally(logical ? 1 : 0);
        }
        else if (value.TryPickNumber(out var number))
        {
            state = state.Tally(number);
        }
        else if (value.IsText && (!isArray || !_ignoreArrayText))
        {
            state = state.Tally(0);
        }
        else if (value.TryPickError(out var error))
        {
            if (!_includeErrors)
                return error;

            state = state.Tally(0);
        }

        return true;
    }
}
