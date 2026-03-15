using System;
using System.Collections.Generic;

namespace XLibur.Excel.CalcEngine.Functions;

internal sealed class TallyNumbers : ITally
{
    private readonly bool _ignoreScalarBlank;
    private readonly bool _ignoreErrors;
    private readonly Func<CalcContext, Reference, IEnumerable<ScalarValue>> _getNonBlankValues;

    /// <summary>
    /// Tally numbers.
    /// </summary>
    internal static readonly TallyNumbers Default = new();

    /// <summary>
    /// Ignore blank from scalar values. Basically used for <c>PRODUCT</c> function, so it doesn't end up with 0.
    /// </summary>
    internal static readonly TallyNumbers WithoutScalarBlank = new(ignoreScalarBlank: true);

    /// <summary>
    /// Tally algorithm for <c>SUBTOTAL</c> functions 1..11.
    /// </summary>
    internal static readonly TallyNumbers Subtotal10 = new(static (ctx, reference) => ctx.GetFilteredNonBlankValues(reference, "SUBTOTAL"));

    /// <summary>
    /// Tally algorithm for <c>SUBTOTAL</c> functions 101..111.
    /// </summary>
    internal static readonly TallyNumbers Subtotal100 = new(static (ctx, reference) => ctx.GetFilteredNonBlankValues(reference, "SUBTOTAL", skipHiddenRows: true));

    /// <summary>
    /// Tally numbers. Any error (including conversion), logical, text is ignored and not tallied.
    /// </summary>
    internal static readonly TallyNumbers IgnoreErrors = new(ignoreErrors: true);

    private TallyNumbers(Func<CalcContext, Reference, IEnumerable<ScalarValue>>? getNonBlankValues = null, bool ignoreScalarBlank = false, bool ignoreErrors = false)
    {
        _ignoreScalarBlank = ignoreScalarBlank;
        _ignoreErrors = ignoreErrors;
        _getNonBlankValues = getNonBlankValues ?? (static (ctx, reference) => ctx.GetNonBlankValues(reference));
    }

    /// <summary>
    /// The method tries to convert scalar arguments to numbers, but ignores non-numbers in
    /// reference/array. Any error found is propagated to the result.
    /// </summary>
    public OneOf<T, XLError> Tally<T>(CalcContext ctx, Span<AnyValue> args, T initialState)
        where T : ITallyState<T>
    {
        var tally = initialState;
        foreach (var arg in args)
        {
            if (arg.TryPickScalar(out var scalar, out var collection))
            {
                if (!TallyScalar(ctx, scalar, ref tally, out var error))
                    return error;
            }
            else
            {
                if (!TallyCollection(ctx, collection, ref tally, out var error))
                    return error;
            }
        }

        return tally;
    }

    private bool TallyScalar<T>(CalcContext ctx, ScalarValue scalar, ref T tally, out XLError error)
        where T : ITallyState<T>
    {
        error = default;

        if (_ignoreScalarBlank && scalar.IsBlank)
            return true;

        // Scalars are converted to number.
        if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out error))
            return _ignoreErrors;

        tally = tally.Tally(number);
        return true;
    }

    private bool TallyCollection<T>(CalcContext ctx, OneOf<Array, Reference> collection, ref T tally, out XLError error)
        where T : ITallyState<T>
    {
        error = default;
        var valuesIterator = !collection.TryPickT0(out var array, out var reference)
            ? _getNonBlankValues(ctx, reference)
            : array;

        foreach (var value in valuesIterator)
        {
            if (value.TryPickError(out error))
            {
                if (_ignoreErrors)
                    continue;

                return false;
            }

            // For arrays and references, only the number type is used. Other types are ignored.
            if (value.TryPickNumber(out var number))
                tally = tally.Tally(number);
        }

        return true;
    }
}
