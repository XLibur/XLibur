using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using XLibur.Excel.CalcEngine.Exceptions;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// The workbook function library, callable outside a worksheet — the same functions a cell formula
/// can call, evaluated without a grid to be relative to.
/// </summary>
/// <remarks>
/// <para>
/// This is for calling one function with arguments you already have, which is what a reporting or
/// templating layer needs. It is not a formula evaluator: there is no parsing, no cell references
/// and no ranges. To evaluate a formula string against real data, use
/// <see cref="IXLWorksheet.Evaluate(string, string)"/>.
/// </para>
/// <para>
/// An instance holds no per-call state and is safe for concurrent use. Construct one per culture
/// and share it — constructing one builds the whole function table, so making one per call is
/// wasteful in a way that is easy not to notice.
/// </para>
/// </remarks>
public sealed class XLFunctionLibrary
{
    private readonly XLCalcEngine _engine;
    private readonly CultureInfo _culture;
    private readonly IReadOnlyCollection<string> _names;

    /// <summary>
    /// Creates a library whose functions parse and format according to <paramref name="culture"/>.
    /// </summary>
    /// <param name="culture">
    /// Culture for parsing and formatting. Defaults to <see cref="CultureInfo.InvariantCulture"/>.
    /// </param>
    public XLFunctionLibrary(CultureInfo? culture = null)
    {
        _culture = culture ?? CultureInfo.InvariantCulture;
        _engine = new XLCalcEngine(_culture);

        // Snapshotted rather than deferred to the registry's key collection, so that Names is a
        // stable, safely-enumerable value for a type that promises to be shareable across threads.
        _names = _engine.Functions.Names.ToList().AsReadOnly();
    }

    /// <summary>
    /// Names of every available function, in Excel's own casing (<c>SUM</c>, <c>VLOOKUP</c>).
    /// </summary>
    /// <remarks>
    /// <see cref="TryInvoke"/> itself matches names case-insensitively, as Excel does; the casing
    /// here is for display and for callers registering these functions under a name of their own.
    /// </remarks>
    public IReadOnlyCollection<string> Names => _names;

    /// <summary>
    /// Calls <paramref name="name"/> with <paramref name="arguments"/>.
    /// </summary>
    /// <param name="name">Function name, matched case-insensitively.</param>
    /// <param name="arguments">
    /// The arguments, already flattened to scalars. A function that takes a range in a cell formula
    /// — <c>SUM</c>, <c>AVERAGE</c>, <c>COUNT</c> — takes those cells as a run of arguments here.
    /// </param>
    /// <param name="result">The function's result, or <c>default</c> if there is no such function.</param>
    /// <returns>
    /// <c>false</c> if no function has that name, leaving <paramref name="result"/> default.
    /// Otherwise <c>true</c>; a call that was made but could not succeed — wrong arity, wrong
    /// argument type, a division by zero — returns <c>true</c> with an <see cref="XLError"/>
    /// result, which is how Excel itself reports these.
    /// </returns>
    /// <exception cref="ArgumentNullException"><paramref name="name"/> is <c>null</c>.</exception>
    /// <exception cref="XLNoWorksheetContextException">
    /// The function needs a worksheet to be relative to (<c>ROW</c>, <c>OFFSET</c>, <c>INDIRECT</c>
    /// and the like). Those belong in a real cell formula.
    /// </exception>
    public bool TryInvoke(string name, ReadOnlySpan<XLCellValue> arguments, out XLCellValue result)
    {
        ArgumentNullException.ThrowIfNull(name);

        if (!_engine.Functions.TryGetFunc(name, out var definition) || definition is null)
        {
            result = default;
            return false;
        }

        if (arguments.Length < definition.MinParams || arguments.Length > definition.MaxParams)
        {
            result = XLError.IncompatibleValue;
            return true;
        }

        // System.Array, spelled out: the calc engine has its own Array type in this namespace.
        var args = arguments.Length == 0 ? System.Array.Empty<AnyValue>() : new AnyValue[arguments.Length];
        for (var i = 0; i < arguments.Length; ++i)
            args[i] = ((ScalarValue)arguments[i]).ToAnyValue();

        // No workbook and no cell: the whole point is a call made before there is a grid to be
        // relative to. Functions that reach for one throw, and are translated below.
        var context = new CalcContext(_engine, _culture, workbook: null, worksheet: null, formulaAddress: null);

        AnyValue value;
        try
        {
            value = definition.CallFunction(context, args.AsSpan());
        }
        catch (MissingContextException ex)
        {
            throw new XLNoWorksheetContextException(
                $"'{name}' needs a worksheet to be relative to, and was called without one. "
                + "Use it in a cell formula instead.",
                ex);
        }

        result = ToCellValue(value);
        return true;
    }

    /// <summary>
    /// The result as a cell would hold it.
    /// </summary>
    /// <remarks>
    /// A blank result stays blank rather than becoming <c>0</c> — unlike
    /// <see cref="ScalarValue.ToCellValue"/>, which applies the rule for a formula *in a cell*,
    /// where a blank result is written as zero. There is no cell here, so the caller is told what
    /// the function actually returned and decides what blank means to it.
    /// </remarks>
    private static XLCellValue ToCellValue(AnyValue value)
    {
        // An array or reference result has no single value a cell could hold. Reported the way
        // Excel reports a value it cannot place, rather than throwing.
        if (!value.TryPickScalar(out var scalar, out _))
            return XLError.IncompatibleValue;

        return scalar.Match<XLCellValue>(
            () => Blank.Value,
            logical => logical,
            number => number,
            text => text,
            error => error);
    }
}
