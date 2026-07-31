using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Scriban.Runtime;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Exceptions;
using XLibur.Report.Expressions;

namespace XLibur.Report.Functions;

/// <summary>
/// Makes XLibur's Excel function library callable from template expressions, so
/// <c>{{ SUM(items.Total) }}</c> and <c>{{ ROUND(item.Price, 2) }}</c> mean in a template what
/// they mean in a cell.
/// </summary>
/// <remarks>
/// One adapter over the calc engine's function registry rather than a wrapper per function: a
/// function added to the engine becomes available to templates without any change here.
/// <para>
/// Functions are registered under their upper-case Excel names. That is not cosmetic — Scriban's
/// keywords are lower case, so <c>if</c> would be parsed as a block conditional while <c>IF</c>
/// parses as an ordinary function call.
/// </para>
/// <para>
/// Evaluation happens with no grid: at the time a template expression runs there is no cell to be
/// relative to. Functions that need one (<c>ROW</c>, <c>OFFSET</c>, <c>INDIRECT</c> and the like)
/// report that they were called outside a worksheet rather than returning something misleading.
/// Cell references belong in real Excel formulas, which generation expands and Excel evaluates.
/// </para>
/// <para>
/// Registration is built once per culture and shared by every template. It has to be: importing
/// the ~400 functions into a Scriban <see cref="ScriptObject"/> costs about 30 KB each, which made
/// constructing an <see cref="XLTemplate"/> allocate some 12 MB — more than nine tenths of what
/// generating a small report allocated in total, and twenty times the cost of opening the template
/// workbook. Sharing the imported object turns that into one <c>PushGlobal</c>.
/// </para>
/// </remarks>
internal static class ExcelFunctionBridge
{
    /// <summary>A template-callable Excel function. Variadic, matching Excel's own arity rules.</summary>
    internal delegate object? ExcelFunction(params object?[] arguments);

    /// <summary>
    /// The adapters for one culture, and — for the default engine — those same adapters already
    /// imported into a Scriban globals object.
    /// </summary>
    /// <remarks>
    /// Keyed by culture because the adapters close over the one they convert arguments with.
    /// In practice this holds a single entry: <see cref="XLTemplate"/> registers under the
    /// invariant culture, matching <see cref="ScribanExpressionEngine"/>'s own default.
    /// </remarks>
    private static readonly ConcurrentDictionary<CultureInfo, IReadOnlyList<(string Name, Delegate Function)>> Adapters = new();

    private static readonly ConcurrentDictionary<CultureInfo, ScriptObject> ScribanGlobals = new();

    /// <summary>
    /// Makes every function the calc engine knows about callable from <paramref name="engine"/>.
    /// Engines that cannot host functions are left alone.
    /// </summary>
    public static void Register(IExpressionEngine engine, CultureInfo? culture = null)
    {
        if (!engine.SupportsFunctions)
        {
            return;
        }

        culture ??= CultureInfo.InvariantCulture;

        // The default engine takes the whole library as one shared globals object rather than
        // ~400 separate imports, which is the entire cost of this bridge.
        if (engine is ScribanExpressionEngine scriban)
        {
            scriban.UseExcelFunctions(ScribanGlobals.GetOrAdd(culture, BuildScribanGlobals));
            return;
        }

        // Any other engine — XLibur.Report.DynamicLinq, or a caller's own — is fed one at a time
        // through the interface. The adapters themselves are still shared.
        foreach (var (name, function) in AdaptersFor(culture))
        {
            engine.AddFunction(name, function);
        }
    }

    private static ScriptObject BuildScribanGlobals(CultureInfo culture)
    {
        var globals = new ScriptObject();

        foreach (var (name, function) in AdaptersFor(culture))
        {
            globals.Import(name, function);
        }

        return globals;
    }

    private static IReadOnlyList<(string Name, Delegate Function)> AdaptersFor(CultureInfo culture) =>
        Adapters.GetOrAdd(culture, BuildAdapters);

    /// <summary>
    /// One adapter per registered function, closing over a single calc engine.
    /// </summary>
    /// <remarks>
    /// Sharing that engine across templates — and so across threads — is safe because these calls
    /// have no grid: <see cref="CalcContext.CalcEngine"/> is reached only when reading a cell, which
    /// needs a worksheet, and without one the call has already failed with the
    /// <see cref="MissingContextException"/> translated below. The engine is here to be handed to a
    /// <see cref="CalcContext"/>, not to hold state for the call.
    /// </remarks>
    private static IReadOnlyList<(string Name, Delegate Function)> BuildAdapters(CultureInfo culture)
    {
        var calcEngine = new XLCalcEngine(culture);
        var adapters = new List<(string Name, Delegate Function)>();

        foreach (var name in calcEngine.Functions.Names)
        {
            if (!calcEngine.Functions.TryGetFunc(name, out var definition) || definition is null)
            {
                continue;
            }

            var function = definition;
            adapters.Add((
                name.ToUpperInvariant(),
                new ExcelFunction(arguments => Invoke(calcEngine, function, culture, arguments))));
        }

        return adapters;
    }

    private static object? Invoke(
        XLCalcEngine calcEngine,
        FunctionDefinition definition,
        CultureInfo culture,
        object?[] arguments)
    {
        var flattened = Flatten(arguments, culture);

        if (flattened.Count < definition.MinParams || flattened.Count > definition.MaxParams)
        {
            return XLError.IncompatibleValue;
        }

        // No workbook and no cell: a template expression is evaluated before there is a grid to be
        // relative to. Functions that reach for one throw, and are reported as such below.
        var context = new CalcContext(calcEngine, culture, workbook: null, worksheet: null, formulaAddress: null);
        var values = flattened.ToArray();

        try
        {
            return FromAnyValue(definition.CallFunction(context, values.AsSpan()));
        }
        catch (MissingContextException)
        {
            throw new ExpressionEvaluationException(
                "This function needs a worksheet to work against, which a template expression does not have. " +
                "Use it in a cell formula instead.");
        }
    }

    /// <summary>
    /// Turns the arguments an expression supplied into scalar function arguments, spreading a
    /// collection across several.
    /// </summary>
    /// <remarks>
    /// <c>SUM(items.Total)</c> passes one collection where Excel's <c>SUM</c> expects a run of
    /// arguments, so a collection is spread rather than boxed into an array value. That is what
    /// the aggregate functions — <c>SUM</c>, <c>AVERAGE</c>, <c>COUNT</c>, <c>MIN</c>,
    /// <c>MAX</c> — mean by a range argument, and it keeps the bridge clear of the engine's
    /// internal array representation.
    /// </remarks>
    private static List<AnyValue> Flatten(object?[] arguments, CultureInfo culture)
    {
        var values = new List<AnyValue>(arguments.Length);

        foreach (var argument in arguments)
        {
            if (argument is not string && argument is IEnumerable enumerable)
            {
                foreach (var item in enumerable)
                {
                    values.Add(ToAnyValue(item, culture));
                }

                continue;
            }

            values.Add(ToAnyValue(argument, culture));
        }

        return values;
    }

    private static AnyValue ToAnyValue(object? value, CultureInfo culture) => value switch
    {
        null => ScalarValue.Blank.ToAnyValue(),
        bool logical => AnyValue.From(logical),
        string text => AnyValue.From(text),

        // Excel keeps dates as serial numbers, and its date functions expect them that way.
        DateTime dateTime => AnyValue.From(dateTime.ToOADate()),
        DateTimeOffset dateTimeOffset => AnyValue.From(dateTimeOffset.DateTime.ToOADate()),
        TimeSpan timeSpan => AnyValue.From(timeSpan.TotalDays),

        XLError error => AnyValue.From(error),
        XLCellValue cellValue => ((ScalarValue)cellValue).ToAnyValue(),

        IConvertible convertible => FromConvertible(convertible, culture),
        _ => AnyValue.From(value.ToString() ?? string.Empty),
    };

    /// <summary>
    /// A number where the value converts to one, its text otherwise.
    /// </summary>
    /// <remarks>
    /// <see cref="IConvertible"/> is not a promise of being a number: <see cref="char"/> and
    /// <see cref="DBNull"/> implement it and throw from <c>ToDouble</c>. Letting that throw would
    /// abort the whole run over one cell's argument, which is the opposite of what generation
    /// promises — so a value that will not convert is passed as its text, the same as anything else
    /// with no numeric form.
    /// </remarks>
    private static AnyValue FromConvertible(IConvertible convertible, CultureInfo culture)
    {
        try
        {
            return AnyValue.From(convertible.ToDouble(culture));
        }
        catch (Exception ex) when (ex is InvalidCastException or FormatException or OverflowException)
        {
            return AnyValue.From(convertible.ToString(culture) ?? string.Empty);
        }
    }

    private static object? FromAnyValue(AnyValue value)
    {
        if (!value.TryPickScalar(out var scalar, out _))
        {
            // An array or reference result has no single value a cell could hold.
            return XLError.IncompatibleValue;
        }

        return scalar.Match<object?>(
            () => null,
            logical => logical,
            number => number,
            text => text,
            error => error);
    }
}
