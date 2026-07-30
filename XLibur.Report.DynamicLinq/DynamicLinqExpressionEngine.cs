using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Dynamic.Core;
using System.Linq.Dynamic.Core.Exceptions;
using System.Linq.Expressions;
using System.Text;
using XLibur.Report.Expressions;

namespace XLibur.Report.DynamicLinq;

/// <summary>
/// An expression engine that runs <b>ClosedXML.Report's</b> expression syntax — C# expressions,
/// lambdas and LINQ inside <c>{{ }}</c> — so templates written for that library run unmodified.
/// </summary>
/// <remarks>
/// <para>
/// A template's <em>structure</em> is engine-independent: defined-name-bound ranges, the options row,
/// <c>&lt;&lt;tags&gt;&gt;</c> and the <c>&amp;=</c> formula prefix mean the same thing whichever engine
/// evaluates the expressions. So the expression language is the only difference between an
/// upstream-authored template and an XLibur.Report one, and this is that difference, in one class.
/// </para>
/// <code>
/// using var template = new XLTemplate("LegacyReport.xlsx", new DynamicLinqExpressionEngine());
/// template.AddVariable("Items", sales);
/// template.Generate();
/// </code>
/// <para>
/// <b>What it supports.</b> Anything
/// <see href="https://dynamic-linq.net/expression-language">Dynamic LINQ</see> parses: property and
/// method access (<c>item.Name.ToUpper()</c>), arithmetic and comparison, the conditional operator,
/// and LINQ over collections in scope (<c>items.Where(x =&gt; x.Qty &gt; 10).Sum(x =&gt; x.Total)</c>).
/// Inside a bound range the names are <c>item</c>, <c>index</c> and <c>items</c>, and a workbook
/// variable may be reached with an <c>@</c> prefix (<c>@Company</c>) as upstream allowed, as well as by
/// its plain name.
/// </para>
/// <para>
/// <b>What it does not.</b> <see cref="SupportsFunctions"/> is <c>false</c>: the Excel-function bridge
/// is a feature of the default engine, upstream syntax never had it, and Dynamic LINQ's static-method
/// extension mechanism is not a good place to put four hundred Excel functions. Templates for this
/// engine use .NET methods, which is what their authors wrote.
/// </para>
/// <para>
/// <b>Security: trusted templates only.</b> Dynamic LINQ has no sandbox. An expression it parses can
/// reach the methods and properties of any object in scope, and the library's own history includes
/// CVE-2023-32571 — arbitrary method invocation — so the version floor here is not negotiable. Do not
/// point this engine at a template a user uploaded. The default
/// <see cref="ScribanExpressionEngine"/> is the engine for that: Scriban has real execution limits and
/// no reflection escape, which is why it is the default and this is opt-in.
/// </para>
/// <para>This type is not thread-safe; a template generates on one thread at a time.</para>
/// </remarks>
public sealed class DynamicLinqExpressionEngine : IExpressionEngine
{
    /// <summary>
    /// Parsed expressions, keyed by the expression text <em>and</em> the shape of the scope it was
    /// parsed against.
    /// </summary>
    /// <remarks>
    /// The key has to include the parameter names and types, not just the text: Dynamic LINQ compiles a
    /// lambda over a fixed parameter list, so the same expression evaluated against a different set of
    /// names is a different lambda. Range expansion evaluates one expression against thousands of rows
    /// with the same scope shape each time, which is what makes caching worth having at all.
    /// </remarks>
    private readonly ConcurrentDictionary<string, Delegate> _lambdas = new(StringComparer.Ordinal);

    private readonly ParsingConfig _config;
    private readonly CultureInfo _culture;

    /// <summary>Creates an engine evaluating under <paramref name="culture"/>, invariant by default.</summary>
    /// <remarks>
    /// Invariant by default for the same reason the default engine is: a report should not change shape
    /// with the machine's locale, and a number interpolated into a formula with a comma for a decimal
    /// point is not a formula.
    /// </remarks>
    public DynamicLinqExpressionEngine(CultureInfo? culture = null)
    {
        _culture = culture ?? CultureInfo.InvariantCulture;

        _config = new ParsingConfig
        {
            // Upstream templates say item.Name, not item.name.
            IsCaseSensitive = true,

            // The template author wrote C#, so let them index and call as they would in C#.
            AllowNewToEvaluateAnyType = false,
        };
    }

    /// <inheritdoc />
    /// <remarks>
    /// Always <c>false</c>. See the remarks on <see cref="DynamicLinqExpressionEngine"/>: the
    /// Excel-function bridge belongs to the default engine, and <c>XLTemplate</c> skips registering it
    /// for engines that decline.
    /// </remarks>
    public bool SupportsFunctions => false;

    /// <inheritdoc />
    /// <exception cref="NotSupportedException">Always.</exception>
    public void AddFunction(string name, Delegate function) =>
        throw new NotSupportedException(
            "DynamicLinqExpressionEngine does not host functions. It runs ClosedXML.Report's syntax, "
            + "where a template calls .NET methods directly; the Excel-function bridge is a feature of "
            + "the default Scriban engine.");

    /// <inheritdoc />
    public object? Evaluate(string expression, ExpressionScope scope)
    {
        if (expression is null)
        {
            throw new ArgumentNullException(nameof(expression));
        }

        if (scope is null)
        {
            throw new ArgumentNullException(nameof(scope));
        }

        var names = Flatten(scope);

        return Invoke(expression, names);
    }

    /// <inheritdoc />
    public string Interpolate(string text, ExpressionScope scope)
    {
        if (text is null)
        {
            throw new ArgumentNullException(nameof(text));
        }

        if (scope is null)
        {
            throw new ArgumentNullException(nameof(scope));
        }

        var names = Flatten(scope);
        var result = new StringBuilder(text.Length);
        var index = 0;

        while (index < text.Length)
        {
            var open = text.IndexOf("{{", index, StringComparison.Ordinal);
            if (open < 0)
            {
                result.Append(text, index, text.Length - index);
                break;
            }

            var close = text.IndexOf("}}", open + 2, StringComparison.Ordinal);
            if (close < 0)
            {
                // An unclosed marker is literal text, not an error: a report is better off showing it
                // than failing over it, and the author will see it in the output.
                result.Append(text, index, text.Length - index);
                break;
            }

            result.Append(text, index, open - index);

            var expression = text.Substring(open + 2, close - open - 2).Trim();
            if (expression.Length > 0)
            {
                result.Append(Format(Invoke(expression, names)));
            }

            index = close + 2;
        }

        return result.ToString();
    }

    /// <summary>
    /// Flattens the scope chain into one name-to-value map, innermost winning, and adds the
    /// <c>@</c>-prefixed aliases upstream templates use for workbook variables.
    /// </summary>
    /// <remarks>
    /// Dynamic LINQ takes a flat parameter list rather than a chain of scopes, so the shadowing the
    /// scope chain expresses has to be resolved here. Outermost first, so an inner scope's
    /// <c>item</c> overwrites an outer one's.
    /// </remarks>
    private static List<Binding> Flatten(ExpressionScope scope)
    {
        var byName = new Dictionary<string, Binding>(StringComparer.Ordinal);
        var order = new List<string>();

        foreach (var level in scope.FromOutermost())
        {
            foreach (var pair in level.Values)
            {
                var binding = Binding.For(pair.Key, pair.Value);

                Put(byName, order, binding);

                // Upstream wrote @Company for a workbook variable read from inside a range. Both forms
                // work here: the alias is added rather than the plain name being replaced, because a
                // template mixing the two is common and neither is wrong.
                Put(byName, order, binding with { Name = "@" + pair.Key });
            }
        }

        return order.Select(name => byName[name]).ToList();
    }

    /// <summary>
    /// Records a binding, keeping the order names were first seen in so that the parameter list and the
    /// argument list stay aligned however the scopes shadow each other.
    /// </summary>
    private static void Put(Dictionary<string, Binding> byName, List<string> order, Binding binding)
    {
        if (!byName.ContainsKey(binding.Name))
        {
            order.Add(binding.Name);
        }

        byName[binding.Name] = binding;
    }

    private object? Invoke(string expression, List<Binding> bindings)
    {
        try
        {
            var lambda = GetLambda(expression, bindings);

            return lambda.DynamicInvoke(bindings.Select(binding => binding.Value).ToArray());
        }
        catch (Exception ex) when (ex is ParseException or InvalidOperationException or ArgumentException)
        {
            throw new ExpressionEvaluationException(expression, ex.Message, ex);
        }
        catch (System.Reflection.TargetInvocationException ex)
        {
            // What DynamicInvoke wraps anything the expression itself threw in — a null reference
            // inside item.Customer.Name, a divide by zero, a method that failed.
            var inner = ex.InnerException ?? ex;
            throw new ExpressionEvaluationException(expression, inner.Message, inner);
        }
    }

    private Delegate GetLambda(string expression, List<Binding> bindings)
    {
        var parameters = bindings
            .Select(binding => Expression.Parameter(binding.Type, binding.Name))
            .ToArray();

        var key = CacheKey(expression, parameters);

        return _lambdas.GetOrAdd(
            key,
            _ => DynamicExpressionParser.ParseLambda(_config, parameters, resultType: null, expression).Compile());
    }

    /// <summary>
    /// One name in scope, as Dynamic LINQ needs it: the type to compile the parameter as, and the value
    /// to pass for it.
    /// </summary>
    /// <remarks>
    /// The two travel together because they have to agree. Compiling against the value's concrete type
    /// is what makes <c>item.Name.ToUpper()</c> and <c>items.Sum(x =&gt; x.Total)</c> resolve at all —
    /// against <see cref="object"/> there is no <c>Name</c> and no <c>Sum</c> — and the invoked delegate
    /// then rejects any argument that is not of exactly that type.
    /// </remarks>
    private readonly record struct Binding(string Name, Type Type, object? Value)
    {
        public static Binding For(string name, object? value)
        {
            if (value is null)
            {
                // Nothing to ask for a type. It compiles as object, and a member access on it fails the
                // way it would in C#.
                return new Binding(name, typeof(object), null);
            }

            var type = value.GetType();

            // A non-generic IEnumerable gives LINQ nothing to work with: no element type means no
            // Sum(x => …). Re-cast to the element type its first item reports — which is what upstream
            // did for DataTable rows and untyped collections — and materialise it, because the cast is
            // lazy and the same collection is enumerated once per row.
            if (value is IEnumerable enumerable && type != typeof(string) && !IsGenericEnumerable(type))
            {
                var elementType = ElementType(enumerable);
                if (elementType is not null)
                {
                    return new Binding(
                        name,
                        typeof(IEnumerable<>).MakeGenericType(elementType),
                        CastTo(enumerable, elementType));
                }
            }

            return new Binding(name, type, value);
        }
    }

    /// <summary>
    /// <c>Enumerable.Cast&lt;T&gt;().ToList()</c>, reached by reflection because <c>T</c> is only known
    /// at runtime.
    /// </summary>
    private static object CastTo(IEnumerable enumerable, Type elementType)
    {
        var cast = typeof(Enumerable).GetMethod(nameof(Enumerable.Cast))!.MakeGenericMethod(elementType);
        var toList = typeof(Enumerable).GetMethod(nameof(Enumerable.ToList))!.MakeGenericMethod(elementType);

        return toList.Invoke(null, new[] { cast.Invoke(null, new object[] { enumerable }) })!;
    }

    private static bool IsGenericEnumerable(Type type) =>
        type.GetInterfaces().Any(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>));

    private static Type? ElementType(IEnumerable enumerable)
    {
        foreach (var item in enumerable)
        {
            if (item is not null)
            {
                return item.GetType();
            }
        }

        return null;
    }

    private static string CacheKey(string expression, ParameterExpression[] parameters)
    {
        var key = new StringBuilder(expression.Length + (parameters.Length * 24));
        key.Append(expression);

        // A separator that cannot occur in an expression or a type name, so no two different scope
        // shapes can produce the same key by concatenation.
        foreach (var parameter in parameters)
        {
            key.Append('\u001F').Append(parameter.Name).Append(':').Append(parameter.Type.FullName);
        }

        return key.ToString();
    }

    /// <summary>
    /// Renders a value for interpolation into text or into a formula.
    /// </summary>
    /// <remarks>
    /// Culture-invariant, because the other use of this is building a formula and a formula cannot
    /// contain a locale's decimal comma. A <see cref="DateTime"/> becomes its OLE Automation serial for
    /// the same reason — that is the only form a formula can do arithmetic on, and it is what upstream
    /// produced.
    /// </remarks>
    private string Format(object? value) => value switch
    {
        null => string.Empty,
        DateTime date => date.ToOADate().ToString(_culture),
        IFormattable formattable => formattable.ToString(null, _culture),
        _ => value.ToString() ?? string.Empty,
    };
}
