using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using Scriban;
using Scriban.Parsing;
using Scriban.Runtime;
using Scriban.Syntax;

namespace XLibur.Report.Expressions;

/// <summary>
/// The default expression engine, evaluating template expressions with
/// <see href="https://github.com/scriban/scriban">Scriban</see>.
/// </summary>
/// <remarks>
/// Expressions are parsed once and cached, then evaluated per data row — the shape range
/// expansion needs, where one cell's expression runs against thousands of items.
/// <para>
/// Three deliberate departures from Scriban's defaults:
/// <list type="bullet">
/// <item>Members keep their C# names (Scriban would rename <c>UnitPrice</c> to
/// <c>unit_price</c>), so <c>{{ item.UnitPrice }}</c> binds the way a template author expects.</item>
/// <item>Relaxed member, target and indexer access are on, so a missing property or a null
/// target yields a blank cell instead of throwing — report data is routinely sparse.</item>
/// <item>Evaluation is culture-invariant by default, so a generated report does not change
/// shape with the machine's locale. Pass a culture to the constructor to override.</item>
/// </list>
/// Scriban's execution limits (loop, recursion, output size) are left at their defaults, which
/// is what makes this engine safe to point at templates you do not control.
/// </para>
/// <para>This type is not thread-safe; a template generates on one thread at a time.</para>
/// </remarks>
public sealed class ScribanExpressionEngine : IExpressionEngine
{
    private static readonly LexerOptions ExpressionLexerOptions = new() { Mode = ScriptMode.ScriptOnly };

    private readonly ConcurrentDictionary<string, Template> _expressionCache = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, Template> _textCache = new(StringComparer.Ordinal);
    private readonly ConditionalWeakTable<ExpressionScope, ScriptObject> _scopeObjects = new();
    private readonly ScriptObject _functions = new();
    private readonly CultureInfo _culture;
    private TemplateContext? _context;

    /// <summary>Creates an engine evaluating under <paramref name="culture"/>, invariant by default.</summary>
    public ScribanExpressionEngine(CultureInfo? culture = null) => _culture = culture ?? CultureInfo.InvariantCulture;

    /// <inheritdoc />
    public CultureInfo Culture => _culture;

    /// <inheritdoc />
    public bool SupportsFunctions => true;

    /// <inheritdoc />
    public void AddFunction(string name, Delegate function)
    {
        if (string.IsNullOrEmpty(name))
        {
            throw new ArgumentException("A function name is required.", nameof(name));
        }

        if (function is null)
        {
            throw new ArgumentNullException(nameof(function));
        }

        _functions.Import(name, function);
    }

    /// <inheritdoc />
    public object? Evaluate(string expression, ExpressionScope scope)
    {
        if (expression is null)
        {
            throw new ArgumentNullException(nameof(expression));
        }

        var template = GetTemplate(_expressionCache, expression, expression, ExpressionLexerOptions);
        return Run(template, expression, scope, static (t, ctx) => t.Evaluate(ctx));
    }

    /// <inheritdoc />
    public string Interpolate(string text, ExpressionScope scope)
    {
        if (text is null)
        {
            throw new ArgumentNullException(nameof(text));
        }

        var template = GetTemplate(_textCache, text, text, lexerOptions: null);
        return Run(template, text, scope, static (t, ctx) => t.Render(ctx)) ?? string.Empty;
    }

    private T Run<T>(Template template, string source, ExpressionScope scope, Func<Template, TemplateContext, T> run)
    {
        var context = GetContext();
        var pushed = PushScopes(context, scope);

        try
        {
            return run(template, context);
        }
        catch (ScriptRuntimeException ex)
        {
            throw new ExpressionEvaluationException(source, ex.Message, ex);
        }
        finally
        {
            for (var i = 0; i < pushed; i++)
            {
                context.PopGlobal();
            }
        }
    }

    private TemplateContext GetContext()
    {
        if (_context is not null)
        {
            return _context;
        }

        var context = new TemplateContext
        {
            // Keep C# member names verbatim; Scriban would otherwise expose UnitPrice as unit_price.
            MemberRenamer = member => member.Name,
            EnableRelaxedMemberAccess = true,
            EnableRelaxedTargetAccess = true,
            EnableRelaxedIndexerAccess = true,
            StrictVariables = false,
        };

        context.PushCulture(_culture);

        // Pushed once and never popped, so functions registered later through AddFunction are
        // still visible: the same ScriptObject instance stays on the stack.
        context.PushGlobal(_functions);

        _context = context;
        return context;
    }

    private int PushScopes(TemplateContext context, ExpressionScope scope)
    {
        var chain = scope.FromOutermost();
        var pushed = 0;

        foreach (var level in chain)
        {
            if (level.Values.Count == 0)
            {
                continue;
            }

            context.PushGlobal(GetScriptObject(level));
            pushed++;
        }

        return pushed;
    }

    private ScriptObject GetScriptObject(ExpressionScope scope)
    {
        if (_scopeObjects.TryGetValue(scope, out var existing))
        {
            return existing;
        }

        var scriptObject = new ScriptObject();
        foreach (var pair in scope.Values)
        {
            scriptObject[pair.Key] = pair.Value;
        }

        _scopeObjects.Add(scope, scriptObject);
        return scriptObject;
    }

    private static Template GetTemplate(
        ConcurrentDictionary<string, Template> cache,
        string key,
        string source,
        LexerOptions? lexerOptions)
    {
        var template = cache.GetOrAdd(key, _ => Template.Parse(source, lexerOptions: lexerOptions));

        if (template.HasErrors)
        {
            var message = string.Join("; ", template.Messages.Select(m => m.Message));
            throw new ExpressionEvaluationException(source, message);
        }

        return template;
    }
}
