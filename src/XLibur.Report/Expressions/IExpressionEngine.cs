using System;
using System.Globalization;

namespace XLibur.Report.Expressions;

/// <summary>
/// Evaluates the expressions a template embeds in <c>{{ }}</c>. The implementation defines the
/// template expression language.
/// </summary>
/// <remarks>
/// XLibur.Report ships <see cref="ScribanExpressionEngine"/> as the default. The separate
/// <c>XLibur.Report.DynamicLinq</c> package supplies an engine that runs ClosedXML.Report's C#
/// expression syntax instead.
/// <para>
/// Implementations are not required to be thread-safe: a template generates on one thread at a
/// time, and the engine may cache mutable parse state.
/// </para>
/// </remarks>
public interface IExpressionEngine
{
    /// <summary>
    /// The culture the report is generated under — how a value formats, and how two text keys
    /// collate.
    /// </summary>
    /// <remarks>
    /// The engine owns the answer because it is the thing that already had to be told: both shipped
    /// engines take a culture and default to <see cref="CultureInfo.InvariantCulture"/>. Exposing it
    /// is what lets the rest of the pipeline — sort order, group order, a group's <c>{0} Total</c>
    /// label — follow the culture the <em>template</em> was given rather than the one the machine
    /// happens to be running under, so the same template and the same data produce the same workbook
    /// wherever it is generated.
    /// </remarks>
    CultureInfo Culture { get; }

    /// <summary>
    /// Whether this engine can expose .NET functions to templates through
    /// <see cref="AddFunction"/>. Engines that return <c>false</c> are skipped when the Excel
    /// function bridge is registered.
    /// </summary>
    bool SupportsFunctions { get; }

    /// <summary>
    /// Evaluates a single expression — the text between <c>{{</c> and <c>}}</c>, without the
    /// delimiters — and returns the raw result, preserving its .NET type so a decimal reaches
    /// the cell as a number rather than as text.
    /// </summary>
    /// <exception cref="ExpressionEvaluationException">The expression failed to parse or to evaluate.</exception>
    object? Evaluate(string expression, ExpressionScope scope);

    /// <summary>
    /// Renders text containing zero or more <c>{{ }}</c> expressions interleaved with literal
    /// text, returning the whole thing as a string.
    /// </summary>
    /// <exception cref="ExpressionEvaluationException">The text failed to parse or to evaluate.</exception>
    string Interpolate(string text, ExpressionScope scope);

    /// <summary>
    /// Exposes a .NET function to templates under <paramref name="name"/>.
    /// </summary>
    /// <exception cref="NotSupportedException"><see cref="SupportsFunctions"/> is <c>false</c>.</exception>
    void AddFunction(string name, Delegate function);
}
