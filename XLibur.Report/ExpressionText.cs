using System;
using System.Text.RegularExpressions;

namespace XLibur.Report;

/// <summary>
/// Recognises the <c>{{ }}</c> expression markers a template embeds in cell text.
/// </summary>
internal static class ExpressionText
{
    /// <summary>The prefix marking a cell whose evaluated text becomes a formula.</summary>
    public const string FormulaPrefix = "&=";

    private static readonly Regex Pattern = new(
        @"\{\{.+?\}\}",
        RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.CultureInvariant,
        TimeSpan.FromSeconds(5));

    /// <summary>Whether <paramref name="text"/> embeds at least one expression.</summary>
    public static bool Contains(string? text) => !string.IsNullOrEmpty(text) && Pattern.IsMatch(text);

    /// <summary>
    /// Whether the whole of <paramref name="text"/> is a single expression, in which case
    /// <paramref name="expression"/> receives its body without the delimiters.
    /// </summary>
    /// <remarks>
    /// The distinction decides the cell's type. A cell holding only <c>{{ item.Total }}</c> takes
    /// the expression's value verbatim, so a decimal stays a number; a cell holding
    /// <c>Total: {{ item.Total }}</c> can only be text.
    /// </remarks>
    public static bool TryGetSingleExpression(string? text, out string expression)
    {
        expression = string.Empty;

        if (string.IsNullOrEmpty(text))
        {
            return false;
        }

        var trimmed = text!.Trim();
        if (trimmed.Length < 5 || !trimmed.StartsWith("{{", StringComparison.Ordinal) || !trimmed.EndsWith("}}", StringComparison.Ordinal))
        {
            return false;
        }

        var match = Pattern.Match(trimmed);
        if (!match.Success || match.Index != 0 || match.Length != trimmed.Length)
        {
            return false;
        }

        expression = trimmed.Substring(2, trimmed.Length - 4).Trim();
        return expression.Length > 0;
    }
}
