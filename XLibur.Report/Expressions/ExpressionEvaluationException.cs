using System;

namespace XLibur.Report.Expressions;

/// <summary>
/// Thrown by an <see cref="IExpressionEngine"/> when an expression cannot be parsed or evaluated.
/// </summary>
/// <remarks>
/// Generation catches this per cell and records a <see cref="TemplateError"/> rather than
/// aborting, so one bad expression cannot cost the whole report.
/// </remarks>
public class ExpressionEvaluationException : Exception
{
    /// <summary>Creates an exception describing a failure in <paramref name="expression"/>.</summary>
    public ExpressionEvaluationException(string expression, string message, Exception? innerException = null)
        : base(message, innerException)
    {
        Expression = expression;
    }

    /// <inheritdoc cref="ExpressionEvaluationException"/>
    public ExpressionEvaluationException(string message)
        : base(message)
    {
        Expression = string.Empty;
    }

    /// <summary>The expression that failed, without its <c>{{ }}</c> delimiters.</summary>
    public string Expression { get; }
}
