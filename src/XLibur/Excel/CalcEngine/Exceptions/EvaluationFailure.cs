using System;
using System.Diagnostics.CodeAnalysis;

namespace XLibur.Excel.CalcEngine.Exceptions;

/// <summary>
/// The kinds of failure an evaluation can end in.
/// </summary>
internal enum EvaluationFailureKind
{
    /// <summary>A formula depends on its own value: a circular reference.</summary>
    Cycle,

    /// <summary>A valid Excel construct that XLibur does not evaluate: an unsupported feature.</summary>
    Unsupported,

    /// <summary>Formula text the parser cannot read: a refused formula.</summary>
    Refused,

    /// <summary>Evaluated without the worksheet or cell address it needs.</summary>
    NoContext,

    /// <summary>
    /// A precedent is still dirty. The calc engine's own signal to calculate the precedent first; it
    /// must never reach a public caller.
    /// </summary>
    Pending,

    /// <summary>A bug in XLibur rather than something in the workbook: a defect.</summary>
    Defect,
}

/// <summary>
/// Sorts every evaluation failure into one <see cref="EvaluationFailureKind"/>.
/// </summary>
/// <remarks>
/// A kind is known by its exception type, and each type means one kind. That is why an unsupported
/// feature is raised as <see cref="UnsupportedFeatureException"/> rather than as a plain
/// <see cref="NotImplementedException"/>: a plain one, thrown anywhere in the library, is a defect.
/// What an entry point does with each kind is <see cref="EvaluationPolicy"/>'s business, not this
/// class's.
/// </remarks>
internal static class EvaluationFailure
{
    internal static EvaluationFailureKind Classify(Exception exception) => exception switch
    {
        XLCircularReferenceException => EvaluationFailureKind.Cycle,

        // NotSupportedException is recorded as it was classified before spec 56. The calc engine's
        // own gaps are all UnsupportedFeatureException.
        UnsupportedFeatureException or NotSupportedException => EvaluationFailureKind.Unsupported,
        ExpressionParseException => EvaluationFailureKind.Refused,
        MissingContextException or XLNoWorksheetContextException => EvaluationFailureKind.NoContext,
        GettingDataException => EvaluationFailureKind.Pending,
        _ => EvaluationFailureKind.Defect,
    };

    /// <summary>
    /// Whether <paramref name="kind"/> is an expected failure (ADR 0001): something in the workbook's
    /// content that save and the public <c>TryEvaluate</c> methods tolerate, as opposed to a defect or
    /// the engine's own signals. It is the one definition of "expected": a failure is expected exactly
    /// when it has a <see cref="XLFormulaIssueKind"/>, so adding one is a change to
    /// <see cref="TryGetIssueKind"/> only.
    /// </summary>
    internal static bool IsExpected(EvaluationFailureKind kind) => TryGetIssueKind(kind, out _);

    /// <summary>
    /// The public name of an expected failure.
    /// </summary>
    /// <returns><c>false</c> when <paramref name="kind"/> is not an expected failure.</returns>
    internal static bool TryGetIssueKind(EvaluationFailureKind kind, out XLFormulaIssueKind issueKind)
    {
        switch (kind)
        {
            case EvaluationFailureKind.Cycle:
                issueKind = XLFormulaIssueKind.CircularReference;
                return true;
            case EvaluationFailureKind.Unsupported:
                issueKind = XLFormulaIssueKind.UnsupportedFeature;
                return true;
            case EvaluationFailureKind.Refused:
                issueKind = XLFormulaIssueKind.RefusedFormula;
                return true;
            default:
                issueKind = default;
                return false;
        }
    }

    /// <summary>
    /// Describes an expected failure as the public issue a caller of <c>TryEvaluate</c> gets.
    /// </summary>
    /// <returns><c>false</c> when <paramref name="exception"/> is not an expected failure.</returns>
    internal static bool TryCreateIssue(Exception exception, [NotNullWhen(true)] out XLFormulaIssue? issue)
    {
        if (TryGetIssueKind(Classify(exception), out var issueKind))
        {
            issue = new XLFormulaIssue(issueKind, exception.Message, exception);
            return true;
        }

        issue = null;
        return false;
    }
}
