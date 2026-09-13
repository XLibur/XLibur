using System;

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
/// What an entry point does with each kind is <c>EvaluationPolicy</c>'s business, not this class's.
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
    /// Whether <paramref name="exception"/> is a failure a formula can legitimately produce, rather
    /// than a defect or the engine's own pending signal.
    /// </summary>
    internal static bool IsExpected(Exception exception) =>
        Classify(exception) is not (EvaluationFailureKind.Defect or EvaluationFailureKind.Pending);
}
