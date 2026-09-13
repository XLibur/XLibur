using System;
using XLibur.Excel.CalcEngine.Exceptions;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// The ways a public caller reaches the calc engine. Each decides what a failure looks like to its
/// caller by reading <see cref="EvaluationPolicy"/>.
/// </summary>
internal enum EvaluationEntryPoint
{
    /// <summary><c>IXLCell.Value</c>, and everything that reads it.</summary>
    CellValue,

    /// <summary>
    /// <c>IXLCell.TryGetValue</c>, <c>IXLCell.GetFormattedString</c> and <c>IXLRangeBase.Search</c>,
    /// which promise an answer for a cell whose formula cannot be evaluated.
    /// </summary>
    TolerantRead,

    /// <summary><c>IXLWorksheet.Evaluate</c>, <c>IXLWorkbook.Evaluate</c> and <c>XLWorkbook.EvaluateExpr</c>.</summary>
    Evaluate,

    /// <summary><c>XLFunctionLibrary.TryInvoke</c>.</summary>
    FunctionLibrary,

    /// <summary>
    /// <c>RecalculateAllFormulas</c> on a workbook or a sheet, and <c>LoadOptions.RecalculateAllFormulas</c>.
    /// </summary>
    Recalculation,

    /// <summary>Save, with <c>SaveOptions.EvaluateFormulasBeforeSaving</c>.</summary>
    Save,
}

/// <summary>
/// What an entry point does with a failure of one kind.
/// </summary>
internal enum EvaluationOutcome
{
    /// <summary>
    /// The failure reaches the caller as its public exception: <see cref="XLCircularReferenceException"/>,
    /// <see cref="NotImplementedException"/>, <see cref="ExpressionParseException"/>,
    /// <see cref="XLNoWorksheetContextException"/>, or a defect as it was raised.
    /// </summary>
    Throw,

    /// <summary>
    /// The call answers "no value": <c>TryGetValue</c> returns <c>false</c>, <c>GetFormattedString</c>
    /// shows the cached value, and <c>Search</c> does not match the cell.
    /// </summary>
    NoValue,

    /// <summary>
    /// The cell gets no new value and stays dirty, and the entry point carries on with the other
    /// cells. Save writes the cell with no cached value, so Excel recalculates it on open.
    /// </summary>
    LeaveDirty,
}

/// <summary>
/// The one table that maps an entry point and a kind of failure to what the caller sees (spec 56).
/// </summary>
/// <remarks>
/// <para>
/// A failure is sorted into its kind once, by <see cref="EvaluationFailure.Classify"/>, and every
/// entry point reads its row here instead of deciding for itself. A change to what a caller sees
/// is a change to this table, and <c>EvaluationOutcomeTests</c> observes every cell of it from
/// outside.
/// </para>
/// <para>
/// <see cref="EvaluationFailureKind.Pending"/> never reaches a public caller: every entry point
/// calculates the dirty precedent first. Its column says what would happen if one did, which is
/// what happens to a defect.
/// </para>
/// </remarks>
internal static class EvaluationPolicy
{
    internal static EvaluationOutcome For(EvaluationEntryPoint entry, Exception exception) =>
        For(entry, EvaluationFailure.Classify(exception));

    internal static EvaluationOutcome For(EvaluationEntryPoint entry, EvaluationFailureKind kind) => entry switch
    {
        EvaluationEntryPoint.TolerantRead => kind switch
        {
            // #459: only what a formula can legitimately produce is tolerated.
            EvaluationFailureKind.Cycle
                or EvaluationFailureKind.Unsupported
                or EvaluationFailureKind.Refused
                or EvaluationFailureKind.NoContext => EvaluationOutcome.NoValue,
            _ => EvaluationOutcome.Throw,
        },

        // ADR 0001: an expected failure writes no cached value, and Excel recalculates the cell on
        // open. Anything else throws out of the save, a defect above all (D61).
        EvaluationEntryPoint.Save => kind switch
        {
            EvaluationFailureKind.Cycle
                or EvaluationFailureKind.Unsupported
                or EvaluationFailureKind.Refused => EvaluationOutcome.LeaveDirty,
            _ => EvaluationOutcome.Throw,
        },

        _ => EvaluationOutcome.Throw,
    };

    /// <summary>
    /// Runs <paramref name="evaluate"/>, and reports the calc engine's internal missing-context
    /// signal as the public <see cref="XLNoWorksheetContextException"/>, with the message
    /// <paramref name="explain"/> gives.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The only place that translation is written. <see cref="MissingContextException"/> is internal,
    /// so letting it out of a public entry point hands the caller an exception they cannot name, let
    /// alone catch (D37).
    /// </para>
    /// <para>
    /// The state is passed in rather than captured, so a caller can use static lambdas and a cell's
    /// evaluation allocates nothing for the translation.
    /// </para>
    /// </remarks>
    internal static TResult RaiseMissingContextAsPublic<TState, TResult>(
        TState state,
        Func<TState, TResult> evaluate,
        Func<TState, string> explain)
    {
        try
        {
            return evaluate(state);
        }
        catch (MissingContextException e)
        {
            throw new XLNoWorksheetContextException(explain(state), e);
        }
    }
}
