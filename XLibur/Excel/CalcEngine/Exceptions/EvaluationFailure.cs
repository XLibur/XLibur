using System;

namespace XLibur.Excel.CalcEngine.Exceptions;

/// <summary>
/// Tells the failures a formula can legitimately produce from defects in the library.
/// </summary>
/// <remarks>
/// A few read APIs - <c>IXLCell.TryGetValue</c>, <c>IXLCell.GetFormattedString</c> and
/// <c>IXLRangeBase.Search</c> - promise an answer for a cell whose formula cannot be evaluated. They
/// used to keep that promise by catching everything, which also turned a defect inside a function
/// into "this cell has no value". Only the failures listed here are a property of the formula;
/// anything else reaches the caller.
/// </remarks>
internal static class EvaluationFailure
{
    internal static bool IsExpected(Exception exception) => exception is
        // A function, operator or formula type the engine does not evaluate.
        NotImplementedException or NotSupportedException
        // Text the parser cannot read.
        or ExpressionParseException
        // Evaluated without the worksheet or cell address it needs.
        or MissingContextException or XLNoWorksheetContextException
        // A formula that depends on its own value.
        or CircularReferenceException;
}
