using System;
using System.Collections.Generic;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;

namespace XLibur.Extensions;

internal static class XLErrorExtensions
{
    public static string ToDisplayString(this XLError error) =>
        error switch
        {
            XLError.CellReference => "#REF!",
            XLError.IncompatibleValue => "#VALUE!",
            XLError.DivisionByZero => "#DIV/0!",
            XLError.NameNotRecognized => "#NAME?",
            XLError.NoValueAvailable => "#N/A",
            XLError.NullValue => "#NULL!",
            XLError.NumberInvalid => "#NUM!",
            XLError.GettingData => "#GETTING_DATA",
            XLError.SpillRange => "#SPILL!",
            XLError.Connect => "#CONNECT!",
            XLError.Blocked => "#BLOCKED!",
            XLError.Unknown => "#UNKNOWN!",
            XLError.Field => "#FIELD!",
            XLError.Calc => "#CALC!",
            XLError.Busy => "#BUSY!",
            XLError.External => "#EXTERNAL!",
            XLError.Timeout => "#TIMEOUT!",
            XLError.Python => "#PYTHON!",
            _ => throw new ArgumentOutOfRangeException(nameof(error), error, "Unknown XLError value.")
        };
}

internal static class XLErrorParser
{
    private static readonly Dictionary<string, XLError> ErrorMap = new(StringComparer.Ordinal)
    {
        ["#REF!"] = XLError.CellReference,
        ["#VALUE!"] = XLError.IncompatibleValue,
        ["#DIV/0!"] = XLError.DivisionByZero,
        ["#NAME?"] = XLError.NameNotRecognized,
        ["#N/A"] = XLError.NoValueAvailable,
        ["#NULL!"] = XLError.NullValue,
        ["#NUM!"] = XLError.NumberInvalid,
        ["#GETTING_DATA"] = XLError.GettingData,
        ["#SPILL!"] = XLError.SpillRange,
        ["#CONNECT!"] = XLError.Connect,
        ["#BLOCKED!"] = XLError.Blocked,
        ["#UNKNOWN!"] = XLError.Unknown,
        ["#FIELD!"] = XLError.Field,
        ["#CALC!"] = XLError.Calc,
        ["#BUSY!"] = XLError.Busy,
        ["#EXTERNAL!"] = XLError.External,
        ["#TIMEOUT!"] = XLError.Timeout,
        ["#PYTHON!"] = XLError.Python
    };

    public static bool TryParseError(string input, out XLError error)
        => ErrorMap.TryGetValue(input.Trim(), out error);

    /// <summary>
    /// The <see cref="XLError"/> an error literal in a formula stands for. <see cref="XLError"/> has a
    /// member for every error value the parser reads. Should a later parser read one it lacks, a formula
    /// holding it is refused like unreadable text.
    /// </summary>
    /// <exception cref="ExpressionParseException">The literal has no <see cref="XLError"/>.</exception>
    public static XLError ParseFormulaError(ReadOnlySpan<char> error)
    {
        var text = error.ToString();
        if (!TryParseError(text, out var parsed))
            throw new ExpressionParseException($"'{text}' is not an error value XLibur supports.");

        return parsed;
    }
}
