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
            XLError.SpillRange => "#SPILL!",
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
        ["#SPILL!"] = XLError.SpillRange
    };

    public static bool TryParseError(string input, out XLError error)
        => ErrorMap.TryGetValue(input.Trim(), out error);

    /// <summary>
    /// The <see cref="XLError"/> an error literal in a formula stands for. The parser reads every error
    /// value Excel knows, but <see cref="XLError"/> has no member for most of the newer ones
    /// (<c>#CALC!</c>, <c>#FIELD!</c>, ...), so a formula holding one is refused like unreadable text.
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
