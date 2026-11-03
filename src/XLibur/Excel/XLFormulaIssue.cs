using System;

namespace XLibur.Excel;

/// <summary>
/// Describes why a formula was refused, or could not be evaluated, as a value instead of an exception.
/// </summary>
/// <remarks>
/// An Excel error value such as <c>#DIV/0!</c> is a result, not an issue.
/// </remarks>
public sealed class XLFormulaIssue
{
    internal XLFormulaIssue(XLFormulaIssueKind kind, string message, Exception? inner, int? position = null)
    {
        Kind = kind;
        Message = message;
        Inner = inner;
        Position = position;
    }

    /// <summary>Which of the expected failures this is.</summary>
    public XLFormulaIssueKind Kind { get; }

    /// <summary>A description of the issue, for a person to read.</summary>
    public string Message { get; }

    /// <summary>
    /// The zero-based index of the character in the formula text where the parser stopped. From
    /// <c>XLHelper.TryParseFormula</c> it indexes the text exactly as the caller passed it (a leading
    /// <c>=</c> counts). From <c>TryEvaluate</c> it indexes the cell's formula text, which has no
    /// leading <c>=</c>. It is <c>null</c> for an issue that has no place in the text: a circular
    /// reference, an unsupported feature, and a part of the text that the parser read but the calc
    /// engine cannot take, such as a function with the wrong number of arguments. When the parser
    /// refuses text without a more precise place, it is <c>0</c>. Do not parse <see cref="Message"/>
    /// for a position: its wording is not a contract.
    /// </summary>
    public int? Position { get; }

    /// <summary>
    /// The exception that was raised for the issue, for diagnostics. It is not part of the contract:
    /// its type and message can change.
    /// </summary>
    public Exception? Inner { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Kind}: {Message}";
}
