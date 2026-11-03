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
    internal XLFormulaIssue(XLFormulaIssueKind kind, string message, Exception? inner)
    {
        Kind = kind;
        Message = message;
        Inner = inner;
    }

    /// <summary>Which of the expected failures this is.</summary>
    public XLFormulaIssueKind Kind { get; }

    /// <summary>A description of the issue, for a person to read.</summary>
    public string Message { get; }

    /// <summary>
    /// The zero-based index of the character in the formula text, exactly as the caller passed it
    /// (a leading <c>=</c> counts), where the issue was found. It is always <c>null</c> for now,
    /// because the parser does not report a position. It will stay <c>null</c> for an issue that has
    /// no place in the text, such as a circular reference. Do not parse <see cref="Message"/> for a
    /// position: its wording is not a contract.
    /// </summary>
    public int? Position => null;

    /// <summary>
    /// The exception that was raised for the issue, for diagnostics. It is not part of the contract:
    /// its type and message can change.
    /// </summary>
    public Exception? Inner { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Kind}: {Message}";
}
