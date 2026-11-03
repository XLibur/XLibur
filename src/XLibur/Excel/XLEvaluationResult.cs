using System;

namespace XLibur.Excel;

/// <summary>
/// The outcome of a <c>TryEvaluate</c> call: the value, or the reason there is none.
/// </summary>
/// <remarks>
/// An Excel error value such as <c>#DIV/0!</c> is a successful result, as it is for
/// <see cref="IXLCell.Value"/>. Only a refused formula, a circular reference and an unsupported
/// feature are failures. A defect in XLibur is thrown, not returned.
/// </remarks>
public readonly struct XLEvaluationResult
{
    private readonly XLCellValue _value;
    private readonly XLFormulaIssue? _issue;

    internal XLEvaluationResult(XLCellValue value)
    {
        _value = value;
        _issue = null;
    }

    internal XLEvaluationResult(XLFormulaIssue issue)
    {
        _value = default;
        _issue = issue;
    }

    /// <summary>Whether the evaluation produced a value.</summary>
    public bool IsSuccess => _issue is null;

    /// <summary>
    /// The value of the evaluation.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// <see cref="IsSuccess"/> is <c>false</c>. A failed evaluation has no value, and a blank would be
    /// taken for the value of an empty cell. Read <see cref="Issue"/> instead.
    /// </exception>
    public XLCellValue Value => _issue is null
        ? _value
        : throw new InvalidOperationException($"The evaluation failed, so it has no value: {_issue.Message}");

    /// <summary>Why the evaluation failed, or <c>null</c> when <see cref="IsSuccess"/> is <c>true</c>.</summary>
    public XLFormulaIssue? Issue => _issue;
}
