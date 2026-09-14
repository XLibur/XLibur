using System;
using System.Collections.Generic;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;
using XLibur.Excel.CalcEngine.Exceptions;
using XLibur.Report.Excel;
using XLibur.Report.Expressions;

namespace XLibur.Report.Ranges;

/// <summary>
/// Resolves the expressions in a single cell, recording rather than throwing on failure.
/// </summary>
internal sealed class CellEvaluator
{
    private readonly IExpressionEngine _engine;
    private readonly TemplateErrors _errors;

    /// <summary>The cells already reported as unreadable, so each is reported once.</summary>
    private readonly HashSet<(string Sheet, string? Address)> _unreadable = new();

    public CellEvaluator(IExpressionEngine engine, TemplateErrors errors)
    {
        _engine = engine;
        _errors = errors;
    }

    /// <summary>
    /// Reads <paramref name="cell"/>'s value for the expander, which looks through every cell of a
    /// bound range for tags and expressions.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A formula that is part of, or depends on, a circular reference, a feature XLibur does not
    /// evaluate, or a formula the parser cannot read has no value to read. These are the expected
    /// failures (spec 56, Q22; #488). Each is recorded as a template error, once per cell, and read
    /// as blank, so generation carries on and the cell keeps its formula. A cycle elsewhere in the
    /// workbook does not reach the read (#492). A defect still throws (#459).
    /// </para>
    /// <para>
    /// Which failures are expected is decided by spec 56's policy table, which is internal to
    /// XLibur. <see cref="IXLCell.TryGetValue{T}"/> reads its row for a tolerant read: it answers
    /// <c>false</c> for an expected failure and lets a defect throw. Only then is the formula read
    /// again, to learn which kind of failure it was.
    /// </para>
    /// </remarks>
    public XLCellValue ReadValue(IXLCell cell)
    {
        // Any value converts to text, so false means the formula has no value. A cell without a
        // formula cannot fail, and skips the conversion.
        if (!cell.HasFormula || cell.TryGetValue(out string _))
        {
            // Calculated by the line above, so this reads the value it left.
            return cell.Value;
        }

        try
        {
            return cell.Value;
        }
        catch (Exception ex) when (Describe(ex) is { } problem)
        {
            var sheet = cell.Worksheet.Name;
            var address = cell.Address.ToString();
            if (_unreadable.Add((sheet, address)))
            {
                _errors.Add(new TemplateError(problem, sheet, address, ex));
            }

            return Blank.Value;
        }
    }

    /// <summary>
    /// What went wrong, for each kind of failure the policy table expects, or <c>null</c> for any
    /// other failure, which then reaches the caller.
    /// </summary>
    /// <remarks>
    /// A plain <see cref="NotImplementedException"/> is a defect, not an unsupported feature, but it
    /// never gets here: <see cref="ReadValue"/> reads the formula again only when the policy table
    /// has already called the failure expected.
    /// </remarks>
    private static string? Describe(Exception exception) => exception switch
    {
        XLCircularReferenceException =>
            "The formula in this cell is part of, or depends on, a circular reference, so its value cannot be read.",
        NotImplementedException or NotSupportedException =>
            "The formula in this cell uses, or depends on, a feature XLibur does not evaluate, so its value cannot be read.",
        ExpressionParseException =>
            "The formula in this cell is, or depends on, a formula XLibur cannot parse, so its value cannot be read.",
        _ => null,
    };

    /// <summary>
    /// Evaluates <paramref name="cell"/> against <paramref name="scope"/>. Cells holding a
    /// formula, a non-text value, or text with no expressions are left alone.
    /// </summary>
    public void Evaluate(IXLCell cell, ExpressionScope scope)
    {
        if (cell.HasFormula)
        {
            return;
        }

        var value = cell.Value;
        if (!value.IsText)
        {
            return;
        }

        var text = value.GetText();

        if (text.StartsWith(ExpressionText.FormulaPrefix, StringComparison.Ordinal))
        {
            EvaluateFormula(cell, text, scope);
            return;
        }

        if (!ExpressionText.Contains(text))
        {
            return;
        }

        try
        {
            // A cell that is nothing but one expression keeps that expression's type, so a decimal
            // total reaches Excel as a number. Anything mixed with literal text can only be text.
            cell.Value = ExpressionText.TryGetSingleExpression(text, out var expression)
                ? ReportValueConverter.ToCellValue(_engine.Evaluate(expression, scope))
                : _engine.Interpolate(text, scope);
        }
        catch (ExpressionEvaluationException ex)
        {
            Record(cell, ex);
        }
    }

    private void EvaluateFormula(IXLCell cell, string text, ExpressionScope scope)
    {
        var body = text.Substring(ExpressionText.FormulaPrefix.Length);

        try
        {
            cell.FormulaA1 = _engine.Interpolate(body, scope);
        }
        catch (ExpressionEvaluationException ex)
        {
            Record(cell, ex);
        }
    }

    private void Record(IXLCell cell, ExpressionEvaluationException exception)
    {
        _errors.Add(new TemplateError(
            exception.Message,
            cell.Worksheet.Name,
            cell.Address.ToString(),
            exception));

        // Leave the failure visible in the report rather than silently blanking the cell.
        cell.Value = exception.Message;
        cell.Style.Font.FontColor = XLColor.Red;
    }
}
