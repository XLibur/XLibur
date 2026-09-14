using System;
using System.Collections.Generic;
using XLibur.Excel;
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
    /// A formula that is part of a circular reference has no value to read. It is recorded as a
    /// template error, once per cell, and read as blank, so generation carries on and the cell keeps
    /// its formula. Every other failure still throws.
    /// </remarks>
    public XLCellValue ReadValue(IXLCell cell)
    {
        try
        {
            return cell.Value;
        }
        catch (XLCircularReferenceException ex)
        {
            var sheet = cell.Worksheet.Name;
            var address = cell.Address.ToString();
            if (_unreadable.Add((sheet, address)))
            {
                _errors.Add(new TemplateError(
                    "The formula in this cell is part of a circular reference, so its value cannot be read.",
                    sheet,
                    address,
                    ex));
            }

            return Blank.Value;
        }
    }

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
