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

    /// <summary>
    /// The formulas found to have no value, by where they are, each with the formula it held, so
    /// each is evaluated and reported once.
    /// </summary>
    private readonly Dictionary<(string Sheet, int Row, int Column), string> _unreadable = new();

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
    /// again, to learn which kind of failure it was. That read cannot overturn the verdict: a failure
    /// it does not recognise as expected still throws.
    /// </para>
    /// <para>
    /// The expander reads a cell up to three times, and each evaluation of a failing formula can be
    /// a full recalculation, so a formula found to have no value is not evaluated again. That holds
    /// only while the cell keeps the same formula: a range bound to no data has its rows deleted,
    /// which can move another cell into the same place.
    /// </para>
    /// </remarks>
    public XLCellValue ReadValue(IXLCell cell)
    {
        if (!cell.HasFormula)
        {
            // A cell without a formula cannot fail.
            return cell.Value;
        }

        var address = cell.Address;
        (string Sheet, int Row, int Column) key = (cell.Worksheet.Name, address.RowNumber, address.ColumnNumber);
        if (_unreadable.TryGetValue(key, out var formula) && formula == cell.FormulaA1)
        {
            return Blank.Value;
        }

        // Any value converts to text, so false means the formula has no value.
        if (cell.TryGetValue(out string _))
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
            _unreadable[key] = cell.FormulaA1;
            _errors.Add(new TemplateError(problem, key.Sheet, address.ToString(), ex));
            return Blank.Value;
        }
    }

    /// <summary>
    /// What went wrong, for each kind of failure the policy table expects, or <c>null</c> for any
    /// other failure, which then reaches the caller.
    /// </summary>
    private static string? Describe(Exception exception) => exception switch
    {
        XLCircularReferenceException =>
            "The formula in this cell is part of, or depends on, a circular reference, so its value cannot be read.",
        ExpressionParseException =>
            "The formula in this cell is, or depends on, a formula XLibur cannot parse, so its value cannot be read.",
        _ when IsUnsupportedFeature(exception) =>
            "The formula in this cell uses, or depends on, a feature XLibur does not evaluate, so its value cannot be read.",
        _ => null,
    };

    /// <summary>
    /// Whether <paramref name="exception"/> is how the calc engine raises an unsupported feature.
    /// </summary>
    /// <remarks>
    /// The calc engine raises one as its own sealed subclass of <see cref="NotImplementedException"/>,
    /// which is internal to XLibur, so it is recognised as a subclass that XLibur declares. A plain
    /// <see cref="NotImplementedException"/>, or a subclass declared anywhere else, such as in a
    /// function a caller registers, is a defect. <see cref="NotSupportedException"/> counts as
    /// unsupported, as it does in the policy table.
    /// </remarks>
    private static bool IsUnsupportedFeature(Exception exception) => exception switch
    {
        NotSupportedException => true,
        NotImplementedException => exception.GetType() != typeof(NotImplementedException)
            && exception.GetType().Assembly == typeof(IXLCell).Assembly,
        _ => false,
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
