using System;
using System.Collections.Generic;
using System.Globalization;
using XLibur.Excel;

namespace XLibur.Report.Tags;

/// <summary>
/// Totals a column of a generated range. Written in the options row under the column to total —
/// <c>&lt;&lt;Sum&gt;&gt;</c>, <c>&lt;&lt;Avg&gt;&gt;</c>, <c>&lt;&lt;Count&gt;&gt;</c> and the rest.
/// </summary>
/// <remarks>
/// The tag leaves a <c>SUBTOTAL</c> formula rather than a computed number, for two reasons: the
/// total stays live if someone edits a generated row, and <c>SUBTOTAL</c> ignores rows that are
/// filtered out or that are themselves subtotals, so a grand total does not double-count its own
/// group totals.
/// <para>
/// <c>over</c> totals a different column: <c>&lt;&lt;Sum over=D&gt;&gt;</c>.
/// </para>
/// </remarks>
public sealed class SummaryFunctionTag : OptionTag
{
    /// <summary>
    /// Excel's SUBTOTAL function numbers, which is how SUBTOTAL selects the aggregate to apply.
    /// </summary>
    private static readonly Dictionary<string, int> FunctionNumbers = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Avg"] = 1,
        ["Average"] = 1,
        ["Count"] = 2,
        ["CountA"] = 3,
        ["Max"] = 4,
        ["Min"] = 5,
        ["Product"] = 6,
        ["StdDev"] = 7,
        ["StdDevP"] = 8,
        ["Sum"] = 9,
        ["Var"] = 10,
        ["VarP"] = 11,
    };

    /// <inheritdoc />
    public override void Execute(ProcessingContext context)
    {
        if (context.OptionsRow is null)
        {
            return;
        }

        if (!FunctionNumbers.TryGetValue(Token.Name, out var functionNumber))
        {
            context.Errors.Add(new TemplateError($"<<{Token.Name}>> is not a summary function.", context.Worksheet.Name));
            return;
        }

        // The tag's own column decides where the total appears; `over` only changes what is
        // totalled, so a label column can carry the total of a value column.
        var totalled = ResolveColumn(context);
        if (totalled is null)
        {
            return;
        }

        var optionsRowNumber = context.OptionsRow.RangeAddress.FirstAddress.RowNumber;
        var target = context.Worksheet.Cell(optionsRowNumber, Column);

        var generated = context.GeneratedRange.RangeAddress;
        var firstRow = generated.FirstAddress.RowNumber;
        var lastRow = generated.LastAddress.RowNumber;

        if (lastRow < firstRow)
        {
            // Nothing was generated; a total over no rows is zero, not a broken reference.
            target.Value = 0;
            return;
        }

        var columnLetter = XLHelper.GetColumnLetterFromNumber(totalled.Value);

        target.FormulaA1 = string.Create(
            CultureInfo.InvariantCulture,
            $"SUBTOTAL({functionNumber},{columnLetter}{firstRow}:{columnLetter}{lastRow})");
    }

    private int? ResolveColumn(ProcessingContext context)
    {
        var over = Token.Value("over");
        if (over.Length == 0)
        {
            return Column;
        }

        if (int.TryParse(over, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number))
        {
            return number;
        }

        try
        {
            return XLHelper.GetColumnNumberFromLetter(over);
        }
        catch (ArgumentException)
        {
            context.Errors.Add(new TemplateError(
                $"<<{Token.Name} over={over}>> does not name a column.",
                context.Worksheet.Name));
            return null;
        }
    }
}
