using System.ComponentModel;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using XLibur.Excel;
using XLibur.Operations.Workbooks;

namespace XLibur.Operations;

public sealed record CheckFormulaRequest
{
    [Argument(0), Description("The formula, with or without the leading '='.")]
    public required string Formula { get; init; }

    [Argument(1), Description("A workbook to evaluate against. Without one, the formula can use no references.")]
    public string? Path { get; init; }

    [Description("The sheet to evaluate on. Defaults to the first sheet.")]
    public string? Sheet { get; init; }

    [Description("The cell the formula is evaluated as if it were in, for relative references and ROW(). Defaults to A1.")]
    public string? Cell { get; init; }

    [Description("Evaluate as well as parse. Default true.")]
    public bool Evaluate { get; init; } = true;
}

/// <param name="Position">The character the parser stopped at, counted from 0 in the formula without its '=', when it says.</param>
/// <param name="Failure">Why evaluation produced no value, when it did not.</param>
public sealed record CheckFormulaResult(
    string Formula,
    bool Valid,
    string? Refusal,
    int? Position,
    XLCellValue? Value,
    CellFailure? Failure,
    string? Detail) : IOperationResult
{
    public int ExitCode => Valid && Failure is null ? 0 : 1;

    public void WriteText(StringBuilder text)
    {
        if (!Valid)
            text.Append("refused: ").Append(Refusal).Append('\n');
        else if (Failure is { } failure)
            text.Append("parses; evaluation ").Append(Rendering.Name(failure)).Append(": ").Append(Detail).Append('\n');
        else if (Value is { } value)
            text.Append(Rendering.ToText(value)).Append('\n');
        else
            text.Append("parses\n");
    }
}

public sealed partial class CheckFormulaOperation : Operation<CheckFormulaRequest, CheckFormulaResult>
{
    public override string Name => "check_formula";

    public override string Description =>
        "Parse a formula and, optionally, evaluate it against a workbook without writing it anywhere. Returns why XLibur refuses it, or its computed value.";

    public override OperationEffect Effect => OperationEffect.ReadOnly;

    public override CheckFormulaResult Execute(CheckFormulaRequest request, OperationContext context)
    {
        var text = FormulaCells.WithoutEquals(request.Formula.Trim());
        if (text.Length == 0)
            throw OperationException.BadInput("formula_empty", "A formula is required.");

        if (FormulaCells.Refusal(text) is { } refusal)
        {
            var position = ParserPosition().Match(refusal);
            return new CheckFormulaResult(text, false, refusal,
                position.Success ? int.Parse(position.Groups[1].Value, CultureInfo.InvariantCulture) : null, null, null, null);
        }

        if (!request.Evaluate)
            return new CheckFormulaResult(text, true, null, null, null, null, null);

        CellRead read;
        if (request.Path is null)
        {
            if (request.Sheet is not null || request.Cell is not null)
                throw OperationException.BadInput("path_required", "sheet and cell need a workbook path.");
            read = CellReader.Read(text, static t => XLWorkbook.EvaluateExpr(t));
        }
        else
        {
            using var loaded = context.LoadWorkbook(request.Path);
            var target = Addressing.Cell(loaded.Workbook, request.Sheet, request.Cell ?? "A1");
            var state = (Sheet: target.Worksheet, Text: text, Cell: target.Address.ToStringRelative());
            read = CellReader.Read(state, static s => s.Sheet.Evaluate(s.Text, s.Cell));
        }

        return read.Failure is { } failure
            ? new CheckFormulaResult(text, true, null, null, null, failure, read.Message)
            : new CheckFormulaResult(text, true, null, null, read.Value, null, null);
    }

    [GeneratedRegex(@"\bat char (\d+)\b")]
    private static partial Regex ParserPosition();
}
