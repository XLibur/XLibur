using System.ComponentModel;
using System.Globalization;
using System.Text;
using XLibur.Excel;
using XLibur.Operations.Workbooks;

namespace XLibur.Operations;

/// <summary>One change to one cell: a value, a formula, or a clear. Exactly one of the three.</summary>
public sealed record CellEdit
{
    [Description("The cell: 'Sheet1!B2', or 'B2' on the request's sheet.")]
    public required string Cell { get; init; }

    [Description("A value: a number, a string, true/false, or {\"date\":\"2026-01-31\"}.")]
    public XLCellValue? Value { get; init; }

    [Description("A formula, with or without the leading '='.")]
    public string? Formula { get; init; }

    [Description("Clear the cell's contents.")]
    public bool Clear { get; init; }

    /// <summary>
    /// The CLI form: <c>B2=42</c> sets a value, <c>B2==SUM(A1:A9)</c> a formula, <c>B2=</c> clears.
    /// A value that reads as an invariant number is a number, TRUE or FALSE a boolean, anything else text.
    /// </summary>
    public static CellEdit Parse(string text)
    {
        var equals = text.IndexOf('=');
        if (equals <= 0)
            throw OperationException.BadInput("bad_edit", $"'{text}' is not CELL=VALUE, CELL==FORMULA or CELL=.");

        var cell = text[..equals];
        var rest = text[(equals + 1)..];
        if (rest.Length == 0)
            return new CellEdit { Cell = cell, Clear = true };
        if (rest[0] == '=')
            return new CellEdit { Cell = cell, Formula = rest[1..] };
        if (double.TryParse(rest, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            return new CellEdit { Cell = cell, Value = number };
        if (bool.TryParse(rest, out var logical))
            return new CellEdit { Cell = cell, Value = logical };
        return new CellEdit { Cell = cell, Value = rest };
    }
}

public sealed record EditRequest
{
    [Argument(0), Description("The .xlsx or .xlsm file to edit.")]
    public required string Path { get; init; }

    [Description("The changes, applied together: if any is invalid, none is applied and nothing is written.")]
    public required IReadOnlyList<CellEdit> Edits { get; init; }

    [Description("The sheet an unqualified cell is on. Defaults to the first sheet.")]
    public string? Sheet { get; init; }

    [Description("Write the result here instead of over the source, leaving the source untouched.")]
    public string? OutputPath { get; init; }

    [Description("Replace outputPath if it exists. Without outputPath the edit always saves over the source.")]
    public bool Overwrite { get; init; }

    [Description("The most dependent value changes to list. Default 50.")]
    public int MaxChanges { get; init; } = 50;
}

/// <param name="Edited">Every edited cell with its value after recalculation.</param>
/// <param name="NewIssues">
/// Formula cells that now hold an error value or fail to calculate and did not before the edit, and
/// edited cells that do, whatever they held before.
/// </param>
/// <param name="Changed">Other formula cells whose value the edit changed, up to maxChanges.</param>
public sealed record EditResult(
    string Saved,
    IReadOnlyList<EditedCell> Edited,
    IReadOnlyList<CellIssue> NewIssues,
    IReadOnlyList<ValueChange> Changed,
    int ChangedTotal) : IOperationResult
{
    public void WriteText(StringBuilder text)
    {
        text.Append("saved ").Append(Saved).Append('\n');
        foreach (var cell in Edited)
        {
            text.Append(cell.Cell).Append('\t').Append(cell.Value);
            if (cell.Formula is not null)
                text.Append("\t=").Append(Rendering.Escape(cell.Formula));
            text.Append('\n');
        }

        foreach (var issue in NewIssues)
            text.Append("new ").Append(Rendering.Name(issue.Kind)).Append(' ').Append(issue.Cell).Append('\t').Append(issue.Value).Append('\n');
        foreach (var change in Changed)
            text.Append("changed ").Append(change.Cell).Append('\t').Append(change.Before).Append(" -> ").Append(change.After).Append('\n');
        if (ChangedTotal > Changed.Count)
            text.Append(ChangedTotal - Changed.Count).Append(" more changed\n");
    }
}

public sealed record EditedCell(string Cell, string Value, string? Formula);

public sealed class EditOperation : Operation<EditRequest, EditResult>
{
    public override string Name => "edit";

    public override string Description =>
        "Set values or formulas, or clear cells, all at once; recalculate; save atomically (over the source, or to outputPath). Returns each edited cell's computed value, every formula that turned into an error, and the dependent values that changed.";

    public override OperationEffect Effect => OperationEffect.Writes;

    public override EditResult Execute(EditRequest request, OperationContext context)
    {
        if (request.Edits.Count == 0)
            throw OperationException.BadInput("no_edits", "Give at least one edit.");
        foreach (var edit in request.Edits)
            Check(edit);

        var target = context.Paths.ResolveWrite(request.OutputPath ?? request.Path, overwrite: request.OutputPath is null || request.Overwrite);
        using var loaded = context.LoadWorkbook(request.Path);
        if (!string.Equals(Path.GetExtension(target.FullPath), Path.GetExtension(loaded.FullPath), StringComparison.OrdinalIgnoreCase))
            throw OperationException.BadInput("extension_mismatch", "outputPath must have the same extension as the source (.xlsx or .xlsm).");

        var workbook = loaded.Workbook;
        var targets = request.Edits.Select(e => (Edit: e, Cell: Addressing.Cell(workbook, request.Sheet, e.Cell))).ToList();

        var before = FormulaCells.ReadAll(workbook);
        foreach (var (edit, cell) in targets)
            Apply(edit, cell);

        var edited = targets
            .Select(t => new EditedCell(Rendering.Address(t.Cell), CellReader.Read(t.Cell).ToText(), t.Cell.HasFormula ? t.Cell.FormulaA1 : null))
            .ToList();
        var editedKeys = targets.Select(t => FormulaCells.Key(t.Cell)).ToHashSet();

        var issues = new List<CellIssue>();
        var changed = new List<ValueChange>();
        foreach (var cell in FormulaCells.All(workbook))
        {
            var key = FormulaCells.Key(cell);
            var after = CellReader.Read(cell);
            var had = before.TryGetValue(key, out var old);
            var isEdited = editedKeys.Contains(key);
            if (FormulaCells.IsProblem(after) && (!had || !FormulaCells.IsProblem(old) || isEdited))
                issues.Add(FormulaCells.IssueOf(cell, after)!);
            else if (had && !isEdited && !FormulaCells.Same(old, after, 0))
                changed.Add(new ValueChange(Rendering.Address(cell), old.ToText(), after.ToText()));
        }

        context.SaveWorkbook(workbook, target);
        return new EditResult(target.FullPath, edited, issues, changed.Take(request.MaxChanges).ToList(), changed.Count);
    }

    private static void Check(CellEdit edit)
    {
        var parts = (edit.Value is not null ? 1 : 0) + (edit.Formula is not null ? 1 : 0) + (edit.Clear ? 1 : 0);
        if (parts != 1)
            throw OperationException.BadInput("bad_edit", $"The edit of '{edit.Cell}' must have exactly one of value, formula or clear.");
        if (edit.Formula is { } formula && FormulaCells.Refusal(FormulaCells.WithoutEquals(formula)) is { } refusal)
            throw OperationException.BadInput("formula_refused", $"XLibur cannot read the formula for '{edit.Cell}': {refusal}");
    }

    private static void Apply(CellEdit edit, IXLCell cell)
    {
        if (edit.Clear)
            cell.Clear(XLClearOptions.Contents);
        else if (edit.Formula is { } formula)
            cell.FormulaA1 = FormulaCells.WithoutEquals(formula);
        else
            cell.Value = edit.Value!.Value;
    }
}
