using System.ComponentModel;
using System.Text;
using XLibur.Excel;

namespace XLibur.Operations;

public sealed record InspectRequest
{
    [Argument(0), Description("The .xlsx or .xlsm file.")]
    public required string Path { get; init; }
}

public sealed record InspectResult(string Path, IReadOnlyList<SheetSummary> Sheets, IReadOnlyList<NameSummary> DefinedNames)
    : IOperationResult
{
    public void WriteText(StringBuilder text)
    {
        text.Append(Path).Append('\n');
        foreach (var sheet in Sheets)
            sheet.WriteText(text);

        foreach (var name in DefinedNames)
            text.Append("name ").Append(name.Sheet is null ? name.Name : $"{name.Sheet}!{name.Name}").Append(" = ").Append(name.RefersTo).Append('\n');
    }
}

public sealed record SheetSummary(
    string Name,
    int Position,
    string Visibility,
    string? UsedRange,
    int Cells,
    int Formulas,
    IReadOnlyList<TableSummary> Tables,
    IReadOnlyList<PivotSummary> Pivots,
    IReadOnlyList<ChartSummary> Charts,
    int Pictures)
{
    internal void WriteText(StringBuilder text)
    {
        text.Append(Position).Append(' ').Append(Name);
        if (Visibility != "visible")
            text.Append(" (").Append(Visibility).Append(')');
        text.Append(" used=").Append(UsedRange ?? "empty")
            .Append(" cells=").Append(Cells)
            .Append(" formulas=").Append(Formulas);
        AppendList(text, "tables", Tables.Select(t => $"{t.Name}({t.Range})"));
        AppendList(text, "pivots", Pivots.Select(p => $"{p.Name}@{p.TargetCell}"));
        AppendList(text, "charts", Charts.Select(c => c.Title is null ? c.Type : $"{c.Type}:{c.Title}"));
        if (Pictures > 0)
            text.Append(" pictures=").Append(Pictures);
        text.Append('\n');
    }

    private static void AppendList(StringBuilder text, string label, IEnumerable<string> items)
    {
        var list = items.ToList();
        if (list.Count > 0)
            text.Append(' ').Append(label).Append('=').AppendJoin(',', list);
    }
}

public sealed record TableSummary(string Name, string Range);

public sealed record PivotSummary(string Name, string TargetCell);

public sealed record ChartSummary(string Type, string? Title, int Series);

/// <param name="Sheet">The sheet the name is scoped to, or null for a workbook-scoped name.</param>
public sealed record NameSummary(string Name, string? Sheet, string RefersTo);

public sealed class InspectOperation : Operation<InspectRequest, InspectResult>
{
    public override string Name => "inspect";

    public override string Description =>
        "Summarise a workbook: its sheets with used ranges, cell and formula counts, tables, pivot tables, charts and pictures, and its defined names. Start here.";

    public override OperationEffect Effect => OperationEffect.ReadOnly;

    public override InspectResult Execute(InspectRequest request, OperationContext context)
    {
        using var loaded = context.LoadWorkbook(request.Path);
        var workbook = loaded.Workbook;

        var sheets = workbook.Worksheets.Select(ws =>
        {
            var cells = 0;
            var formulas = 0;
            foreach (var cell in ws.CellsUsed(XLCellsUsedOptions.Contents))
            {
                cells++;
                if (cell.HasFormula)
                    formulas++;
            }

            return new SheetSummary(
                ws.Name,
                ws.Position,
                ws.Visibility switch
                {
                    XLWorksheetVisibility.Hidden => "hidden",
                    XLWorksheetVisibility.VeryHidden => "veryHidden",
                    _ => "visible",
                },
                ws.RangeUsed(XLCellsUsedOptions.Contents) is { } used ? Rendering.Area(used.RangeAddress) : null,
                cells,
                formulas,
                ws.Tables.Select(t => new TableSummary(t.Name, t.RangeAddress.ToStringRelative())).ToList(),
                ws.PivotTables.Select(p => new PivotSummary(p.Name, p.TargetCell.Address.ToStringRelative())).ToList(),
                ws.Charts.Select(c => new ChartSummary(c.ChartType.ToString(), c.Title, c.Series.Count)).ToList(),
                ws.Pictures.Count);
        }).ToList();

        var names = workbook.DefinedNames.Select(n => new NameSummary(n.Name, null, n.RefersTo))
            .Concat(workbook.Worksheets.SelectMany(ws => ws.DefinedNames.Select(n => new NameSummary(n.Name, ws.Name, n.RefersTo))))
            .ToList();

        return new InspectResult(loaded.FullPath, sheets, names);
    }
}
