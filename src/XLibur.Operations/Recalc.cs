using System.ComponentModel;
using System.Text;
using XLibur.Excel;
using XLibur.Operations.Workbooks;

namespace XLibur.Operations;

public sealed record RecalcRequest
{
    [Argument(0), Description("The .xlsx or .xlsm file.")]
    public required string Path { get; init; }

    [Description("The most cells to list in each category; 0 lists only the counts. Default 50.")]
    public int MaxIssues { get; init; } = 50;

    [Description("A recalculated number equals the cached one when they differ by at most this times the larger magnitude (or times 1 below 1). Default 1e-12.")]
    public double Tolerance { get; init; } = 1e-12;
}

/// <summary>
/// A full recalculation, compared with the values the file cached. <see cref="Changed"/> is where
/// XLibur's result differs from what the program that saved the file computed.
/// </summary>
public sealed record RecalcResult(
    string Path,
    int Formulas,
    IssueList ErrorValues,
    IssueList Circular,
    IssueList Refused,
    IssueList Unsupported,
    IssueList Defects,
    ChangeList Changed) : IOperationResult
{
    public void WriteText(StringBuilder text)
    {
        text.Append(Path).Append(": ").Append(Formulas).Append(" formulas recalculated; ")
            .Append(ErrorValues.Total).Append(" error values, ")
            .Append(Circular.Total).Append(" circular, ")
            .Append(Refused.Total).Append(" refused, ")
            .Append(Unsupported.Total).Append(" unsupported, ")
            .Append(Defects.Total).Append(" defects, ")
            .Append(Changed.Total).Append(" differ from cached\n");
        ErrorValues.WriteText(text, IssueKind.ErrorValue);
        Circular.WriteText(text, IssueKind.Circular);
        Refused.WriteText(text, IssueKind.Refused);
        Unsupported.WriteText(text, IssueKind.Unsupported);
        Defects.WriteText(text, IssueKind.Defect);
        Changed.WriteText(text);
    }
}

public sealed record IssueList(int Total, IReadOnlyList<CellIssue> Cells)
{
    internal void WriteText(StringBuilder text, IssueKind kind)
    {
        foreach (var issue in Cells)
        {
            text.Append(Rendering.Name(issue.Kind)).Append(' ').Append(issue.Cell).Append('\t').Append(issue.Value);
            if (issue.Formula is not null)
                text.Append("\t=").Append(Rendering.Escape(issue.Formula));
            if (issue.Detail is not null && issue.Kind != IssueKind.Circular)
                text.Append('\t').Append(Rendering.Escape(issue.Detail));
            text.Append('\n');
        }

        if (Total > Cells.Count)
            text.Append(Total - Cells.Count).Append(" more ").Append(Rendering.Name(kind)).Append('\n');
    }

    internal static IssueList From(IEnumerable<CellIssue> issues, int max)
    {
        var all = issues.ToList();
        return new IssueList(all.Count, all.Take(max).ToList());
    }
}

public sealed record ChangeList(int Total, IReadOnlyList<ValueChange> Cells)
{
    internal void WriteText(StringBuilder text)
    {
        foreach (var change in Cells)
            text.Append("differs ").Append(change.Cell).Append("\tcached ").Append(change.Before).Append("\tnow ").Append(change.After).Append('\n');
        if (Total > Cells.Count)
            text.Append(Total - Cells.Count).Append(" more differ\n");
    }
}

public sealed class RecalcOperation : Operation<RecalcRequest, RecalcResult>
{
    public override string Name => "recalc";

    public override string Description =>
        "Recalculate every formula with XLibur's engine and report error values, circular references, refused formulas, unsupported features, and cells whose value differs from the one cached in the file. Writes nothing.";

    public override OperationEffect Effect => OperationEffect.ReadOnly;

    public override RecalcResult Execute(RecalcRequest request, OperationContext context)
    {
        using var loaded = context.LoadWorkbook(request.Path);
        return Run(loaded.FullPath, loaded.Workbook, request.MaxIssues, request.Tolerance);
    }

    internal static RecalcResult Run(string path, XLWorkbook workbook, int maxIssues, double tolerance)
    {
        if (maxIssues < 0)
            throw OperationException.BadInput("bad_max_issues", "maxIssues cannot be negative.");

        var cells = FormulaCells.All(workbook).ToList();
        var cached = cells.Select(c => new CellRead(c.CachedValue, null, null)).ToList();
        workbook.RecalculateAllFormulas();

        var issues = new List<CellIssue>();
        var changes = new List<ValueChange>();
        for (var i = 0; i < cells.Count; i++)
        {
            var now = CellReader.Read(cells[i]);
            if (FormulaCells.IssueOf(cells[i], now) is { } issue)
                issues.Add(issue);
            if (now.Failure is null && !FormulaCells.Same(cached[i], now, tolerance))
                changes.Add(new ValueChange(Rendering.Address(cells[i]), cached[i].ToText(), now.ToText()));
        }

        IssueList Of(IssueKind kind) => IssueList.From(issues.Where(x => x.Kind == kind), maxIssues);
        return new RecalcResult(path, cells.Count, Of(IssueKind.ErrorValue), Of(IssueKind.Circular), Of(IssueKind.Refused),
            Of(IssueKind.Unsupported), Of(IssueKind.Defect),
            new ChangeList(changes.Count, changes.Take(maxIssues).ToList()));
    }
}
