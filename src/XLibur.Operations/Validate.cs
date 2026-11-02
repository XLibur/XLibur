using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using XLibur.Excel;
using XLibur.Operations.Workbooks;

namespace XLibur.Operations;

public sealed record ValidateRequest
{
    [Argument(0), Description("The .xlsx or .xlsm file.")]
    public required string Path { get; init; }

    [Description("Save the workbook to memory, load it again, and compare the two. Default true.")]
    public bool RoundTrip { get; init; } = true;

    [Description("The most cells to list in each category. Default 20.")]
    public int MaxIssues { get; init; } = 20;
}

/// <param name="Ok">Loaded, no refused formula, no defect, and a round trip with no difference.</param>
public sealed record ValidateResult(
    string Path,
    bool Ok,
    bool Loaded,
    string? LoadError,
    long LoadMilliseconds,
    RecalcResult? Recalc,
    RoundTripReport? RoundTrip) : IOperationResult
{
    public int ExitCode => Ok ? 0 : 1;

    public void WriteText(StringBuilder text)
    {
        text.Append(Ok ? "ok " : "FAILED ").Append(Path).Append('\n');
        if (!Loaded)
        {
            text.Append("load failed: ").Append(LoadError).Append('\n');
            return;
        }

        text.Append("loaded in ").Append(LoadMilliseconds).Append(" ms\n");
        Recalc?.WriteText(text);
        if (RoundTrip is { } roundTrip)
        {
            if (roundTrip.Differences is not { } differences)
                text.Append("round trip failed: ").Append(roundTrip.Error).Append('\n');
            else if (differences.Identical)
                text.Append("round trip: identical\n");
            else
            {
                text.Append("round trip: differs\n");
                differences.WriteText(text);
            }
        }
    }
}

public sealed record RoundTripReport(string? Error, DiffResult? Differences);

public sealed class ValidateOperation : Operation<ValidateRequest, ValidateResult>
{
    public override string Name => "validate";

    public override string Description =>
        "Check a workbook: that it loads, how its formulas recalculate (refused formulas, unsupported features, defects), and that a save and reload by XLibur changes nothing. Exit code 1 when a check fails.";

    public override OperationEffect Effect => OperationEffect.ReadOnly;

    public override ValidateResult Execute(ValidateRequest request, OperationContext context)
    {
        var (fullPath, bytes) = context.ReadFile(request.Path);
        if (!OperationContext.IsWorkbookExtension(fullPath))
            throw OperationException.BadInput("not_a_workbook", $"'{request.Path}' is not an .xlsx or .xlsm file.");

        var clock = Stopwatch.StartNew();
        XLWorkbook workbook;
        try
        {
            workbook = WorkbookLoader.Load(bytes, request.Path, context.Limits, recalculate: false);
        }
        catch (OperationException e)
        {
            return new ValidateResult(fullPath, false, false, e.Error.ToText(), clock.ElapsedMilliseconds, null, null);
        }

        var loadMs = clock.ElapsedMilliseconds;
        RecalcResult recalc;
        using (workbook)
            recalc = RecalcOperation.Run(fullPath, workbook, request.MaxIssues, 1e-12);

        var roundTrip = request.RoundTrip ? RoundTrip(fullPath, bytes, context.Limits, request.MaxIssues) : null;
        var ok = recalc.Refused.Total == 0 && recalc.Defects.Total == 0
                 && roundTrip is null or { Error: null, Differences.Identical: true };
        return new ValidateResult(fullPath, ok, true, null, loadMs, recalc, roundTrip);
    }

    private static RoundTripReport RoundTrip(string path, byte[] bytes, Limits limits, int maxIssues)
    {
        using var original = WorkbookLoader.Load(bytes, path, limits, recalculate: false);
        var saved = new MemoryStream();
        try
        {
            original.SaveAs(saved);
        }
        catch (Exception e)
        {
            return new RoundTripReport($"save threw {e.GetType().FullName}: {e.Message}", null);
        }

        XLWorkbook reloaded;
        try
        {
            reloaded = WorkbookLoader.Load(saved.ToArray(), path + " (saved)", limits, recalculate: false);
        }
        catch (OperationException e)
        {
            return new RoundTripReport($"the saved file does not load: {e.Error.Message}", null);
        }

        using (reloaded)
        {
            var diff = WorkbookComparer.Compare(path, original, path + " (saved)", reloaded,
                new ComparerOptions([], [], Styles: true, Tolerance: 0, MaxChanges: maxIssues, Recalculated: false));
            return new RoundTripReport(null, diff);
        }
    }
}
