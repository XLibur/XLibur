using System.ComponentModel;
using System.Globalization;
using System.Text;
using System.Text.Json;
using XLibur.Operations.Workbooks;
using XLibur.Report;

namespace XLibur.Operations;

public sealed record RenderReportRequest
{
    [Argument(0), Description("The .xlsx template, written for XLibur.Report.")]
    public required string Template { get; init; }

    [Argument(1), Description("A .json file holding one object; each property becomes a template variable. Write a date as {\"date\":\"2026-01-31\"}.")]
    public required string Data { get; init; }

    [Argument(2), Description("The .xlsx to generate.")]
    public required string Output { get; init; }

    [Description("Replace the output if it exists.")]
    public bool Overwrite { get; init; }
}

public sealed record RenderReportResult(string Output, IReadOnlyList<TemplateProblem> Errors) : IOperationResult
{
    public int ExitCode => Errors.Count == 0 ? 0 : 1;

    public void WriteText(StringBuilder text)
    {
        text.Append("wrote ").Append(Output).Append(", ").Append(Errors.Count).Append(" template errors\n");
        foreach (var error in Errors)
            text.Append(error.Location.Length == 0 ? "" : error.Location + ": ").Append(error.Message).Append('\n');
    }
}

public sealed record TemplateProblem(string Location, string Message);

public sealed class RenderReportOperation : Operation<RenderReportRequest, RenderReportResult>
{
    public override string Name => "render_report";

    public override string Description =>
        "Generate a report from an XLibur.Report .xlsx template and a JSON data file. Template errors do not stop generation; they are listed, and the cell shows the error.";

    public override OperationEffect Effect => OperationEffect.Writes;

    public override RenderReportResult Execute(RenderReportRequest request, OperationContext context)
    {
        var target = context.Paths.ResolveWrite(request.Output, request.Overwrite);
        if (!OperationContext.IsWorkbookExtension(target.FullPath))
            throw OperationException.BadInput("not_a_workbook", "The report output must be an .xlsx or .xlsm file.");

        var (_, data) = context.ReadFile(request.Data);
        JsonElement root;
        try
        {
            using var document = JsonDocument.Parse(data);
            root = document.RootElement.Clone();
        }
        catch (JsonException e)
        {
            throw OperationException.BadInput("bad_data", $"'{request.Data}' is not JSON: {e.Message}", e);
        }

        if (root.ValueKind != JsonValueKind.Object)
            throw OperationException.BadInput("bad_data", $"'{request.Data}' must hold one JSON object; each property is a variable.");

        using var loaded = context.LoadWorkbook(request.Template);
        using var report = new XLTemplate(loaded.Workbook);
        foreach (var property in root.EnumerateObject())
            report.AddVariable(property.Name, ToClr(property.Value));

        var result = report.Generate();
        context.Paths.WriteAtomically(target, stream => report.SaveAs(stream));
        return new RenderReportResult(target.FullPath,
            result.ParsingErrors.Select(e => new TemplateProblem(e.Location, e.Message)).ToList());
    }

    /// <summary>
    /// JSON as the objects a template reads: an object becomes a dictionary, an array a list, and a
    /// number a decimal when it fits. A date is written as a cell value is, <c>{"date":"2026-01-31"}</c>;
    /// a plain string stays a string, however it looks.
    /// </summary>
    internal static object? ToClr(JsonElement element) => element.ValueKind switch
    {
        JsonValueKind.Object when TaggedDate(element) is { } date => date,
        JsonValueKind.Object => element.EnumerateObject().ToDictionary(p => p.Name, p => ToClr(p.Value)),
        JsonValueKind.Array => element.EnumerateArray().Select(ToClr).ToList(),
        JsonValueKind.Number => element.TryGetDecimal(out var d) ? d : element.GetDouble(),
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.String => element.GetString(),
        _ => null,
    };

    private static DateTime? TaggedDate(JsonElement element)
        => element.EnumerateObject().Count() == 1 && element.TryGetProperty("date", out var date) &&
           date.ValueKind == JsonValueKind.String &&
           DateTime.TryParse(date.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var parsed)
            ? parsed
            : null;
}
