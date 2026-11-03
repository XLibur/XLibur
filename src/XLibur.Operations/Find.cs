using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using XLibur.Excel;
using XLibur.Operations.Workbooks;

namespace XLibur.Operations;

public sealed record FindRequest
{
    [Argument(0), Description("The .xlsx or .xlsm file.")]
    public required string Path { get; init; }

    [Argument(1), Description("Text to find in cell values (a substring match).")]
    public string? Text { get; init; }

    [Description("A .NET regular expression to match against cell values instead of text.")]
    public string? Regex { get; init; }

    [Description("Find cells whose value is an error value (#DIV/0!, #N/A, ...) or that fail to calculate.")]
    public bool Errors { get; init; }

    [Description("Match case. Off by default.")]
    public bool MatchCase { get; init; }

    [Description("Also match formula text, not only values.")]
    public bool InFormulas { get; init; }

    [Description("Search one sheet only. Defaults to every sheet.")]
    public string? Sheet { get; init; }

    [Description("The most matches to return. Default 100.")]
    public int MaxResults { get; init; } = 100;
}

public sealed record FindMatch(string Cell, string Value, string? Formula);

public sealed record FindResult(IReadOnlyList<FindMatch> Matches, int Total) : IOperationResult
{
    public void WriteText(StringBuilder text)
    {
        foreach (var match in Matches)
        {
            text.Append(match.Cell).Append('\t').Append(match.Value);
            if (match.Formula is not null)
                text.Append("\t=").Append(Rendering.Escape(match.Formula));
            text.Append('\n');
        }

        text.Append(Total).Append(Total == 1 ? " match" : " matches");
        if (Total > Matches.Count)
            text.Append(", first ").Append(Matches.Count).Append(" shown");
        text.Append('\n');
    }
}

public sealed class FindOperation : Operation<FindRequest, FindResult>
{
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(1);
    private static readonly TimeSpan SearchBudget = TimeSpan.FromSeconds(10);

    public override string Name => "find";

    public override string Description =>
        "Find cells by text, by regular expression, or by error value, across every sheet or one. Returns each match's address, value and formula.";

    public override OperationEffect Effect => OperationEffect.ReadOnly;

    public override FindResult Execute(FindRequest request, OperationContext context)
    {
        var criteria = (request.Text is not null ? 1 : 0) + (request.Regex is not null ? 1 : 0) + (request.Errors ? 1 : 0);
        if (criteria != 1)
            throw OperationException.BadInput("bad_criteria", "Give exactly one of text, regex or errors.");
        if (request.MaxResults < 1)
            throw OperationException.BadInput("bad_max_results", "maxResults must be at least 1.");

        var isHit = HitTest(request);
        using var loaded = context.LoadWorkbook(request.Path);
        var sheets = request.Sheet is null
            ? loaded.Workbook.Worksheets.ToList()
            : [Addressing.Sheet(loaded.Workbook, request.Sheet)];

        var matches = new List<FindMatch>();
        var total = 0;
        foreach (var cell in sheets.SelectMany(s => s.CellsUsed(XLCellsUsedOptions.Contents)))
        {
            var read = CellReader.Read(cell);
            var formula = cell.HasFormula ? cell.FormulaA1 : null;
            if (!isHit(read, formula))
                continue;

            total++;
            if (matches.Count < request.MaxResults)
                matches.Add(new FindMatch(Rendering.Address(cell), read.ToText(), formula));
        }

        return new FindResult(matches, total);
    }

    private static Func<CellRead, string?, bool> HitTest(FindRequest request)
    {
        if (request.Errors)
            return (read, _) => read.Failure is not null || read.Value.IsError;

        var matches = Matcher(request);
        return (read, formula) => matches(read.Value.ToString(CultureInfo.InvariantCulture))
                                  || request.InFormulas && formula is not null && matches(formula);
    }

    private static Func<string, bool> Matcher(FindRequest request)
    {
        if (request.Text is { } text)
        {
            var comparison = request.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            return value => value.Contains(text, comparison);
        }

        var pattern = request.Regex!;

        Regex regex;
        try
        {
            regex = new Regex(pattern, request.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase, RegexTimeout);
        }
        catch (ArgumentException e)
        {
            throw OperationException.BadInput("bad_regex", $"'{pattern}' is not a valid regular expression: {e.Message}");
        }

        var clock = Stopwatch.StartNew();
        return value =>
        {
            if (clock.Elapsed > SearchBudget)
                throw OperationException.BadInput("regex_timeout", $"The regular expression '{pattern}' took longer than {SearchBudget.TotalSeconds:0} s over the workbook.");
            try
            {
                return regex.IsMatch(value);
            }
            catch (RegexMatchTimeoutException)
            {
                throw OperationException.BadInput("regex_timeout", $"The regular expression '{pattern}' took longer than {RegexTimeout.TotalSeconds:0} s on one cell.");
            }
        };
    }
}
