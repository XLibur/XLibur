using System.Text.RegularExpressions;
using XLibur.Excel;
using XLibur.Excel.CalcEngine;

namespace XLibur.Operations.Workbooks;

/// <summary>What is wrong with a cell after calculation: an error value, or one of the failures of <see cref="CellFailure"/>.</summary>
public enum IssueKind
{
    ErrorValue,
    Circular,
    Refused,
    Unsupported,
    Defect,
}

public sealed record CellIssue(string Cell, IssueKind Kind, string Value, string? Formula, string? Detail);

public sealed record ValueChange(string Cell, string Before, string After);

internal readonly record struct CellKey(string Sheet, int Row, int Column);

internal static class FormulaCells
{
    internal static IEnumerable<IXLCell> All(XLWorkbook workbook)
        => workbook.Worksheets.SelectMany(ws => ws.CellsUsed(XLCellsUsedOptions.Contents).Where(c => c.HasFormula));

    internal static CellKey Key(IXLCell cell) => new(cell.Worksheet.Name, cell.Address.RowNumber, cell.Address.ColumnNumber);

    internal static Dictionary<CellKey, CellRead> ReadAll(XLWorkbook workbook)
        => All(workbook).ToDictionary(Key, CellReader.Read);

    internal static CellIssue? IssueOf(IXLCell cell, CellRead read)
    {
        if (read.Failure is { } failure)
        {
            var kind = failure switch
            {
                CellFailure.Circular => IssueKind.Circular,
                CellFailure.Refused => IssueKind.Refused,
                CellFailure.Unsupported => IssueKind.Unsupported,
                _ => IssueKind.Defect,
            };
            return new CellIssue(Rendering.Address(cell), kind, read.ToText(), cell.FormulaA1, read.Message);
        }

        if (!read.Value.IsError)
            return null;

        var formula = cell.HasFormula ? cell.FormulaA1 : null;
        if (formula is not null && read.Value.GetError() == XLError.NameNotRecognized &&
            UnimplementedFunctions(formula, cell.Worksheet) is { Count: > 0 } missing)
            return new CellIssue(Rendering.Address(cell), IssueKind.Unsupported, read.ToText(), formula,
                $"XLibur does not implement {string.Join(", ", missing)}, so the formula evaluates to #NAME?.");
        return new CellIssue(Rendering.Address(cell), IssueKind.ErrorValue, read.ToText(), formula, null);
    }

    private static readonly Lazy<HashSet<string>> Implemented =
        new(() => new HashSet<string>(new XLFunctionLibrary().Names, StringComparer.OrdinalIgnoreCase));

    private static readonly Regex FunctionCall = new(@"(?<![A-Za-z0-9_.\]!])([A-Za-z_][A-Za-z0-9_.]*)\s*\(", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));

    /// <summary>
    /// The functions a formula calls that XLibur's calc engine does not have. XLibur evaluates such a
    /// call to <c>#NAME?</c>, as Excel does for a name it does not know, which would hide an
    /// unsupported feature behind an error value. A defined name, such as a LAMBDA's, is not counted.
    /// String literals are skipped, so a function name inside quotes is not counted either.
    /// </summary>
    internal static List<string> UnimplementedFunctions(string formula, IXLWorksheet sheet)
    {
        var code = Regex.Replace(formula, "\"(?:[^\"]|\"\")*\"", "\"\"", RegexOptions.None, TimeSpan.FromSeconds(1));
        return FunctionCall.Matches(code)
            .Select(m => m.Groups[1].Value)
            .Select(name => name.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase) ? name[6..] : name)
            .Select(name => name.StartsWith("_xlws.", StringComparison.OrdinalIgnoreCase) ? name[6..] : name)
            .Where(name => !Implemented.Value.Contains(name) && !sheet.Workbook.DefinedNames.Contains(name) && !sheet.DefinedNames.Contains(name)
                           && !name.StartsWith("_xlpm.", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(name => name.ToUpperInvariant())
            .ToList();
    }

    internal static bool IsProblem(CellRead read) => read.Failure is not null || read.Value.IsError;

    /// <summary>
    /// Whether two reads hold the same result. Two numbers are the same when they differ by at most
    /// the tolerance times the larger magnitude, or times 1 when both are below 1.
    /// </summary>
    internal static bool Same(CellRead a, CellRead b, double tolerance)
    {
        if (a.Failure != b.Failure)
            return false;
        if (a.Failure is not null)
            return true;
        return SameValue(a.Value, b.Value, tolerance);
    }

    internal static bool SameValue(XLCellValue a, XLCellValue b, double tolerance)
    {
        if (a.IsUnifiedNumber && b.IsUnifiedNumber && tolerance > 0)
        {
            var x = a.GetUnifiedNumber();
            var y = b.GetUnifiedNumber();
            return Math.Abs(x - y) <= tolerance * Math.Max(1, Math.Max(Math.Abs(x), Math.Abs(y)));
        }

        return a.Equals(b);
    }

    /// <summary>
    /// Why XLibur refuses formula text, or null when it reads it. XLibur has no public parse call, and
    /// setting a formula does not parse it; reading it back as R1C1 does, and throws on a refusal. The
    /// formula goes into a scratch workbook, so checking it changes nothing in the caller's.
    /// </summary>
    internal static string? Refusal(string formula)
    {
        using var scratch = WorkbookLoader.Create();
        var cell = scratch.AddWorksheet().Cell(1, 1);
        try
        {
            cell.FormulaA1 = formula;
            _ = cell.FormulaR1C1;
            return null;
        }
        catch (ExpressionParseException e)
        {
            return e.InnerException?.Message ?? e.Message;
        }
    }

    internal static string WithoutEquals(string formula) => formula.StartsWith('=') ? formula[1..] : formula;
}
