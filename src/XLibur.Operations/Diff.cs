using System.ComponentModel;
using System.Text;
using XLibur.Excel;
using XLibur.Operations.Workbooks;

namespace XLibur.Operations;

public enum ValueSource
{
    /// <summary>Compare the values the files cached, as Excel saved them.</summary>
    Cached,

    /// <summary>Recalculate both workbooks with XLibur first, then compare.</summary>
    Recalculated,
}

public sealed record DiffRequest
{
    [Argument(0), Description("The first workbook.")]
    public required string Left { get; init; }

    [Argument(1), Description("The second workbook.")]
    public required string Right { get; init; }

    [Description("cached (default) compares the values in the files; recalculated recalculates both first.")]
    public ValueSource Values { get; init; } = ValueSource.Cached;

    [Description("Sheets to leave out of the comparison.")]
    public IReadOnlyList<string> IgnoreSheets { get; init; } = [];

    [Description("Ranges to leave out, such as 'Sheet1!A1:B9'.")]
    public IReadOnlyList<string> IgnoreRanges { get; init; } = [];

    [Description("Compare cell styles too. Default true.")]
    public bool Styles { get; init; } = true;

    [Description("Two numbers are equal when they differ by at most this times the larger magnitude (or times 1 below 1). Default 0, exact.")]
    public double Tolerance { get; init; }

    [Description("The most cell changes to list. Default 200.")]
    public int MaxChanges { get; init; } = 200;
}

public enum CellChangeKind
{
    Value,
    Formula,

    /// <summary>The style differs; <see cref="CellChange.Left"/> names the parts that do, such as <c>font,fill</c>.</summary>
    Style,
}

public enum ItemChangeKind
{
    Added,
    Removed,
    Changed,
}

public sealed record CellChange(string Cell, CellChangeKind Kind, string Left, string Right);

public sealed record ItemChange(string Name, ItemChangeKind Kind, string? Left, string? Right);

public sealed record DiffResult(
    string Left,
    string Right,
    bool Identical,
    IReadOnlyList<string> SheetsAdded,
    IReadOnlyList<string> SheetsRemoved,
    int CellChangeTotal,
    IReadOnlyList<CellChange> CellChanges,
    IReadOnlyList<ItemChange> DefinedNames,
    IReadOnlyList<ItemChange> Tables) : IOperationResult
{
    public int ExitCode => Identical ? 0 : 1;

    public void WriteText(StringBuilder text)
    {
        var count = SheetsAdded.Count + SheetsRemoved.Count + CellChangeTotal + DefinedNames.Count + Tables.Count;
        text.Append(Left).Append(" vs ").Append(Right).Append(": ")
            .Append(Identical ? "identical" : $"{count} difference{(count == 1 ? "" : "s")}").Append('\n');
        foreach (var sheet in SheetsAdded)
            text.Append("sheet added ").Append(sheet).Append('\n');
        foreach (var sheet in SheetsRemoved)
            text.Append("sheet removed ").Append(sheet).Append('\n');
        foreach (var change in CellChanges)
        {
            text.Append(Rendering.Name(change.Kind)).Append(' ').Append(change.Cell).Append('\t').Append(change.Left);
            if (change.Kind != CellChangeKind.Style)
                text.Append(" -> ").Append(change.Right);
            text.Append('\n');
        }

        if (CellChangeTotal > CellChanges.Count)
            text.Append(CellChangeTotal - CellChanges.Count).Append(" more cell changes\n");
        foreach (var name in DefinedNames)
            text.Append("name ").Append(Rendering.Name(name.Kind)).Append(' ').Append(name.Name).Append('\t').Append(name.Left).Append(" -> ").Append(name.Right).Append('\n');
        foreach (var table in Tables)
            text.Append("table ").Append(Rendering.Name(table.Kind)).Append(' ').Append(table.Name).Append('\t').Append(table.Left).Append(" -> ").Append(table.Right).Append('\n');
    }
}

public sealed class DiffOperation : Operation<DiffRequest, DiffResult>
{
    public override string Name => "diff";

    public override string Description =>
        "Compare two workbooks: sheets added or removed, cell value, formula and style changes, defined names and tables. Exit code 0 when identical, 1 when they differ.";

    public override OperationEffect Effect => OperationEffect.ReadOnly;

    public override DiffResult Execute(DiffRequest request, OperationContext context)
    {
        if (request.MaxChanges < 0)
            throw OperationException.BadInput("bad_max_changes", "maxChanges cannot be negative.");

        var recalculate = request.Values == ValueSource.Recalculated;
        using var left = context.LoadWorkbook(request.Left, recalculate);
        using var right = context.LoadWorkbook(request.Right, recalculate);
        return WorkbookComparer.Compare(left.FullPath, left.Workbook, right.FullPath, right.Workbook, new ComparerOptions(
            request.IgnoreSheets, request.IgnoreRanges, request.Styles, request.Tolerance, request.MaxChanges, recalculate));
    }
}

internal sealed record ComparerOptions(
    IReadOnlyList<string> IgnoreSheets,
    IReadOnlyList<string> IgnoreRanges,
    bool Styles,
    double Tolerance,
    int MaxChanges,
    bool Recalculated);

internal static class WorkbookComparer
{
    internal static DiffResult Compare(string leftPath, XLWorkbook left, string rightPath, XLWorkbook right, ComparerOptions options)
    {
        var ignoredSheets = options.IgnoreSheets.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var ignoredRanges = options.IgnoreRanges.Select(ParseRange).ToList();

        var leftSheets = Compared(left, ignoredSheets).ToDictionary(s => s.Name, StringComparer.OrdinalIgnoreCase);
        var rightSheets = Compared(right, ignoredSheets).ToDictionary(s => s.Name, StringComparer.OrdinalIgnoreCase);
        var added = rightSheets.Keys.Where(n => !leftSheets.ContainsKey(n)).ToList();
        var removed = leftSheets.Keys.Where(n => !rightSheets.ContainsKey(n)).ToList();

        var changes = new List<CellChange>();
        var total = 0;
        foreach (var (name, leftSheet) in leftSheets)
        {
            if (!rightSheets.TryGetValue(name, out var rightSheet))
                continue;

            var ignored = ignoredRanges.Where(r => string.Equals(r.Sheet, name, StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (var change in CompareSheets(leftSheet, rightSheet, options))
            {
                var (row, column) = change.Position;
                if (ignored.Any(r => r.Contains(row, column)))
                    continue;
                total++;
                if (changes.Count < options.MaxChanges)
                    changes.Add(change.Change);
            }
        }

        var names = CompareItems(DefinedNames(left, ignoredSheets), DefinedNames(right, ignoredSheets));
        var tables = CompareItems(Tables(left, ignoredSheets), Tables(right, ignoredSheets));
        var identical = added.Count == 0 && removed.Count == 0 && total == 0 && names.Count == 0 && tables.Count == 0;
        return new DiffResult(leftPath, rightPath, identical, added, removed, total, changes, names, tables);
    }

    private static IEnumerable<IXLWorksheet> Compared(XLWorkbook workbook, HashSet<string> ignoredSheets)
        => workbook.Worksheets.Where(s => !ignoredSheets.Contains(s.Name));

    private static IEnumerable<((int Row, int Column) Position, CellChange Change)> CompareSheets(
        IXLWorksheet left, IXLWorksheet right, ComparerOptions options)
    {
        var usage = options.Styles ? XLCellsUsedOptions.All : XLCellsUsedOptions.Contents;
        var positions = new SortedSet<(int Row, int Column)>();
        foreach (var cell in left.CellsUsed(usage).Concat(right.CellsUsed(usage)))
            positions.Add((cell.Address.RowNumber, cell.Address.ColumnNumber));

        foreach (var (row, column) in positions)
        {
            var l = left.Cell(row, column);
            var r = right.Cell(row, column);
            var address = Rendering.Address(r);

            var lf = l.HasFormula ? "=" + l.FormulaA1 : "";
            var rf = r.HasFormula ? "=" + r.FormulaA1 : "";
            if (lf != rf)
                yield return ((row, column), new CellChange(address, CellChangeKind.Formula, lf, rf));

            var lv = Read(l, options.Recalculated);
            var rv = Read(r, options.Recalculated);
            if (!FormulaCells.Same(lv, rv, options.Tolerance))
                yield return ((row, column), new CellChange(address, CellChangeKind.Value, lv.ToText(), rv.ToText()));

            if (options.Styles && !l.Style.Equals(r.Style))
                yield return ((row, column), new CellChange(address, CellChangeKind.Style, StyleParts(l.Style, r.Style), ""));
        }
    }

    private static CellRead Read(IXLCell cell, bool recalculated)
        => cell.HasFormula && !recalculated ? new CellRead(cell.CachedValue, null, null) : CellReader.Read(cell);

    private static string StyleParts(IXLStyle a, IXLStyle b)
    {
        var parts = new List<string>();
        if (a.Font.ToString() != b.Font.ToString())
            parts.Add("font");
        if (a.Fill.ToString() != b.Fill.ToString())
            parts.Add("fill");
        if (a.Border.ToString() != b.Border.ToString())
            parts.Add("border");
        if (a.NumberFormat.Format != b.NumberFormat.Format || a.NumberFormat.NumberFormatId != b.NumberFormat.NumberFormatId)
            parts.Add($"numberFormat({a.NumberFormat.Format}|{b.NumberFormat.Format})");
        if (a.Alignment.ToString() != b.Alignment.ToString())
            parts.Add("alignment");
        if (a.Protection.ToString() != b.Protection.ToString())
            parts.Add("protection");
        return parts.Count == 0 ? "style" : string.Join(',', parts);
    }

    /// <summary>Workbook-scoped names, and the names scoped to every sheet that is not ignored.</summary>
    private static Dictionary<string, string> DefinedNames(XLWorkbook workbook, HashSet<string> ignoredSheets)
        => workbook.DefinedNames.Select(n => (n.Name, n.RefersTo))
            .Concat(Compared(workbook, ignoredSheets).SelectMany(ws => ws.DefinedNames.Select(n => ($"{ws.Name}!{n.Name}", n.RefersTo))))
            .GroupBy(x => x.Item1, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First().RefersTo, StringComparer.OrdinalIgnoreCase);

    private static Dictionary<string, string> Tables(XLWorkbook workbook, HashSet<string> ignoredSheets)
        => Compared(workbook, ignoredSheets).SelectMany(ws => ws.Tables.Select(t => (t.Name, Range: $"{ws.Name}!{t.RangeAddress.ToStringRelative()}")))
            .GroupBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First().Range, StringComparer.OrdinalIgnoreCase);

    private static List<ItemChange> CompareItems(Dictionary<string, string> left, Dictionary<string, string> right)
    {
        var changes = new List<ItemChange>();
        foreach (var (name, value) in left.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase))
        {
            if (!right.TryGetValue(name, out var other))
                changes.Add(new ItemChange(name, ItemChangeKind.Removed, value, null));
            else if (value != other)
                changes.Add(new ItemChange(name, ItemChangeKind.Changed, value, other));
        }

        changes.AddRange(right.Where(p => !left.ContainsKey(p.Key)).OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase)
            .Select(p => new ItemChange(p.Key, ItemChangeKind.Added, null, p.Value)));
        return changes;
    }

    private static IgnoredRange ParseRange(string reference)
    {
        var (sheet, area) = Addressing.Split(reference);
        if (sheet is null)
            throw OperationException.BadInput("bad_ignore_range", $"'{reference}' must name its sheet, as 'Sheet1!A1:B9'.");

        var corners = area.Split(':');
        if (corners.Length > 2 || !TryCell(corners[0], out var first) || !TryCell(corners[^1], out var last))
            throw OperationException.BadInput("bad_ignore_range", $"'{reference}' is not a range such as 'Sheet1!A1:B9'.");
        return new IgnoredRange(sheet, Math.Min(first.Row, last.Row), Math.Min(first.Column, last.Column),
            Math.Max(first.Row, last.Row), Math.Max(first.Column, last.Column));
    }

    private static bool TryCell(string text, out (int Row, int Column) cell)
    {
        cell = default;
        var address = text.Replace("$", "");
        var split = 0;
        while (split < address.Length && char.IsAsciiLetter(address[split]))
            split++;
        if (split == 0 || split == address.Length || !int.TryParse(address[split..], out var row) || row < 1 || row > XLHelper.MaxRowNumber)
            return false;
        var column = XLHelper.GetColumnNumberFromLetter(address[..split]);
        if (column < 1 || column > XLHelper.MaxColumnNumber)
            return false;
        cell = (row, column);
        return true;
    }

    private sealed record IgnoredRange(string Sheet, int Top, int Left, int Bottom, int Right)
    {
        public bool Contains(int row, int column) => row >= Top && row <= Bottom && column >= Left && column <= Right;
    }
}
