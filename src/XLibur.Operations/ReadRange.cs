using System.ComponentModel;
using System.Text;
using XLibur.Excel;
using XLibur.Operations.Workbooks;

namespace XLibur.Operations;

public enum ReadMode
{
    /// <summary>Each cell's value.</summary>
    Values,

    /// <summary>Each formula cell's formula text, and the value of every other cell.</summary>
    Formulas,

    /// <summary>The value of every cell, and the formula text of formula cells.</summary>
    Both,
}

public sealed record ReadRangeRequest
{
    [Argument(0), Description("The .xlsx or .xlsm file.")]
    public required string Path { get; init; }

    [Argument(1), Description("What to read: 'Sheet1!A1:D50', 'A1:D50' (on sheet, or the first sheet), 'A:C', or a defined name. Defaults to the used range. Pass a result's 'next' here to read the following page.")]
    public string? Range { get; init; }

    [Description("The sheet an unqualified range is on. Defaults to the first sheet.")]
    public string? Sheet { get; init; }

    [Description("values (default), formulas, or both.")]
    public ReadMode Mode { get; init; } = ReadMode.Values;

    [Description("The most cells to return; whole rows are returned, and 'next' names the rest. Default 2000.")]
    public int MaxCells { get; init; } = 2000;

    [Description("Recalculate every formula on load instead of returning the values cached in the file.")]
    public bool Recalculate { get; init; }
}

/// <summary>
/// A block of cells. Text form is TSV: a line of column letters, then one line per row headed by its
/// row number, trailing blanks dropped. JSON form is row-major arrays.
/// </summary>
public sealed record ReadRangeResult(
    string Sheet,
    string? Range,
    int FirstRow,
    int FirstColumn,
    IReadOnlyList<IReadOnlyList<XLCellValue>> Values,
    IReadOnlyList<IReadOnlyList<string?>>? Formulas,
    IReadOnlyDictionary<string, string>? Failures,
    string? Next) : IOperationResult
{
    internal ReadMode Mode { get; init; }

    public void WriteText(StringBuilder text)
    {
        text.Append(Range ?? $"{Sheet}: empty").Append('\n');
        if (Values.Count > 0)
        {
            var width = Values.Max(r => r.Count);
            for (var c = 0; c < width; c++)
                text.Append('\t').Append(XLHelper.GetColumnLetterFromNumber(FirstColumn + c));
            text.Append('\n');
        }

        for (var r = 0; r < Values.Count; r++)
        {
            text.Append(FirstRow + r);
            var fields = Enumerable.Range(0, Values[r].Count).Select(c => Field(r, c)).ToList();
            var width = fields.FindLastIndex(f => f.Length > 0) + 1;
            for (var c = 0; c < width; c++)
                text.Append('\t').Append(fields[c]);
            text.Append('\n');
        }

        if (Next is not null)
            text.Append("next: ").Append(Next).Append('\n');
    }

    private string Field(int r, int c)
    {
        var address = XLHelper.GetColumnLetterFromNumber(FirstColumn + c) + (FirstRow + r);
        var value = Failures is not null && Failures.TryGetValue(address, out var failure)
            ? "!" + failure
            : Rendering.ToText(Values[r][c]);
        var formula = Formulas?[r][c];
        if (formula is null)
            return value;
        return Mode == ReadMode.Formulas ? "=" + Rendering.Escape(formula) : $"={Rendering.Escape(formula)} -> {value}";
    }
}

public sealed class ReadRangeOperation : Operation<ReadRangeRequest, ReadRangeResult>
{
    public override string Name => "read_range";

    public override string Description =>
        "Read cell values and/or formulas from a range as compact TSV (or JSON). Large ranges are paged: whole rows up to maxCells, with 'next' naming the remaining range.";

    public override OperationEffect Effect => OperationEffect.ReadOnly;

    public override ReadRangeResult Execute(ReadRangeRequest request, OperationContext context)
    {
        if (request.MaxCells < 1)
            throw OperationException.BadInput("bad_max_cells", "maxCells must be at least 1.");
        var maxCells = Math.Min(request.MaxCells, context.Limits.MaxCells);

        using var loaded = context.LoadWorkbook(request.Path, request.Recalculate);
        var requested = Addressing.Range(loaded.Workbook, request.Sheet, request.Range);
        var sheet = requested.Worksheet;
        var used = sheet.RangeUsed(XLCellsUsedOptions.Contents);
        var area = used is null ? null : Intersect(requested.RangeAddress, used.RangeAddress);
        if (area is null)
            return new ReadRangeResult(sheet.Name, null, 0, 0, [], null, null, null) { Mode = request.Mode };

        var (top, left, bottom, right) = area.Value;
        var width = right - left + 1;
        if (width > maxCells)
            throw OperationException.BadInput("range_too_wide",
                $"The range is {width:N0} columns wide, more than maxCells ({maxCells:N0}). Read fewer columns.");

        var rowsPerPage = maxCells / width;
        var last = Math.Min(bottom, top + rowsPerPage - 1);
        var block = sheet.Range(top, left, last, right);

        var values = new XLCellValue[last - top + 1][];
        for (var r = 0; r < values.Length; r++)
            values[r] = new XLCellValue[width];
        string?[][]? formulas = null;
        Dictionary<string, string>? failures = null;

        foreach (var cell in block.CellsUsed(XLCellsUsedOptions.Contents))
        {
            var r = cell.Address.RowNumber - top;
            var c = cell.Address.ColumnNumber - left;
            var read = CellReader.Read(cell);
            values[r][c] = read.Value;
            if (read.Failure is { } failure)
                (failures ??= new())[cell.Address.ToStringRelative()] = Rendering.Name(failure);
            if (request.Mode != ReadMode.Values && cell.HasFormula)
            {
                formulas ??= Enumerable.Range(0, values.Length).Select(_ => new string?[width]).ToArray();
                formulas[r][c] = cell.FormulaA1;
            }
        }

        var next = last < bottom
            ? $"{QuoteSheet(sheet)}!{XLHelper.GetColumnLetterFromNumber(left)}{last + 1}:{XLHelper.GetColumnLetterFromNumber(right)}{bottom}"
            : null;

        return new ReadRangeResult(sheet.Name, $"{QuoteSheet(sheet)}!{Rendering.Area(block.RangeAddress)}", top, left, values, formulas, failures, next)
        {
            Mode = request.Mode,
        };
    }

    /// <summary>
    /// The sheet name as a reference writes it, quoted where Excel needs quotes. XLibur has no public
    /// quoting call, so the name is cut from the address of the sheet's first cell, which it quotes.
    /// </summary>
    internal static string QuoteSheet(IXLWorksheet sheet)
    {
        var address = sheet.Cell(1, 1).Address.ToStringRelative(includeSheet: true);
        return address[..address.LastIndexOf('!')];
    }

    private static (int Top, int Left, int Bottom, int Right)? Intersect(XLRangeAddress a, XLRangeAddress b)
    {
        var top = Math.Max(a.FirstAddress.RowNumber, b.FirstAddress.RowNumber);
        var left = Math.Max(a.FirstAddress.ColumnNumber, b.FirstAddress.ColumnNumber);
        var bottom = Math.Min(a.LastAddress.RowNumber, b.LastAddress.RowNumber);
        var right = Math.Min(a.LastAddress.ColumnNumber, b.LastAddress.ColumnNumber);
        return top <= bottom && left <= right ? (top, left, bottom, right) : null;
    }
}
