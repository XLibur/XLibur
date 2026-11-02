using System.ComponentModel;
using System.Globalization;
using System.Text;
using System.Text.Json;
using XLibur.Excel;
using XLibur.Operations.Workbooks;

namespace XLibur.Operations;

public sealed record ConvertRequest
{
    [Argument(0), Description("The input: .xlsx, .xlsm, .csv, .tsv or .json. JSON is an array of row arrays, or an array of objects whose keys become the header row.")]
    public required string Input { get; init; }

    [Argument(1), Description("The output; its extension picks the format.")]
    public required string Output { get; init; }

    [Description("The sheet to export from a workbook, or the sheet name to create from a table. Defaults to the first sheet, or 'Sheet1'.")]
    public string? Sheet { get; init; }

    [Description("Replace the output if it exists.")]
    public bool Overwrite { get; init; }

    [Description("Recalculate a workbook's formulas before exporting or saving it.")]
    public bool Recalculate { get; init; }

    [Description("JSON output only: write rows as objects keyed by the first row, instead of arrays.")]
    public bool Header { get; init; }
}

public sealed record ConvertResult(string Output, string Format, int Rows, int Columns, int Failures) : IOperationResult
{
    public void WriteText(StringBuilder text)
    {
        text.Append("wrote ").Append(Output).Append(" (").Append(Format).Append(", ")
            .Append(Rows).Append(" rows x ").Append(Columns).Append(" columns");
        if (Failures > 0)
            text.Append(", ").Append(Failures).Append(" cells failed to calculate and were left blank");
        text.Append(")\n");
    }
}

public sealed class ConvertOperation : Operation<ConvertRequest, ConvertResult>
{
    private static readonly Dictionary<string, TableFormat> Formats = new(StringComparer.OrdinalIgnoreCase)
    {
        [".csv"] = new("csv", bytes => Delimited.Read(Decode(bytes), ','), (grid, _) => Delimited.Write(grid, ',')),
        [".tsv"] = new("tsv", bytes => Delimited.Read(Decode(bytes), '\t'), (grid, _) => Delimited.Write(grid, '\t')),
        [".json"] = new("json", JsonTable.Read, JsonTable.Write),
    };

    public override string Name => "convert";

    public override string Description =>
        "Convert between .xlsx/.xlsm and .csv, .tsv or .json (one sheet, values only), or between table formats. xlsx to xlsx with recalculate saves recalculated cached values.";

    public override OperationEffect Effect => OperationEffect.Writes;

    public override ConvertResult Execute(ConvertRequest request, OperationContext context)
    {
        var target = context.Paths.ResolveWrite(request.Output, request.Overwrite);
        var destination = target.FullPath;
        var inputIsWorkbook = OperationContext.IsWorkbookExtension(request.Input);
        var outputIsWorkbook = OperationContext.IsWorkbookExtension(destination);
        var outputExtension = Path.GetExtension(destination);

        if (inputIsWorkbook && outputIsWorkbook)
        {
            if (!string.Equals(Path.GetExtension(request.Input), outputExtension, StringComparison.OrdinalIgnoreCase))
                throw OperationException.Unsupported("xlsx_xlsm_conversion", "Converting between .xlsx and .xlsm is not supported; the macro part would be lost or invented.");
            using var loaded = context.LoadWorkbook(request.Input, request.Recalculate);
            var used = Addressing.Sheet(loaded.Workbook, request.Sheet).RangeUsed(XLCellsUsedOptions.Contents);
            context.SaveWorkbook(loaded.Workbook, target);
            return new ConvertResult(destination, outputExtension.TrimStart('.'), used?.RowCount() ?? 0, used?.ColumnCount() ?? 0, 0);
        }

        if (outputIsWorkbook && !string.Equals(outputExtension, ".xlsx", StringComparison.OrdinalIgnoreCase))
            throw OperationException.Unsupported("table_to_xlsm", "A table converts to .xlsx only; a new workbook has no macros to make it an .xlsm.");

        var (grid, failures) = inputIsWorkbook ? ReadWorkbook(context, request) : ReadTable(context, request.Input);
        var columns = grid.Count == 0 ? 0 : grid.Max(r => r.Length);

        if (outputIsWorkbook)
        {
            using var workbook = WorkbookLoader.Create();
            var sheet = workbook.AddWorksheet(request.Sheet ?? "Sheet1");
            for (var r = 0; r < grid.Count; r++)
            for (var c = 0; c < grid[r].Length; c++)
            {
                if (!grid[r][c].IsBlank)
                    sheet.Cell(r + 1, c + 1).Value = grid[r][c];
            }

            context.SaveWorkbook(workbook, target);
            return new ConvertResult(destination, outputExtension.TrimStart('.'), grid.Count, columns, failures);
        }

        var format = FormatOf(destination);
        var bytes = Encoding.UTF8.GetBytes(format.Write(grid, request.Header));
        context.Paths.WriteAtomically(target, stream => stream.Write(bytes));
        return new ConvertResult(destination, format.Name, grid.Count, columns, failures);
    }

    private static (List<XLCellValue[]> Grid, int Failures) ReadWorkbook(OperationContext context, ConvertRequest request)
    {
        using var loaded = context.LoadWorkbook(request.Input, request.Recalculate);
        var sheet = Addressing.Sheet(loaded.Workbook, request.Sheet);
        var used = sheet.RangeUsed(XLCellsUsedOptions.Contents);
        if (used is null)
            return ([], 0);

        var rows = used.RangeAddress.LastAddress.RowNumber;
        var columns = used.RangeAddress.LastAddress.ColumnNumber;
        var grid = Enumerable.Range(0, rows).Select(_ => new XLCellValue[columns]).ToList();
        var failures = 0;
        foreach (var cell in used.CellsUsed(XLCellsUsedOptions.Contents))
        {
            var read = CellReader.Read(cell);
            if (read.Failure is not null)
                failures++;
            grid[cell.Address.RowNumber - 1][cell.Address.ColumnNumber - 1] = read.Value;
        }

        return (grid, failures);
    }

    private static (List<XLCellValue[]> Grid, int Failures) ReadTable(OperationContext context, string input)
    {
        var (fullPath, bytes) = context.ReadFile(input);
        try
        {
            return (FormatOf(fullPath).Read(bytes), 0);
        }
        catch (Exception e) when (e is JsonException or FormatException)
        {
            throw OperationException.BadInput("unreadable_table", $"'{input}' cannot be read: {e.Message}", e);
        }
    }

    private static TableFormat FormatOf(string path)
        => Formats.TryGetValue(Path.GetExtension(path), out var format)
            ? format
            : throw OperationException.BadInput("unknown_format", $"'{path}' is not a format convert reads or writes.");

    private static string Decode(byte[] bytes) => new UTF8Encoding(false, true).GetString(bytes).TrimStart('\uFEFF');

    private sealed record TableFormat(string Name, Func<byte[], List<XLCellValue[]>> Read, Func<List<XLCellValue[]>, bool, string> Write);
}

/// <summary>RFC 4180 CSV, and TSV with the escapes read_range uses (\t, \n, \r, \\).</summary>
internal static class Delimited
{
    internal static List<XLCellValue[]> Read(string text, char separator)
    {
        var rows = new List<XLCellValue[]>();
        var row = new List<XLCellValue>();
        var field = new StringBuilder();
        var quoted = false;
        var wasQuoted = false;
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (quoted)
            {
                if (c == '"' && i + 1 < text.Length && text[i + 1] == '"')
                {
                    field.Append('"');
                    i++;
                }
                else if (c == '"')
                    quoted = false;
                else
                    field.Append(c);
            }
            else if (c == '"' && field.Length == 0 && separator == ',')
            {
                quoted = true;
                wasQuoted = true;
            }
            else if (c == separator)
            {
                row.Add(Value(field, wasQuoted, separator));
                wasQuoted = false;
            }
            else if (c is '\n' or '\r')
            {
                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                    i++;
                row.Add(Value(field, wasQuoted, separator));
                wasQuoted = false;
                rows.Add([.. row]);
                row.Clear();
            }
            else
                field.Append(c);
        }

        if (quoted)
            throw new FormatException("A quoted field is not closed.");
        if (field.Length > 0 || row.Count > 0)
        {
            row.Add(Value(field, wasQuoted, separator));
            rows.Add([.. row]);
        }

        return rows;
    }

    private static XLCellValue Value(StringBuilder field, bool quoted, char separator)
    {
        var text = field.ToString();
        field.Clear();
        if (separator == '\t')
            text = Unescape(text);
        if (quoted)
            return text;
        if (text.Length == 0)
            return Blank.Value;
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            return number;
        if (text is "TRUE" or "FALSE")
            return text == "TRUE";
        return text;
    }

    private static string Unescape(string text)
    {
        if (!text.Contains('\\'))
            return text;
        var result = new StringBuilder(text.Length);
        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] != '\\' || i + 1 == text.Length)
            {
                result.Append(text[i]);
                continue;
            }

            result.Append(text[++i] switch
            {
                't' => '\t',
                'n' => '\n',
                'r' => '\r',
                var other => other,
            });
        }

        return result.ToString();
    }

    internal static string Write(List<XLCellValue[]> grid, char separator)
    {
        var text = new StringBuilder();
        foreach (var row in grid)
        {
            for (var c = 0; c < row.Length; c++)
            {
                if (c > 0)
                    text.Append(separator);
                var cell = row[c];
                if (cell.IsText)
                {
                    var safe = NeutraliseFormula(cell.GetText());
                    text.Append(separator == '\t' ? Rendering.Escape(safe) : CsvField(safe));
                }
                else
                {
                    var value = Rendering.ToText(cell);
                    text.Append(separator == '\t' ? value : CsvField(value));
                }
            }

            text.Append("\r\n");
        }

        return text.ToString();
    }

    /// <summary>Text that Excel would read as a formula when the file is opened gets a leading quote. Numbers never reach here.</summary>
    private static string NeutraliseFormula(string text)
        => text.Length > 0 && text[0] is '=' or '+' or '-' or '@' or '\t' or '\r' ? "'" + text : text;

    private static string CsvField(string value)
        => value.AsSpan().IndexOfAny(",\"\r\n") < 0 ? value : "\"" + value.Replace("\"", "\"\"") + "\"";
}

internal static class JsonTable
{
    internal static List<XLCellValue[]> Read(byte[] bytes)
    {
        using var document = JsonDocument.Parse(bytes);
        if (document.RootElement.ValueKind != JsonValueKind.Array)
            throw new FormatException("A JSON table is an array of rows.");

        var rows = document.RootElement.EnumerateArray().ToList();
        if (rows.Count > 0 && rows.All(r => r.ValueKind == JsonValueKind.Object))
        {
            var headers = rows.SelectMany(r => r.EnumerateObject().Select(p => p.Name)).Distinct().ToList();
            var grid = new List<XLCellValue[]> { headers.Select(h => (XLCellValue)h).ToArray() };
            grid.AddRange(rows.Select(r => headers.Select(h => r.TryGetProperty(h, out var v) ? Cell(v) : Blank.Value).ToArray()));
            return grid;
        }

        return rows.Select(r => r.ValueKind == JsonValueKind.Array
            ? r.EnumerateArray().Select(Cell).ToArray()
            : throw new FormatException("Every row of a JSON table is an array, or every row is an object.")).ToList();
    }

    private static XLCellValue Cell(JsonElement element) => element.Deserialize<XLCellValue>(Rendering.JsonOptions);

    internal static string Write(List<XLCellValue[]> grid, bool header)
    {
        if (!header || grid.Count == 0)
            return JsonSerializer.Serialize(grid, Rendering.JsonOptions);

        var keys = grid[0].Select((v, i) => v.IsBlank ? $"column{i + 1}" : v.ToString(CultureInfo.InvariantCulture)).ToArray();
        var rows = grid.Skip(1).Select(row =>
        {
            var item = new Dictionary<string, XLCellValue>();
            for (var c = 0; c < keys.Length; c++)
                item[keys[c]] = c < row.Length ? row[c] : Blank.Value;
            return item;
        });
        return JsonSerializer.Serialize(rows, Rendering.JsonOptions);
    }
}
