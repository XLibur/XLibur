using XLibur.Excel;

namespace XLibur.Operations.Workbooks;

internal static class Addressing
{
    internal static IXLWorksheet Sheet(XLWorkbook workbook, string? name)
    {
        if (workbook.Worksheets.Count == 0)
            throw OperationException.BadInput("no_sheets", "The workbook has no worksheets.");
        if (name is null)
            return workbook.Worksheet(1);
        if (workbook.TryGetWorksheet(name, out var sheet))
            return sheet;
        throw OperationException.BadInput("sheet_not_found",
            $"No sheet '{name}'. Sheets: {string.Join(", ", workbook.Worksheets.Select(s => s.Name))}.");
    }

    internal static IXLRange Range(XLWorkbook workbook, string? sheet, string? reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
        {
            var ws = Sheet(workbook, sheet);
            return ws.RangeUsed(XLCellsUsedOptions.Contents) ?? ws.Range("A1");
        }

        var (sheetPart, area) = Split(reference.Trim());
        if (sheetPart is null && sheet is null && workbook.DefinedNames.TryGetValue(area, out var name))
        {
            return name.Ranges.FirstOrDefault()
                   ?? throw OperationException.BadInput("name_not_a_range", $"The defined name '{area}' does not refer to a range.");
        }

        var target = Sheet(workbook, sheetPart ?? sheet);
        if (sheetPart is null && target.DefinedNames.TryGetValue(area, out var local) && local.Ranges.FirstOrDefault() is { } localRange)
            return localRange;

        try
        {
            return target.Range(area);
        }
        catch (Exception e) when (e is ArgumentException or FormatException or InvalidOperationException)
        {
            throw OperationException.BadInput("bad_address", $"'{reference}' is not a range address or a defined name.", e);
        }
    }

    internal static IXLCell Cell(XLWorkbook workbook, string? sheet, string reference)
    {
        var range = Range(workbook, sheet, reference);
        if (range.RangeAddress.FirstAddress.RowNumber != range.RangeAddress.LastAddress.RowNumber ||
            range.RangeAddress.FirstAddress.ColumnNumber != range.RangeAddress.LastAddress.ColumnNumber)
            throw OperationException.BadInput("not_a_cell", $"'{reference}' names a range, not one cell.");
        return range.FirstCell();
    }

    /// <summary>Splits <c>'It''s'!A1</c> into <c>It's</c> and <c>A1</c>. A sheet name may hold '!', an area never does.</summary>
    internal static (string? Sheet, string Area) Split(string reference)
    {
        var bang = reference.LastIndexOf('!');
        if (bang < 0)
            return (null, reference);

        var sheet = reference[..bang];
        if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[^1] == '\'')
            sheet = sheet[1..^1].Replace("''", "'");
        return (sheet, reference[(bang + 1)..]);
    }
}
