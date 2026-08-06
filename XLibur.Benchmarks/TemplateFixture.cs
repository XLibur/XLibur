using System;
using System.IO;
using XLibur.Excel;

namespace XLibur.Benchmarks;

/// <summary>
/// Builds the synthetic reporting template the round-trip benchmarks and probes measure against,
/// and the cell-writing routines that fill it.
/// </summary>
/// <remarks>
/// A generated fixture is a poor stand-in for a genuine reporting workbook: shared strings,
/// styles, tables, spilled dynamic-array formulas, conditional formatting, images and external
/// links all live in a real template and none are reproduced here. Measured side by side against
/// the template that motivated this work, the generated fixture round-tripped roughly an order of
/// magnitude faster, so conclusions drawn from it alone understate real-world cost. It is used
/// because it is reproducible and checked in; point <see cref="TemplateRoundTripProfile"/> at a
/// real file when reproducing a production number matters.
/// </remarks>
internal static class TemplateFixture
{
    public const string DataSheet = "Data";
    public const string LookupSheetPrefix = "Lookup";
    public const string FirstLookupSheet = LookupSheetPrefix + "1";
    public const string LookupRangeName = "LookupRange";
    public const int HeaderRow = 1;
    public const int FirstDataRow = 3;
    public const int GridColumns = 21;
    public const string DateFormat = "mmm/yyyy";

    /// <summary>
    /// Default rows per lookup sheet, enough that a defined name spans a realistic range.
    /// </summary>
    public const int DefaultLookupRows = 100;

    /// <summary>
    /// Builds a workbook approximating a real reporting template: one data sheet plus a number of
    /// lookup sheets, workbook-scoped defined names over the lookup columns, and list data
    /// validations on the data sheet pointing at those names.
    /// </summary>
    /// <param name="lookupRows">
    /// Rows on each lookup sheet. Zero builds sheets that hold nothing but a header, which is what
    /// isolates the structural cost of a worksheet from the cost of the rows on it — the two were
    /// confounded while this was a constant, and the per-sheet slope was reported as a cost per
    /// *empty* sheet when every sheet in it carried 101 rows.
    /// </param>
    public static byte[] Build(
        int sheetCount,
        int definedNames,
        int validations,
        int dataRows,
        int lookupRows = DefaultLookupRows)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(sheetCount, 1);
        ArgumentOutOfRangeException.ThrowIfNegative(lookupRows);

        var lookupSheets = sheetCount - 1;

        // Every defined name spans a lookup column, and every validation points at the first
        // defined name, so neither can be built without a lookup sheet that has rows. Rejecting the
        // combination keeps a probe from quietly measuring a fixture that is not what it asked for.
        if ((lookupSheets == 0 || lookupRows == 0) && (definedNames > 0 || validations > 0))
        {
            throw new ArgumentException(
                "Defined names and validations span a lookup range, so they need at least one lookup sheet with at least one row.",
                nameof(sheetCount));
        }

        using var workbook = new XLWorkbook();

        var data = workbook.AddWorksheet(DataSheet);
        for (var c = 1; c <= GridColumns; c++)
        {
            data.Cell(HeaderRow, c).Value = $"Column{c}";
            data.Cell(HeaderRow, c).Style.Font.Bold = true;
        }

        if (dataRows > 0)
            WriteGrid(data, dataRows, GridColumns, perCellNumberFormat: true);

        // The data sheet is one of the requested sheets, so the rest are lookups. sheetCount: 1 is
        // the data sheet alone; anything else made the workbook one sheet wider than asked for,
        // which mattered because the per-sheet cost is read off the slope between these counts.
        for (var s = 1; s <= lookupSheets; s++)
        {
            var lookup = workbook.AddWorksheet($"{LookupSheetPrefix}{s}");
            lookup.Cell(HeaderRow, 1).Value = "Value";

            for (var r = 1; r <= lookupRows; r++)
                lookup.Cell(HeaderRow + r, 1).Value = $"Sheet {s} value {r}";
        }

        if (lookupSheets > 0 && lookupRows > 0)
        {
            // The first defined name is the one the refresh probe repoints.
            var firstLookup = workbook.Worksheet(FirstLookupSheet);
            workbook.DefinedNames.Add(LookupRangeName, firstLookup.Range(HeaderRow + 1, 1, HeaderRow + lookupRows, 1));

            for (var n = 1; n < definedNames; n++)
            {
                var sheet = workbook.Worksheet($"{LookupSheetPrefix}{(n % lookupSheets) + 1}");
                workbook.DefinedNames.Add($"Name{n}", sheet.Range(HeaderRow + 1, 1, HeaderRow + lookupRows, 1));
            }
        }

        for (var v = 0; v < validations; v++)
        {
            var column = (v % GridColumns) + 1;
            var validation = data.Range(FirstDataRow, column, 1_000, column).CreateDataValidation();
            validation.List($"={LookupRangeName}", true);
        }

        using var buffer = new MemoryStream();
        workbook.SaveAs(buffer);
        return buffer.ToArray();
    }

    /// <summary>Writes a wide grid of mixed cell types, the shape of a data export.</summary>
    public static void WriteGrid(IXLWorksheet sheet, int rows, int columns, bool perCellNumberFormat)
    {
        var date = new DateTime(2026, 7, 1, 0, 0, 0, DateTimeKind.Unspecified);

        for (var r = 0; r < rows; r++)
        {
            var rowIndex = FirstDataRow + r;

            for (var c = 1; c <= columns; c++)
            {
                var cell = sheet.Cell(rowIndex, c);

                switch (c % 4)
                {
                    case 0:
                        cell.Value = date.AddMonths(r % 24);
                        if (perCellNumberFormat)
                            cell.Style.NumberFormat.Format = DateFormat;
                        break;
                    case 1:
                        cell.Value = $"Row {r} column {c} descriptive text";
                        break;
                    case 2:
                        cell.Value = (r * c) % 100_000;
                        break;
                    default:
                        cell.Value = r % 2 == 0 ? "Yes" : "No";
                        break;
                }
            }
        }
    }

    public static void WriteDateColumn(IXLWorksheet sheet, int rows, int column, bool perCellFormat)
    {
        var date = new DateTime(2026, 7, 1, 0, 0, 0, DateTimeKind.Unspecified);

        for (var r = 0; r < rows; r++)
        {
            var cell = sheet.Cell(FirstDataRow + r, column);
            cell.Value = date.AddMonths(r % 24);

            if (perCellFormat)
                cell.Style.NumberFormat.Format = DateFormat;
        }
    }
}
