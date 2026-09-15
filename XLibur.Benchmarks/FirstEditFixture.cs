using System;
using System.IO;
using XLibur.Excel;

namespace XLibur.Benchmarks;

/// <summary>
/// The workbook that <see cref="FirstEditAfterLoadBenchmarks"/> and <see cref="FirstEditProfile"/>
/// load: every formula is calculated before the save, so each loads clean with a cached value and
/// the first edit after the load has to build the whole dependency tree (#504, #513).
/// </summary>
internal static class FirstEditFixture
{
    /// <summary>
    /// Build the package.
    /// </summary>
    /// <param name="rows">Rows of the fixture.</param>
    /// <param name="formulasPerRow">
    /// <c>2</c> is the narrow shape #504 measured: two value columns and two formulas. <c>8</c> is
    /// the wide shape of the 200,000-formula row in #513: seven value columns and eight formulas,
    /// 15 columns in all, as in <c>LoadAndReadAllCells</c>.
    /// </param>
    public static byte[] Build(int rows, int formulasPerRow)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Data");
        for (var row = 1; row <= rows; row++)
        {
            switch (formulasPerRow)
            {
                case 2:
                    AddNarrowRow(sheet, row);
                    break;
                case 8:
                    AddWideRow(sheet, row);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(formulasPerRow), formulasPerRow, "Only 2 or 8.");
            }
        }

        // Calculated, so that each formula is saved with a cached value and loads clean.
        workbook.RecalculateAllFormulas();

        using var buffer = new MemoryStream();
        workbook.SaveAs(buffer);
        return buffer.ToArray();
    }

    private static void AddNarrowRow(IXLWorksheet sheet, int row)
    {
        sheet.Cell(row, 1).Value = row;
        sheet.Cell(row, 2).Value = row * 2;
        sheet.Cell(row, 3).FormulaA1 = $"A{row}*B{row}+1";
        sheet.Cell(row, 4).FormulaA1 = $"IF(C{row}>100,SUM(A{row}:C{row}),C{row}/2)";
    }

    /// <remarks>
    /// The formulas mix single cells, row ranges, a formula that reads other formulas, and one
    /// absolute reference that every row shares, so one precedent area has 25,000 dependents.
    /// No value is zero, so no formula divides by zero.
    /// </remarks>
    private static void AddWideRow(IXLWorksheet sheet, int row)
    {
        sheet.Cell(row, 1).Value = row;
        sheet.Cell(row, 2).Value = row * 2;
        sheet.Cell(row, 3).Value = row % 7 + 1;
        sheet.Cell(row, 4).Value = row % 13 + 1;
        sheet.Cell(row, 5).Value = row % 5 + 1;
        sheet.Cell(row, 6).Value = row * 0.5;
        sheet.Cell(row, 7).Value = row % 3 + 1;
        sheet.Cell(row, 8).FormulaA1 = $"A{row}*B{row}+1";
        sheet.Cell(row, 9).FormulaA1 = $"IF(H{row}>100,SUM(A{row}:G{row}),H{row}/2)";
        sheet.Cell(row, 10).FormulaA1 = $"SUM(A{row}:G{row})";
        sheet.Cell(row, 11).FormulaA1 = $"AVERAGE(B{row}:F{row})*C{row}";
        sheet.Cell(row, 12).FormulaA1 = $"ROUND(D{row}/E{row},2)";
        sheet.Cell(row, 13).FormulaA1 = $"MAX(A{row},B{row},C{row})-MIN(D{row},E{row})";
        sheet.Cell(row, 14).FormulaA1 = $"J{row}+K{row}-L{row}";
        sheet.Cell(row, 15).FormulaA1 = $"IF(ISERROR(N{row}),0,N{row}*$G$1)";
    }
}
