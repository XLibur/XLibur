using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// Deletes a set of whole rows — contiguous or not — re-pointing every formula once for the whole set
/// rather than once per row.
/// </summary>
/// <remarks>
/// Deleting rows one at a time makes the formula pass, which visits every formula in the workbook, run
/// once per deleted row. On a sheet with a formula per row that is the dominant cost of the whole
/// operation and it grows quadratically: 1,333 deletes over 2,667 surviving formulas is 3.6M parses,
/// nearly all of which conclude that nothing needed changing.
/// <para>
/// Here the rows become one <see cref="XLRowDeletionMap"/>, every formula is re-pointed against it in a
/// single pass, and only then are the rows removed run by run with the per-run formula pass switched
/// off. The removal itself is still per-run — the live ranges, conditional formats, data validations,
/// merges, page breaks and hyperlinks all still shift a block at a time, which is correct because those
/// shifts compose, just not yet batched.
/// </para>
/// </remarks>
internal static class XLRowBatchDelete
{
    internal static void Delete(XLWorksheet worksheet, IEnumerable<int> rowNumbers)
    {
        var map = XLRowDeletionMap.Create(rowNumbers);
        if (map is null)
            return;

        var batched = CanBatch(worksheet.Workbook);
        if (batched)
            ShiftAllFormulas(worksheet.Workbook, worksheet.Name, map);

        foreach (var (firstRow, lastRow) in map.GetRunsBottomUp())
        {
            var range = (XLRange)worksheet.Range(firstRow, 1, lastRow, XLHelper.MaxColumnNumber);
            range.Delete(XLShiftDeletedCells.ShiftCellsUp, shiftFormulas: !batched);

            for (var row = lastRow; row >= firstRow; row--)
                worksheet.DeleteRow(row);
        }
    }

    /// <summary>
    /// Whether the composite pass can stand in for the per-run passes.
    /// </summary>
    /// <remarks>
    /// Array, data-table and dynamic-array formulas carry a stored range or a spill footprint that the
    /// per-cell shift relocates as a side effect of re-pointing the text. The composite pass rewrites
    /// text only, so a workbook containing any of them keeps the per-run path — correctness over speed,
    /// and the case is rare enough that the scan to detect it is cheaper than the work it guards.
    /// </remarks>
    private static bool CanBatch(XLWorkbook workbook)
    {
        foreach (var sheet in workbook.WorksheetsInternal)
        {
            var enumerator = sheet.Internals.CellsCollection.FormulaSlice.GetForwardEnumerator(Area.Full);
            while (enumerator.MoveNext())
            {
                var formula = enumerator.Current;
                if (formula.Type != FormulaType.Normal || formula.IsDynamicArray)
                    return false;
            }
        }

        return true;
    }

    private static void ShiftAllFormulas(XLWorkbook workbook, string shiftedSheetName, XLRowDeletionMap map)
    {
        var firstDeletedRow = map.FirstDeletedRow;

        foreach (var sheet in workbook.WorksheetsInternal)
        {
            var cellsCollection = sheet.Internals.CellsCollection;

            // Materialise the points first: rewriting a formula replaces the entry at its own point,
            // and mutating a slice while its enumerator is live is not a contract the slice offers.
            var points = new List<Point>();
            var enumerator = cellsCollection.FormulaSlice.GetForwardEnumerator(Area.Full);
            while (enumerator.MoveNext())
                points.Add(enumerator.Point);

            foreach (var point in points)
            {
                var cell = cellsCollection.GetCell(point);
                var formula = cell.Formula;

                // Same pre-filter the per-run pass uses: a formula whose furthest reference stops above
                // every deleted row cannot be rewritten by any of them.
                if (formula is null || formula.MaxShiftableRow < firstDeletedRow)
                    continue;

                var shifted = XLCellFormulaShifter.ShiftFormulaRows(formula.A1, sheet, shiftedSheetName, map);
                if (string.Equals(shifted, formula.A1, System.StringComparison.Ordinal))
                    continue;

                cell.FormulaA1 = shifted;
                cell.Formula?.SeedShiftedExtentFrom(formula, map);
            }
        }
    }

    /// <summary>
    /// The rows an <see cref="IXLRows"/> covers, grouped by the sheet they live on.
    /// </summary>
    internal static IEnumerable<IGrouping<IXLWorksheet, int>> GroupBySheet(IEnumerable<IXLRow> rows)
        => rows.GroupBy(r => r.Worksheet, r => r.RowNumber());
}
