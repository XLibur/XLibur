using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// Re-points the formulas of every worksheet in a workbook after a row or column shift.
/// </summary>
/// <remarks>
/// Every worksheet is visited, not just the shifted one, because a formula on any sheet may refer to
/// the sheet being shifted; <c>XLCellFormulaShifter</c> is what decides which references that shift
/// actually reaches.
/// <para>
/// The walk is over the formula slice rather than over used cells. The cell-level enumeration visited
/// every used cell in the sheet and materialised an <c>XLCell</c> facade for each — through a 16-slot
/// direct-mapped cache that a linear scan defeats — only to discard the ones with no formula. The
/// formula slice already knows where the formulas are, so a value-only sheet costs nothing here.
/// </para>
/// </remarks>
internal static class XLFormulaShiftPass
{
    internal static void Run(XLWorkbook workbook, XLRange shiftedRange, bool shiftRows, int shift)
    {
        // Array and data-table formulas share one instance across every cell of their range and must
        // be shifted exactly once; the set is shared across sheets because the pass is.
        var processedArrayFormulas = new HashSet<XLCellFormula>();

        foreach (var sheet in workbook.WorksheetsInternal)
        {
            var cellsCollection = sheet.Internals.CellsCollection;
            var enumerator = cellsCollection.FormulaSlice.GetForwardEnumerator(Coordinates.Area.Full);

            // Materialise the points first. The shift writes back into the same slice (a rewritten
            // normal formula is a new XLCellFormula instance at the same point), and mutating a slice
            // while its enumerator is live is not a contract the slice offers.
            var points = new List<Coordinates.Point>();
            while (enumerator.MoveNext())
                points.Add(enumerator.Point);

            foreach (var point in points)
            {
                var cell = cellsCollection.GetCell(point);
                if (shiftRows)
                    cell.ShiftFormulaRows(shiftedRange, shift, processedArrayFormulas);
                else
                    cell.ShiftFormulaColumns(shiftedRange, shift, processedArrayFormulas);
            }
        }
    }
}
