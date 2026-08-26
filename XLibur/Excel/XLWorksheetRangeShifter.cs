using System;
using System.Linq;
using XLibur.Excel.ConditionalFormats;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel;

/// <summary>
/// Handles range shifting (insert/delete rows/columns) for a worksheet: conditional formats, data
/// validations, page breaks, defined names, sparklines and calc-engine notifications. Was six mirror
/// pairs against three shared methods; now one implementation per step, bound to an
/// <see cref="IGridAxis"/> through a generic type argument (spec 26 task 8).
/// </summary>
internal sealed class XLWorksheetRangeShifter(XLWorksheet worksheet)
{
    public void ShiftColumns(XLRange range, int columnsShifted) => Shift<ColumnAxis>(range, columnsShifted);

    public void ShiftRows(XLRange range, int rowsShifted) => Shift<RowAxis>(range, rowsShifted);

    /// <summary>
    /// The per-shift passes, in one fixed order for both axes. <b>The order is pinned deliberately.</b>
    /// The two copies used to disagree — <c>ShiftColumns</c> ran page breaks then sparkline cleanup and
    /// <c>ShiftRows</c> the reverse, with nothing saying whether it mattered. It does not:
    /// <see cref="RemoveInvalidSparklines"/> reads only <c>SparklineGroups</c> and address validity while
    /// <see cref="ShiftPageBreaks{TAxis}"/> touches only <c>PageSetup.*Breaks</c>, so they commute.
    /// <c>GridAxisSymmetryTests.Page_breaks_and_sparklines_survive_a_shift_on_both_axes</c> pins it, and
    /// ran green on the two-order code before this collapse, so one order is provably a no-op.
    /// The first three steps are <em>not</em> interchangeable: coverage is shifted from the pre-shift
    /// address, and the validation formula pass must follow the range pass, which deletes a rule whose
    /// coverage transforms to nothing.
    /// </summary>
    private void Shift<TAxis>(XLRange range, int shift)
        where TAxis : struct, IGridAxis
    {
        SplitMergedRangesCrossingTheShift<TAxis>(range);

        worksheet.Workbook.WorksheetsInternal.ForEach<XLWorksheet>(ws => MoveDefinedNames<TAxis>(range, shift, ws.DefinedNames));
        MoveDefinedNames<TAxis>(range, shift, worksheet.Workbook.DefinedNamesInternal);
        ShiftConditionalFormatting<TAxis>(range, shift);
        ShiftDataValidation<TAxis>(range, shift);
        ShiftDataValidationFormula<TAxis>(range, shift);
        ShiftPageBreaks<TAxis>(range, shift);
        RemoveInvalidSparklines();

        Notify<TAxis>(range, shift);
    }

    /// <summary>
    /// Hands the edit to every listener <see cref="XLWorksheet.GetSheetListeners"/> yields, in that
    /// order. This is the only place a sheet listener is reached; adding a feature that must survive
    /// a structural edit is one <see cref="ISheetListener"/> implementation and one
    /// <c>yield return</c>, with nothing to change here.
    /// </summary>
    /// <remarks>
    /// A shift of zero notifies nobody, which is what the two hardcoded blocks this replaced did —
    /// they were an <c>if</c>/<c>else if</c> with no <c>else</c>.
    /// </remarks>
    private void Notify<TAxis>(XLRange range, int shift)
        where TAxis : struct, IGridAxis
    {
        if (shift == 0)
            return;

        var axis = default(TAxis);
        var edit = new SheetEdit
        {
            Sheet = range.Worksheet,
            Area = shift > 0
                ? axis.ExtendAlongIndex(Area.FromRangeAddress(range.RangeAddress), shift - 1)
                : Area.FromRangeAddress(range.RangeAddress),
            Range = range,
            Shift = shift,
        };

        foreach (var listener in worksheet.GetSheetListeners())
        {
            if (shift > 0)
                axis.OnInsertAreaAndShift(listener, in edit);
            else
                axis.OnDeleteAreaAndShift(listener, in edit);
        }
    }

    /// <summary>A merged range the shift would tear — one extending past the shifted range on the cross
    /// axis — is unmerged rather than moved. An entire-line range cannot be torn, so it is left alone.</summary>
    private void SplitMergedRangesCrossingTheShift<TAxis>(XLRange range)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        if (axis.IsEntireLine(range))
            return;

        var first = range.RangeAddress.FirstAddress;
        var last = range.RangeAddress.LastAddress;
        var model = new XLRangeAddress((XLAddress)first, axis.AddressAtMaxIndex(last));
        var rangesToSplit = worksheet.MergedRanges
            .GetIntersectedRanges(model)
            .Where(r => axis.CrossOf(r.RangeAddress.FirstAddress) < axis.CrossOf(first) ||
                        axis.CrossOf(r.RangeAddress.LastAddress) > axis.CrossOf(last))
            .ToList();
        foreach (var rangeToSplit in rangesToSplit)
            worksheet.MergedRanges.Remove(rangeToSplit);
    }

    private void ShiftPageBreaks<TAxis>(XLRange range, int shift)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        var breaks = axis.PageBreaks(worksheet);
        var firstIndex = axis.IndexOf(range.RangeAddress.FirstAddress);
        for (var i = 0; i < breaks.Count; i++)
        {
            if (firstIndex <= breaks[i])
                breaks[i] += shift;
        }
    }

    private void ShiftConditionalFormatting<TAxis>(XLRange range, int shift)
        where TAxis : struct, IGridAxis
    {
        if (shift == 0 || !worksheet.ConditionalFormats.Any()) return;
        var affected = AffectedArea<TAxis>(range, shift);
        ShiftConditionalFormats(cf => MoveAreas<TAxis>(cf.Areas, affected, shift));
    }

    private void ShiftDataValidation<TAxis>(XLRange range, int shift)
        where TAxis : struct, IGridAxis
    {
        if (shift == 0 || !worksheet.DataValidations.Any()) return;
        var affected = AffectedArea<TAxis>(range, shift);
        ShiftDataValidations(dv => MoveAreas<TAxis>(dv.Areas, affected, shift));
    }

    private static XLAreaList MoveAreas<TAxis>(XLAreaList areas, Area affected, int shift)
        where TAxis : struct, IGridAxis
        => shift > 0
            ? default(TAxis).InsertAndShift(areas, affected)
            : default(TAxis).DeleteAndShift(areas, affected);

    /// <summary>The region a shift inserts or deletes: the range's extent on the cross axis, and
    /// <c>|shift|</c> lines from its leading edge on the shift axis, mirroring
    /// <see cref="XLRangeInsertHelper"/>. The area model handles every insert, including one at the
    /// first line — the old range-based path short-circuited there and let the blanket range shifter
    /// move the coverage instead.</summary>
    private static Area AffectedArea<TAxis>(XLRange range, int shift)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        var first = range.RangeAddress.FirstAddress;
        return new Area(
            Point.FromAddress(first),
            axis.PointAt(axis.IndexOf(first) + Math.Abs(shift) - 1,
                axis.CrossOf(range.RangeAddress.LastAddress)));
    }

    /// <summary>Applies a value-typed area transform to every conditional-format rule and writes it back
    /// with <see cref="XLConditionalFormat.SetAreas"/>. Coverage is the value-typed
    /// <see cref="XLAreaList"/>, not live repository ranges, so the transform is pure and can never alias
    /// or double-shift across overlapping coverage (ClosedXML issue #2850). A rule whose coverage
    /// transforms to nothing is removed. Mirrors <see cref="ShiftDataValidations"/>.</summary>
    private void ShiftConditionalFormats(Func<XLConditionalFormat, XLAreaList> transform)
    {
        foreach (var cf in worksheet.ConditionalFormats.OfType<XLConditionalFormat>().ToList())
        {
            var newAreas = transform(cf);
            if (newAreas.Count == 0)
                worksheet.ConditionalFormats.Remove(f => f == cf);
            else
                cf.SetAreas(newAreas);
        }
    }

    /// <summary>Applies a value-typed area transform to every data-validation rule and writes it back
    /// with <see cref="XLDataValidation.SetAreas"/>. Coverage is the area model, not live repository
    /// ranges, so the transform is pure and can never alias or double-shift (ClosedXML issue #2850), and
    /// the write-back reindexes without split-on-add — an extended range can no longer split a
    /// not-yet-shifted rule it transiently overlaps (the drop-on-insert bug). A rule whose coverage
    /// transforms to nothing is deleted. Mirrors <see cref="ShiftConditionalFormats"/>.</summary>
    private void ShiftDataValidations(Func<XLDataValidation, XLAreaList> transform)
    {
        foreach (var dv in worksheet.DataValidations.OfType<XLDataValidation>().ToList())
        {
            var newAreas = transform(dv);
            if (newAreas.Count == 0)
                worksheet.DataValidations.Delete(v => v == dv);
            else
                dv.SetAreas(newAreas);
        }
    }

    /// <summary>Shifts cell references inside data-validation criteria formulas (formula1/formula2, in
    /// <see cref="IXLDataValidation.MinValue"/> / <see cref="IXLDataValidation.MaxValue"/>). The
    /// validation <em>ranges</em> (sqref) are handled by <see cref="ShiftDataValidation{TAxis}"/>; this
    /// re-points list/custom/comparison rules whose formula refers to other cells (e.g. dependent
    /// dropdowns on OFFSET/MATCH), mirroring <see cref="MoveDefinedNames{TAxis}"/>. Every worksheet is
    /// visited, not just the mutated one, so a formula elsewhere that references it is re-pointed too;
    /// <see cref="XLCellFormulaShifter"/> touches only references to the shifted sheet.</summary>
    private void ShiftDataValidationFormula<TAxis>(XLRange range, int shift)
        where TAxis : struct, IGridAxis
    {
        worksheet.Workbook.WorksheetsInternal.ForEach<XLWorksheet>(ws =>
        {
            var axis = default(TAxis);
            foreach (var dv in ws.DataValidations.ToList())
            {
                if (!string.IsNullOrEmpty(dv.MinValue))
                    dv.MinValue = axis.ShiftFormula(dv.MinValue, ws, range, shift);
                if (!string.IsNullOrEmpty(dv.MaxValue))
                    dv.MaxValue = axis.ShiftFormula(dv.MaxValue, ws, range, shift);
            }
        });
    }

    private void RemoveInvalidSparklines()
    {
        var invalidSparklines = worksheet.SparklineGroups.SelectMany(g => g)
            .Where(sl => !((XLAddress)sl.Location.Address).IsValid)
            .ToList();

        foreach (var sparkline in invalidSparklines)
        {
            worksheet.SparklineGroups.Remove(sparkline.Location);
        }
    }

    private static void MoveDefinedNames<TAxis>(XLRange range, int shift, XLDefinedNames definedNames)
        where TAxis : struct, IGridAxis
    {
        var axis = default(TAxis);
        var ws = range.Worksheet;
        foreach (var definedName in definedNames)
        {
            var sheetRefs = definedName.GetSheetReferencesList();
            if (sheetRefs.Count == 0)
                continue;
            var newRangeList = sheetRefs
                .Select(r => axis.ShiftFormula(r, ws, range, shift))
                .Where(newReference => newReference.Length > 0)
                .ToList();
            definedName.SetRefersTo(string.Join(",", newRangeList));
        }
    }
}
