using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.ConditionalFormats;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel;

/// <summary>
/// Handles range shifting operations (insert/delete rows/columns) for a worksheet,
/// including updating conditional formats, data validations, page breaks, defined names,
/// sparklines, and calc engine notifications.
/// </summary>
internal sealed class XLWorksheetRangeShifter(XLWorksheet worksheet)
{
    public void ShiftColumns(XLRange range, int columnsShifted)
    {
        if (!range.IsEntireColumn())
        {
            var model = new XLRangeAddress(
                range.RangeAddress.FirstAddress,
                new XLAddress(range.RangeAddress.LastAddress.RowNumber, XLHelper.MaxColumnNumber, false, false));
            var rangesToSplit = worksheet.MergedRanges
                .GetIntersectedRanges(model)
                .Where(r => r.RangeAddress.FirstAddress.RowNumber < range.RangeAddress.FirstAddress.RowNumber ||
                            r.RangeAddress.LastAddress.RowNumber > range.RangeAddress.LastAddress.RowNumber)
                .ToList();
            foreach (var rangeToSplit in rangesToSplit)
            {
                worksheet.MergedRanges.Remove(rangeToSplit);
            }
        }

        worksheet.Workbook.WorksheetsInternal.ForEach<XLWorksheet>(ws => MoveDefinedNamesColumns(range, columnsShifted, ws.DefinedNames));
        MoveDefinedNamesColumns(range, columnsShifted, worksheet.Workbook.DefinedNamesInternal);
        ShiftConditionalFormattingColumns(range, columnsShifted);
        ShiftDataValidationColumns(range, columnsShifted);
        ShiftDataValidationFormulaColumns(range, columnsShifted);
        ShiftPageBreaksColumns(range, columnsShifted);
        RemoveInvalidSparklines();

        ISheetListener hyperlinks = worksheet.Hyperlinks;
        if (columnsShifted > 0)
        {
            var area = XLSheetRange
                .FromRangeAddress(range.RangeAddress)
                .ExtendRight(columnsShifted - 1);
            worksheet.Workbook.CalcEngine.OnInsertAreaAndShiftRight(range.Worksheet, area);
            hyperlinks.OnInsertAreaAndShiftRight(range.Worksheet, area);
        }
        else if (columnsShifted < 0)
        {
            var area = XLSheetRange.FromRangeAddress(range.RangeAddress);
            worksheet.Workbook.CalcEngine.OnDeleteAreaAndShiftLeft(range.Worksheet, area);
            hyperlinks.OnDeleteAreaAndShiftLeft(range.Worksheet, area);
        }
    }

    public void ShiftRows(XLRange range, int rowsShifted)
    {
        if (!range.IsEntireRow())
        {
            var model = new XLRangeAddress(
                range.RangeAddress.FirstAddress,
                new XLAddress(XLHelper.MaxRowNumber, range.RangeAddress.LastAddress.ColumnNumber, false, false));
            var rangesToSplit = worksheet.MergedRanges
                .GetIntersectedRanges(model)
                .Where(r => r.RangeAddress.FirstAddress.ColumnNumber < range.RangeAddress.FirstAddress.ColumnNumber ||
                            r.RangeAddress.LastAddress.ColumnNumber > range.RangeAddress.LastAddress.ColumnNumber)
                .ToList();
            foreach (var rangeToSplit in rangesToSplit)
            {
                worksheet.MergedRanges.Remove(rangeToSplit);
            }
        }

        worksheet.Workbook.WorksheetsInternal.ForEach<XLWorksheet>(ws => MoveDefinedNamesRows(range, rowsShifted, ws.DefinedNames));
        MoveDefinedNamesRows(range, rowsShifted, worksheet.Workbook.DefinedNamesInternal);
        ShiftConditionalFormattingRows(range, rowsShifted);
        ShiftDataValidationRows(range, rowsShifted);
        ShiftDataValidationFormulaRows(range, rowsShifted);
        RemoveInvalidSparklines();
        ShiftPageBreaksRows(range, rowsShifted);

        ISheetListener hyperlinks = worksheet.Hyperlinks;
        if (rowsShifted > 0)
        {
            var area = XLSheetRange
                .FromRangeAddress(range.RangeAddress)
                .ExtendBelow(rowsShifted - 1);
            worksheet.Workbook.CalcEngine.OnInsertAreaAndShiftDown(range.Worksheet, area);
            hyperlinks.OnInsertAreaAndShiftDown(range.Worksheet, area);
        }
        else if (rowsShifted < 0)
        {
            var area = XLSheetRange.FromRangeAddress(range.RangeAddress);
            worksheet.Workbook.CalcEngine.OnDeleteAreaAndShiftUp(range.Worksheet, area);
            hyperlinks.OnDeleteAreaAndShiftUp(range.Worksheet, area);
        }
    }

    private void ShiftPageBreaksColumns(XLRange range, int columnsShifted)
    {
        for (var i = 0; i < worksheet.PageSetup.ColumnBreaks.Count; i++)
        {
            var br = worksheet.PageSetup.ColumnBreaks[i];
            if (range.RangeAddress.FirstAddress.ColumnNumber <= br)
            {
                worksheet.PageSetup.ColumnBreaks[i] = br + columnsShifted;
            }
        }
    }

    private void ShiftPageBreaksRows(XLRange range, int rowsShifted)
    {
        for (var i = 0; i < worksheet.PageSetup.RowBreaks.Count; i++)
        {
            var br = worksheet.PageSetup.RowBreaks[i];
            if (range.RangeAddress.FirstAddress.RowNumber <= br)
            {
                worksheet.PageSetup.RowBreaks[i] = br + rowsShifted;
            }
        }
    }

    private void ShiftConditionalFormattingColumns(XLRange range, int columnsShifted)
    {
        if (columnsShifted == 0 || !worksheet.ConditionalFormats.Any()) return;
        var first = range.RangeAddress.FirstAddress;
        var last = range.RangeAddress.LastAddress;
        // The affected region spans the range's rows and the inserted/deleted columns.
        var affected = columnsShifted > 0
            ? new XLSheetRange(first.RowNumber, first.ColumnNumber, last.RowNumber, first.ColumnNumber + columnsShifted - 1)
            : new XLSheetRange(first.RowNumber, first.ColumnNumber, last.RowNumber, first.ColumnNumber - columnsShifted - 1);

        ShiftConditionalFormats(cf =>
            columnsShifted > 0 ? cf.Areas.InsertAndShiftRight(affected) : cf.Areas.DeleteAndShiftLeft(affected));
    }

    private void ShiftConditionalFormattingRows(XLRange range, int rowsShifted)
    {
        if (rowsShifted == 0 || !worksheet.ConditionalFormats.Any()) return;
        var first = range.RangeAddress.FirstAddress;
        var last = range.RangeAddress.LastAddress;
        // The affected region spans the range's columns and the inserted/deleted rows, mirroring
        // the inserted range in XLRangeInsertHelper. Structural shifts run on the value-typed
        // XLAreaList so overlapping/adjacent coverage can never alias or double-shift (issue #2850).
        var affected = rowsShifted > 0
            ? new XLSheetRange(first.RowNumber, first.ColumnNumber, first.RowNumber + rowsShifted - 1, last.ColumnNumber)
            : new XLSheetRange(first.RowNumber, first.ColumnNumber, first.RowNumber - rowsShifted - 1, last.ColumnNumber);

        ShiftConditionalFormats(cf =>
            rowsShifted > 0 ? cf.Areas.InsertAndShiftDown(affected) : cf.Areas.DeleteAndShiftUp(affected));
    }

    /// <summary>
    /// Applies a value-typed area transform to every conditional-format rule and writes the result back
    /// with <see cref="XLConditionalFormat.SetAreas"/>. Because coverage is the value-typed
    /// <see cref="XLAreaList"/> (not live repository ranges) the shift is a pure transform that can never
    /// alias or double-shift, even across overlapping/adjacent coverage (ClosedXML issue #2850). A rule
    /// whose coverage transforms to nothing is removed. Mirrors <see cref="ShiftDataValidations"/>.
    /// </summary>
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

    private void ShiftDataValidationColumns(XLRange range, int columnsShifted)
    {
        if (columnsShifted == 0 || !worksheet.DataValidations.Any()) return;
        var first = range.RangeAddress.FirstAddress;
        var last = range.RangeAddress.LastAddress;
        // The area model handles every insert, including a first-column insert. (The old range-based
        // path short-circuited here and let the blanket range shifter move the validation ranges;
        // area-based coverage is no longer a repository range, so it must be shifted here.)
        var affected = columnsShifted > 0
            ? new XLSheetRange(first.RowNumber, first.ColumnNumber, last.RowNumber, first.ColumnNumber + columnsShifted - 1)
            : new XLSheetRange(first.RowNumber, first.ColumnNumber, last.RowNumber, first.ColumnNumber - columnsShifted - 1);

        ShiftDataValidations(dv =>
            columnsShifted > 0 ? dv.Areas.InsertAndShiftRight(affected) : dv.Areas.DeleteAndShiftLeft(affected));
    }

    private void ShiftDataValidationRows(XLRange range, int rowsShifted)
    {
        if (rowsShifted == 0 || !worksheet.DataValidations.Any()) return;
        var first = range.RangeAddress.FirstAddress;
        var last = range.RangeAddress.LastAddress;
        // The area model handles every insert, including a first-row insert. (The old range-based
        // path short-circuited here and let the blanket range shifter move the validation ranges;
        // area-based coverage is no longer a repository range, so it must be shifted here.)
        var affected = rowsShifted > 0
            ? new XLSheetRange(first.RowNumber, first.ColumnNumber, first.RowNumber + rowsShifted - 1, last.ColumnNumber)
            : new XLSheetRange(first.RowNumber, first.ColumnNumber, first.RowNumber - rowsShifted - 1, last.ColumnNumber);

        ShiftDataValidations(dv =>
            rowsShifted > 0 ? dv.Areas.InsertAndShiftDown(affected) : dv.Areas.DeleteAndShiftUp(affected));
    }

    /// <summary>
    /// Applies a value-typed area transform to every data-validation rule and writes the result back
    /// with <see cref="XLDataValidation.SetAreas"/>. Because coverage is the area model (not live
    /// repository ranges) the shift is a pure transform that can never alias or double-shift (ClosedXML
    /// issue #2850), and the write-back reindexes without split-on-add — so an extended range can no
    /// longer split a not-yet-shifted rule it transiently overlaps (the drop-on-insert bug). A rule
    /// whose coverage transforms to nothing is deleted. Mirrors ShiftConditionalFormatting*.
    /// </summary>
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

    /// <summary>
    /// Shifts cell references inside data-validation criteria formulas (formula1/formula2,
    /// stored in <see cref="IXLDataValidation.MinValue"/> / <see cref="IXLDataValidation.MaxValue"/>)
    /// when columns are inserted or deleted. The validation <em>ranges</em> (sqref) are handled by
    /// <see cref="ShiftDataValidationColumns"/>; this re-points list/custom/comparison rules
    /// whose formula refers to other cells (e.g. dependent dropdowns built on OFFSET/MATCH),
    /// mirroring <see cref="MoveDefinedNamesColumns"/>. Every worksheet's validations are visited
    /// (not just the mutated sheet's) so a formula on one sheet that references the mutated sheet is
    /// re-pointed too; <see cref="XLCellFormulaShifter.ShiftFormulaColumns"/> only touches references
    /// to the sheet being shifted, so unrelated references pass through.
    /// </summary>
    private void ShiftDataValidationFormulaColumns(XLRange range, int columnsShifted)
    {
        worksheet.Workbook.WorksheetsInternal.ForEach<XLWorksheet>(ws =>
        {
            foreach (var dv in ws.DataValidations.ToList())
            {
                if (!string.IsNullOrEmpty(dv.MinValue))
                    dv.MinValue = XLCellFormulaShifter.ShiftFormulaColumns(dv.MinValue, ws, range, columnsShifted);
                if (!string.IsNullOrEmpty(dv.MaxValue))
                    dv.MaxValue = XLCellFormulaShifter.ShiftFormulaColumns(dv.MaxValue, ws, range, columnsShifted);
            }
        });
    }

    /// <summary>
    /// Shifts cell references inside data-validation criteria formulas (formula1/formula2,
    /// stored in <see cref="IXLDataValidation.MinValue"/> / <see cref="IXLDataValidation.MaxValue"/>)
    /// when rows are inserted or deleted. The row counterpart of
    /// <see cref="ShiftDataValidationFormulaColumns"/>.
    /// </summary>
    private void ShiftDataValidationFormulaRows(XLRange range, int rowsShifted)
    {
        worksheet.Workbook.WorksheetsInternal.ForEach<XLWorksheet>(ws =>
        {
            foreach (var dv in ws.DataValidations.ToList())
            {
                if (!string.IsNullOrEmpty(dv.MinValue))
                    dv.MinValue = XLCellFormulaShifter.ShiftFormulaRows(dv.MinValue, ws, range, rowsShifted);
                if (!string.IsNullOrEmpty(dv.MaxValue))
                    dv.MaxValue = XLCellFormulaShifter.ShiftFormulaRows(dv.MaxValue, ws, range, rowsShifted);
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

    private static void MoveDefinedNamesRows(XLRange range, int rowsShifted, XLDefinedNames definedNames)
    {
        var ws = range.Worksheet;
        foreach (var definedName in definedNames)
        {
            var sheetRefs = definedName.GetSheetReferencesList();
            if (sheetRefs.Count > 0)
            {
                var newRangeList = sheetRefs
                    .Select(r => XLCellFormulaShifter.ShiftFormulaRows(r, ws, range, rowsShifted))
                    .Where(newReference => newReference.Length > 0)
                    .ToList();
                var unionFormula = string.Join(",", newRangeList);
                definedName.SetRefersTo(unionFormula);
            }
        }
    }

    private static void MoveDefinedNamesColumns(XLRange range, int columnsShifted, XLDefinedNames definedNames)
    {
        var ws = range.Worksheet;
        foreach (var definedName in definedNames)
        {
            var sheetRefs = definedName.GetSheetReferencesList();
            if (sheetRefs.Count > 0)
            {
                var newRangeList = sheetRefs
                    .Select(r => XLCellFormulaShifter.ShiftFormulaColumns(r, ws, range, columnsShifted))
                    .Where(newReference => newReference.Length > 0)
                    .ToList();
                var unionFormula = string.Join(",", newRangeList);
                definedName.SetRefersTo(unionFormula);
            }
        }
    }
}
