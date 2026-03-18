using XLibur.Excel.CalcEngine;
using System;
using System.Linq;
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
        if (!worksheet.ConditionalFormats.Any()) return;
        var firstCol = range.RangeAddress.FirstAddress.ColumnNumber;
        if (firstCol == 1) return;

        var colNum = columnsShifted > 0 ? firstCol - 1 : firstCol;
        var model = worksheet.Column(colNum).AsRange();

        foreach (var cf in worksheet.ConditionalFormats.ToList())
        {
            var cfRanges = cf.Ranges.ToList();
            cf.Ranges.RemoveAll();

            foreach (var cfRange in cfRanges)
            {
                var newRange = ShiftRangeColumns(cfRange, model, firstCol, columnsShifted);
                if (newRange.RangeAddress.IsValid &&
                    newRange.RangeAddress.FirstAddress.ColumnNumber <=
                    newRange.RangeAddress.LastAddress.ColumnNumber)
                    cf.Ranges.Add(newRange);
            }

            if (cf.Ranges.Count == 0)
                worksheet.ConditionalFormats.Remove(f => f == cf);
        }
    }

    private void ShiftConditionalFormattingRows(XLRange range, int rowsShifted)
    {
        if (!worksheet.ConditionalFormats.Any()) return;
        var firstRow = range.RangeAddress.FirstAddress.RowNumber;
        if (firstRow == 1) return;

        var rowNum = rowsShifted > 0 ? firstRow - 1 : firstRow;
        var model = worksheet.Row(rowNum).AsRange();

        foreach (var cf in worksheet.ConditionalFormats.ToList())
        {
            var cfRanges = cf.Ranges.ToList();
            cf.Ranges.RemoveAll();

            foreach (var cfRange in cfRanges)
            {
                var newRange = ShiftRangeRows(cfRange, model, firstRow, rowsShifted);
                if (newRange.RangeAddress.IsValid &&
                    newRange.RangeAddress.FirstAddress.RowNumber <= newRange.RangeAddress.LastAddress.RowNumber)
                    cf.Ranges.Add(newRange);
            }

            if (cf.Ranges.Count == 0)
                worksheet.ConditionalFormats.Remove(f => f == cf);
        }
    }

    private void ShiftDataValidationColumns(XLRange range, int columnsShifted)
    {
        if (!worksheet.DataValidations.Any()) return;
        var firstCol = range.RangeAddress.FirstAddress.ColumnNumber;
        if (firstCol == 1) return;

        var colNum = columnsShifted > 0 ? firstCol - 1 : firstCol;
        var model = worksheet.Column(colNum).AsRange();

        foreach (var dv in worksheet.DataValidations.ToList())
        {
            var dvRanges = dv.Ranges.ToList();
            dv.ClearRanges();

            foreach (var dvRange in dvRanges)
            {
                var newRange = ShiftRangeColumns(dvRange, model, firstCol, columnsShifted);
                if (newRange.RangeAddress.IsValid &&
                    newRange.RangeAddress.FirstAddress.ColumnNumber <=
                    newRange.RangeAddress.LastAddress.ColumnNumber)
                    dv.AddRange(newRange);
            }

            if (!dv.Ranges.Any())
                worksheet.DataValidations.Delete(v => v == dv);
        }
    }

    private void ShiftDataValidationRows(XLRange range, int rowsShifted)
    {
        if (!worksheet.DataValidations.Any()) return;
        var firstRow = range.RangeAddress.FirstAddress.RowNumber;
        if (firstRow == 1) return;

        var rowNum = rowsShifted > 0 ? firstRow - 1 : firstRow;
        var model = worksheet.Row(rowNum).AsRange();

        foreach (var dv in worksheet.DataValidations.ToList())
        {
            var dvRanges = dv.Ranges.ToList();
            dv.ClearRanges();

            foreach (var dvRange in dvRanges)
            {
                var newRange = ShiftRangeRows(dvRange, model, firstRow, rowsShifted);
                if (newRange.RangeAddress.IsValid &&
                    newRange.RangeAddress.FirstAddress.RowNumber <= newRange.RangeAddress.LastAddress.RowNumber)
                    dv.AddRange(newRange);
            }

            if (!dv.Ranges.Any())
                worksheet.DataValidations.Delete(v => v == dv);
        }
    }

    private IXLRange ShiftRangeColumns(IXLRange range, IXLRange model, int firstCol, int columnsShifted)
    {
        var address = range.RangeAddress;
        if (range.Intersects(model))
        {
            return worksheet.Range(address.FirstAddress.RowNumber,
                address.FirstAddress.ColumnNumber,
                address.LastAddress.RowNumber,
                Math.Min(XLHelper.MaxColumnNumber, address.LastAddress.ColumnNumber + columnsShifted));
        }

        if (address.FirstAddress.ColumnNumber >= firstCol)
        {
            return worksheet.Range(address.FirstAddress.RowNumber,
                Math.Max(address.FirstAddress.ColumnNumber + columnsShifted, firstCol),
                address.LastAddress.RowNumber,
                Math.Min(XLHelper.MaxColumnNumber, address.LastAddress.ColumnNumber + columnsShifted));
        }

        return range;
    }

    private IXLRange ShiftRangeRows(IXLRange range, IXLRange model, int firstRow, int rowsShifted)
    {
        var address = range.RangeAddress;
        if (range.Intersects(model))
        {
            return worksheet.Range(address.FirstAddress.RowNumber,
                address.FirstAddress.ColumnNumber,
                Math.Min(XLHelper.MaxRowNumber, address.LastAddress.RowNumber + rowsShifted),
                address.LastAddress.ColumnNumber);
        }

        if (address.FirstAddress.RowNumber >= firstRow)
        {
            return worksheet.Range(Math.Max(address.FirstAddress.RowNumber + rowsShifted, firstRow),
                address.FirstAddress.ColumnNumber,
                Math.Min(XLHelper.MaxRowNumber, address.LastAddress.RowNumber + rowsShifted),
                address.LastAddress.ColumnNumber);
        }

        return range;
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
