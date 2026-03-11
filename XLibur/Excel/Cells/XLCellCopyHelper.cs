using System;
using System.Linq;

namespace ClosedXML.Excel;

internal static class XLCellCopyHelper
{
    internal static void CopyValues(XLCell target, XLCell source)
    {
        // Rich text is basically a superset of a value. Setting a value would override rich text and vice versa.
        var sourceRichText = source.SliceRichText;
        if (sourceRichText is null)
            target.SliceCellValue = source.SliceCellValue;
        else
            target.SliceRichText = sourceRichText;

        target.FormulaR1C1 = source.FormulaR1C1;
        target.SliceComment = source.SliceComment == null
            ? null
            : new XLComment(target, source.SliceComment, source.Style.Font, source.SliceComment.Style);

        if (source.Worksheet.Hyperlinks.TryGet(source.SheetPoint, out var sourceHyperlink))
        {
            target.SetCellHyperlink(new XLHyperlink(sourceHyperlink));
        }
    }

    internal static IXLCell CopyFromInternal(XLCell target, XLCell otherCell, XLCellCopyOptions options)
    {
        if (options.HasFlag(XLCellCopyOptions.Values))
            CopyValues(target, otherCell);

        if (options.HasFlag(XLCellCopyOptions.Styles))
            target.InnerStyle = otherCell.InnerStyle;

        if (options.HasFlag(XLCellCopyOptions.Sparklines))
            CopySparklineFrom(target, otherCell);

        if (options.HasFlag(XLCellCopyOptions.ConditionalFormats))
            CopyConditionalFormatsFrom(target, otherCell);

        if (options.HasFlag(XLCellCopyOptions.DataValidations))
            CopyDataValidationFrom(target, otherCell);

        return target;
    }

    internal static IXLCell CopyFromRange(XLCell target, IXLRangeBase rangeObject)
    {
        if (rangeObject is null)
            throw new ArgumentNullException(nameof(rangeObject));

        var asRange = (XLRangeBase)rangeObject;
        var maxRows = asRange.RowCount();
        var maxColumns = asRange.ColumnCount();

        var targetRow = target.SheetPoint.Row;
        var targetCol = target.SheetPoint.Column;

        var lastRow = Math.Min(targetRow + maxRows - 1, XLHelper.MaxRowNumber);
        var lastColumn = Math.Min(targetCol + maxColumns - 1, XLHelper.MaxColumnNumber);

        var targetRange = target.Worksheet.Range(targetRow, targetCol, lastRow, lastColumn);

        if (!(asRange is XLRow || asRange is XLColumn))
        {
            targetRange.Clear();
        }

        var minRow = asRange.RangeAddress.FirstAddress.RowNumber;
        var minColumn = asRange.RangeAddress.FirstAddress.ColumnNumber;
        var cellsUsed = asRange.CellsUsed(XLCellsUsedOptions.All
                                          & ~XLCellsUsedOptions.ConditionalFormats
                                          & ~XLCellsUsedOptions.DataValidation
                                          & ~XLCellsUsedOptions.MergedRanges);
        foreach (var sourceCell in cellsUsed)
        {
            CopyFromInternal(
                target.Worksheet.Cell(
                    targetRow + sourceCell.Address.RowNumber - minRow,
                    targetCol + sourceCell.Address.ColumnNumber - minColumn
                ),
                (XLCell)sourceCell,
                XLCellCopyOptions.All
                & ~XLCellCopyOptions.ConditionalFormats
                & ~XLCellCopyOptions.DataValidations); // Conditional formats and data validation are copied separately
        }

        var rangesToMerge = asRange.Worksheet.Internals.MergedRanges
            .Where(asRange.Contains)
            .Select(IXLRange (mr) =>
            {
                var firstRow = targetRow +
                               (mr.RangeAddress.FirstAddress.RowNumber - asRange.RangeAddress.FirstAddress.RowNumber);
                var firstColumn = targetCol + (mr.RangeAddress.FirstAddress.ColumnNumber -
                                               asRange.RangeAddress.FirstAddress.ColumnNumber);
                return target.Worksheet.Range
                (
                    firstRow,
                    firstColumn,
                    firstRow + mr.RowCount() - 1,
                    firstColumn + mr.ColumnCount() - 1
                );
            })
            .ToList();

        rangesToMerge.ForEach(r => r.Merge(false));

        var dataValidations = asRange.Worksheet.DataValidations
            .GetAllInRange(asRange.RangeAddress)
            .ToList();

        foreach (var dataValidation in dataValidations)
        {
            XLDataValidation? newDataValidation = null;
            foreach (var dvRange in dataValidation.Ranges.Where(r => r.Intersects(asRange)))
            {
                var dvTargetAddress = dvRange.RangeAddress.Relative(asRange.RangeAddress, targetRange.RangeAddress);
                var dvTargetRange = target.Worksheet.Range(dvTargetAddress);
                if (newDataValidation == null)
                {
                    newDataValidation = (XLDataValidation)dvTargetRange.CreateDataValidation();
                    newDataValidation.CopyFrom(dataValidation);
                }
                else
                    newDataValidation.AddRange(dvTargetRange);
            }
        }

        CopyConditionalFormatsFrom(target, asRange);
        return target;
    }

    internal static void CopyConditionalFormatsFrom(XLCell target, XLCell otherCell)
    {
        var conditionalFormats = otherCell
            .Worksheet
            .ConditionalFormats
            .Where(c => c.Ranges.GetIntersectedRanges(otherCell).Any())
            .ToList();

        foreach (var cf in conditionalFormats)
        {
            if (otherCell.Worksheet == target.Worksheet)
            {
                if (!cf.Ranges.GetIntersectedRanges(target).Any())
                {
                    cf.Ranges.Add(target);
                }
            }
            else
            {
                CopyConditionalFormatsFrom(target, otherCell.AsRange());
            }
        }
    }

    internal static void CopyConditionalFormatsFrom(XLCell target, XLRangeBase fromRange)
    {
        var srcSheet = fromRange.Worksheet;
        var targetRow = target.SheetPoint.Row;
        var targetCol = target.SheetPoint.Column;
        var minRo = fromRange.RangeAddress.FirstAddress.RowNumber;
        var minCo = fromRange.RangeAddress.FirstAddress.ColumnNumber;
        if (srcSheet.ConditionalFormats.Any(r => r.Ranges.GetIntersectedRanges(fromRange.RangeAddress).Any()))
        {
            var fs = srcSheet.ConditionalFormats
                .SelectMany(cf => cf.Ranges.GetIntersectedRanges(fromRange.RangeAddress)).ToArray();
            if (fs.Any())
            {
                minRo = fs.Max(r => r.RangeAddress.LastAddress.RowNumber);
                minCo = fs.Max(r => r.RangeAddress.LastAddress.ColumnNumber);
            }
        }

        var rCnt = minRo - fromRange.RangeAddress.FirstAddress.RowNumber + 1;
        var cCnt = minCo - fromRange.RangeAddress.FirstAddress.ColumnNumber + 1;
        rCnt = Math.Min(rCnt, fromRange.RowCount());
        cCnt = Math.Min(cCnt, fromRange.ColumnCount());
        var toRange = target.Worksheet.Range(target, target.Worksheet.Cell(targetRow + rCnt - 1, targetCol + cCnt - 1));
        var formats =
            srcSheet.ConditionalFormats.Where(f => f.Ranges.GetIntersectedRanges(fromRange.RangeAddress).Any());

        foreach (var cf in formats.ToList())
        {
            var fmtRanges = cf.Ranges
                .GetIntersectedRanges(fromRange.RangeAddress)
                .Select(r => (XLRange)r.RangeAddress.Intersection(fromRange.RangeAddress)
                    .Relative(fromRange.RangeAddress, toRange.RangeAddress).AsRange()!)
                .ToList();

            var c = new XLConditionalFormat(fmtRanges, true);
            c.CopyFrom(cf);
            c.AdjustFormulas((XLCell)cf.Ranges.First().FirstCell(), fmtRanges.First().FirstCell());

            target.Worksheet.ConditionalFormats.Add(c);
        }
    }

    internal static void ClearMerged(XLCell cell)
    {
        var mergeToDelete = cell.Worksheet.Internals.MergedRanges.GetIntersectedRanges(cell.Address).ToList();

        mergeToDelete.ForEach(m => cell.Worksheet.Internals.MergedRanges.Remove(m));
    }

    internal static IXLCell GetTargetCell(string target, XLWorksheet defaultWorksheet)
    {
        var pair = target.Split('!');
        if (pair.Length == 1)
            return defaultWorksheet.Cell(target)!;

        var wsName = pair[0];
        if (wsName.StartsWith("'"))
            wsName = wsName.Substring(1, wsName.Length - 2);
        return defaultWorksheet.Workbook.Worksheet(wsName).Cell(pair[1]);
    }

    internal static void CopySparklineFrom(XLCell target, XLCell otherCell)
    {
        if (!otherCell.HasSparkline) return;

        var sparkline = otherCell.Sparkline!;
        var sourceDataAddress = sparkline.SourceData.RangeAddress.ToString()!;
        var shiftedRangeAddress = target.GetFormulaA1(otherCell.GetFormulaR1C1(sourceDataAddress));
        var sourceDataWorksheet = otherCell.Worksheet == sparkline.SourceData.Worksheet
            ? target.Worksheet
            : (XLWorksheet)sparkline.SourceData.Worksheet;
        var sourceData = sourceDataWorksheet.Range(shiftedRangeAddress)!;

        IXLSparklineGroup group;
        if (otherCell.Worksheet == target.Worksheet)
        {
            group = sparkline.SparklineGroup;
        }
        else
        {
            group = target.Worksheet.SparklineGroups.Add(new XLSparklineGroup(target.Worksheet, sparkline.SparklineGroup));
            if (sparkline.SparklineGroup.DateRange != null)
            {
                var dateRangeWorksheet =
                    otherCell.Worksheet == sparkline.SparklineGroup.DateRange.Worksheet
                        ? target.Worksheet
                        : sparkline.SparklineGroup.DateRange.Worksheet;
                var dateRangeAddress = sparkline.SparklineGroup.DateRange.RangeAddress.ToString()!;
                var shiftedDateRangeAddress = target.GetFormulaA1(otherCell.GetFormulaR1C1(dateRangeAddress));
                group.SetDateRange(dateRangeWorksheet.Range(shiftedDateRangeAddress));
            }
        }

        group.Add(target, sourceData);
    }

    internal static void CopyDataValidationFrom(XLCell target, XLCell otherCell)
    {
        if (otherCell.HasDataValidation)
            CopyDataValidation(target, otherCell, otherCell.GetDataValidation());
        else if (target.HasDataValidation)
        {
            target.Worksheet.DataValidations.Delete(target.AsRange());
        }
    }

    internal static void CopyDataValidation(XLCell target, XLCell otherCell, IXLDataValidation otherDv)
    {
        var thisDv = (XLDataValidation)target.GetDataValidation();
        thisDv.CopyFrom(otherDv);
        thisDv.Value = target.GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.Value));
        thisDv.MinValue = target.GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.MinValue));
        thisDv.MaxValue = target.GetFormulaA1(otherCell.GetFormulaR1C1(otherDv.MaxValue));
    }
}
