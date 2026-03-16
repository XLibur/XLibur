using System.Linq;
using XLibur.Extensions;

namespace XLibur.Excel;

/// <summary>
/// Contains the conditional-format removal algorithm.
/// <see cref="XLRangeBase"/> delegates <c>RemoveConditionalFormatting</c> here.
/// </summary>
internal static class XLRangeConditionalFormatHelper
{
    internal static void RemoveConditionalFormatting(XLRangeBase range)
    {
        var affectedFormats = range.Worksheet.ConditionalFormats
            .Where(x => x.Ranges.GetIntersectedRanges(range.RangeAddress).Any())
            .ToList();

        foreach (var format in affectedFormats)
        {
            SplitFormatRanges(format, range);

            if (format.Ranges.Count == 0)
                range.Worksheet.ConditionalFormats.Remove(x => x == format);
        }
    }

    private static void SplitFormatRanges(IXLConditionalFormat format, XLRangeBase removeRange)
    {
        var cfRanges = format.Ranges.ToList();
        format.Ranges.RemoveAll();

        foreach (var cfRange in cfRanges)
        {
            if (!cfRange.Intersects(removeRange))
            {
                format.Ranges.Add(cfRange);
                continue;
            }

            AddRemainderRanges(format, cfRange, removeRange);
        }
    }

    private static void AddRemainderRanges(
        IXLConditionalFormat format,
        IXLRange cfRange,
        XLRangeBase removeRange)
    {
        var mf = removeRange.RangeAddress.FirstAddress;
        var ml = removeRange.RangeAddress.LastAddress;
        var f = cfRange.RangeAddress.FirstAddress;
        var l = cfRange.RangeAddress.LastAddress;

        var splitByWidth = TrySplitByWidth(format, removeRange.Worksheet, mf, ml, f, l);
        var splitByHeight = TrySplitByHeight(format, removeRange.Worksheet, mf, ml, f, l);

        if (!splitByWidth && !splitByHeight)
            format.Ranges.Add(cfRange); // Not split, preserve original
    }

    private static bool TrySplitByWidth(
        IXLConditionalFormat format,
        IXLWorksheet worksheet,
        IXLAddress mf, IXLAddress ml,
        IXLAddress f, IXLAddress l)
    {
        if (mf.ColumnNumber > f.ColumnNumber || ml.ColumnNumber < l.ColumnNumber)
            return false;

        if (!mf.RowNumber.Between(f.RowNumber, l.RowNumber) && !ml.RowNumber.Between(f.RowNumber, l.RowNumber))
            return true; // Spans full width, but no row overlap produces remainder — still counts as "by width"

        if (mf.RowNumber > f.RowNumber)
            format.Ranges.Add(worksheet.Range(f.RowNumber, f.ColumnNumber, mf.RowNumber - 1, l.ColumnNumber));

        if (ml.RowNumber < l.RowNumber)
            format.Ranges.Add(worksheet.Range(ml.RowNumber + 1, f.ColumnNumber, l.RowNumber, l.ColumnNumber));

        return true;
    }

    private static bool TrySplitByHeight(
        IXLConditionalFormat format,
        IXLWorksheet worksheet,
        IXLAddress mf, IXLAddress ml,
        IXLAddress f, IXLAddress l)
    {
        if (mf.RowNumber > f.RowNumber || ml.RowNumber < l.RowNumber)
            return false;

        if (!mf.ColumnNumber.Between(f.ColumnNumber, l.ColumnNumber) && !ml.ColumnNumber.Between(f.ColumnNumber, l.ColumnNumber))
            return true; // Spans full height but no column overlap produces remainder

        if (mf.ColumnNumber > f.ColumnNumber)
            format.Ranges.Add(worksheet.Range(f.RowNumber, f.ColumnNumber, l.RowNumber, mf.ColumnNumber - 1));

        if (ml.ColumnNumber < l.ColumnNumber)
        {
            format.Ranges.Add(worksheet.Range(f.RowNumber, ml.ColumnNumber + 1, l.RowNumber, l.ColumnNumber));
        }

        return true;
    }
}
