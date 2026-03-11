using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel;

/// <summary>
/// Contains the conditional-format removal algorithm.
/// <see cref="XLRangeBase"/> delegates <c>RemoveConditionalFormatting</c> here.
/// </summary>
internal static class XLRangeConditionalFormatHelper
{
    internal static void RemoveConditionalFormatting(XLRangeBase range)
    {
        var mf = range.RangeAddress.FirstAddress;
        var ml = range.RangeAddress.LastAddress;
        foreach (var format in range.Worksheet.ConditionalFormats.Where(x => x.Ranges.GetIntersectedRanges(range.RangeAddress).Any()).ToList())
        {
            var cfRanges = format.Ranges.ToList();
            format.Ranges.RemoveAll();

            foreach (var cfRange in cfRanges)
            {
                if (!cfRange.Intersects(range))
                {
                    format.Ranges.Add(cfRange);
                    continue;
                }

                var f = cfRange.RangeAddress.FirstAddress;
                var l = cfRange.RangeAddress.LastAddress;
                bool byWidth = false, byHeight = false;
                XLRange? rng1 = null, rng2 = null;
                if (mf.ColumnNumber <= f.ColumnNumber && ml.ColumnNumber >= l.ColumnNumber)
                {
                    if (mf.RowNumber.Between(f.RowNumber, l.RowNumber) || ml.RowNumber.Between(f.RowNumber, l.RowNumber))
                    {
                        if (mf.RowNumber > f.RowNumber)
                            rng1 = range.Worksheet.Range(f.RowNumber, f.ColumnNumber, mf.RowNumber - 1, l.ColumnNumber);
                        if (ml.RowNumber < l.RowNumber)
                            rng2 = range.Worksheet.Range(ml.RowNumber + 1, f.ColumnNumber, l.RowNumber, l.ColumnNumber);
                    }
                    byWidth = true;
                }

                if (mf.RowNumber <= f.RowNumber && ml.RowNumber >= l.RowNumber)
                {
                    if (mf.ColumnNumber.Between(f.ColumnNumber, l.ColumnNumber) || ml.ColumnNumber.Between(f.ColumnNumber, l.ColumnNumber))
                    {
                        if (mf.ColumnNumber > f.ColumnNumber)
                            rng1 = range.Worksheet.Range(f.RowNumber, f.ColumnNumber, l.RowNumber, mf.ColumnNumber - 1);
                        if (ml.ColumnNumber < l.ColumnNumber)
                            rng2 = range.Worksheet.Range(f.RowNumber, ml.ColumnNumber + 1, l.RowNumber, l.ColumnNumber);
                    }
                    byHeight = true;
                }

                if (rng1 != null)
                {
                    format.Ranges.Add(rng1);
                }
                if (rng2 != null)
                {
                    //TODO: reflect the formula for a new range
                    format.Ranges.Add(rng2);
                }

                if (!byWidth && !byHeight)
                    format.Ranges.Add(cfRange); // Not split, preserve original
            }
            if (!format.Ranges.Any())
                range.Worksheet.ConditionalFormats.Remove(x => x == format);
        }
    }
}
