using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.ConditionalFormats;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// Contains the conditional-format removal algorithm.
/// <see cref="XLRangeBase"/> delegates <c>RemoveConditionalFormatting</c> here.
/// </summary>
/// <remarks>
/// Coverage is only cut when the cleared range fully spans a format area's width or height;
/// a partial (corner) overlap preserves the area unchanged. Operates on the value-typed
/// <see cref="XLConditionalFormat.Areas"/> and writes results back via <see cref="XLConditionalFormat.SetAreas"/>.
/// </remarks>
internal static class XLRangeConditionalFormatHelper
{
    internal static void RemoveConditionalFormatting(XLRangeBase range)
    {
        var remove = Area.FromRangeAddress(range.RangeAddress);

        var affectedFormats = range.Worksheet.ConditionalFormats
            .Cast<XLConditionalFormat>()
            .Where(cf => cf.Areas.IntersectsWith(remove))
            .ToList();

        foreach (var format in affectedFormats)
        {
            var result = new List<Area>();
            foreach (var cfArea in format.Areas)
            {
                if (cfArea.Intersect(remove) is null)
                {
                    result.Add(cfArea);
                    continue;
                }

                AddRemainderAreas(result, cfArea, remove);
            }

            if (result.Count == 0)
                range.Worksheet.ConditionalFormats.Remove(x => x == format);
            else
                format.SetAreas(new XLAreaList(result));
        }
    }

    private static void AddRemainderAreas(List<Area> result, Area cf, Area remove)
    {
        var splitByWidth = TrySplitByWidth(result, remove, cf);
        var splitByHeight = TrySplitByHeight(result, remove, cf);

        if (!splitByWidth && !splitByHeight)
            result.Add(cf); // Not split, preserve original
    }

    private static bool TrySplitByWidth(List<Area> result, Area remove, Area cf)
    {
        if (remove.LeftColumn > cf.LeftColumn || remove.RightColumn < cf.RightColumn)
            return false;

        // Spans full width, but no row overlap produces a remainder — still counts as "by width".
        if (!Between(remove.TopRow, cf.TopRow, cf.BottomRow) && !Between(remove.BottomRow, cf.TopRow, cf.BottomRow))
            return true;

        if (remove.TopRow > cf.TopRow)
            result.Add(new Area(cf.TopRow, cf.LeftColumn, remove.TopRow - 1, cf.RightColumn));

        if (remove.BottomRow < cf.BottomRow)
            result.Add(new Area(remove.BottomRow + 1, cf.LeftColumn, cf.BottomRow, cf.RightColumn));

        return true;
    }

    private static bool TrySplitByHeight(List<Area> result, Area remove, Area cf)
    {
        if (remove.TopRow > cf.TopRow || remove.BottomRow < cf.BottomRow)
            return false;

        // Spans full height but no column overlap produces the remainder.
        if (!Between(remove.LeftColumn, cf.LeftColumn, cf.RightColumn) && !Between(remove.RightColumn, cf.LeftColumn, cf.RightColumn))
            return true;

        if (remove.LeftColumn > cf.LeftColumn)
            result.Add(new Area(cf.TopRow, cf.LeftColumn, cf.BottomRow, remove.LeftColumn - 1));

        if (remove.RightColumn < cf.RightColumn)
            result.Add(new Area(cf.TopRow, remove.RightColumn + 1, cf.BottomRow, cf.RightColumn));

        return true;
    }

    private static bool Between(int value, int low, int high) => value >= low && value <= high;
}
