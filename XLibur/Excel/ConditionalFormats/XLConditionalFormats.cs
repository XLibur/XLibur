using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Extensions;

namespace XLibur.Excel;

/// <summary>
/// A container for conditional formatting of a <see cref="XLWorksheet"/>. It contains
/// a collection of <see cref="XLConditionalFormat"/>. Doesn't contain pivot table formats,
/// they are in pivot table <see cref="XLPivotTable.ConditionalFormats"/>,
/// </summary>
internal sealed class XLConditionalFormats : IXLConditionalFormats
{
    private readonly List<IXLConditionalFormat> _conditionalFormats = [];

    private static readonly List<XLConditionalFormatType> CFTypesExcludedFromConsolidation =
    [
        XLConditionalFormatType.DataBar,
        XLConditionalFormatType.ColorScale,
        XLConditionalFormatType.IconSet,
        XLConditionalFormatType.Top10,
        XLConditionalFormatType.AboveAverage,
        XLConditionalFormatType.IsDuplicate,
        XLConditionalFormatType.IsUnique
    ];

    public void Add(IXLConditionalFormat conditionalFormat)
    {
        _conditionalFormats.Add(conditionalFormat);
    }

    public IEnumerator<IXLConditionalFormat> GetEnumerator()
    {
        return _conditionalFormats.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public void Remove(Predicate<IXLConditionalFormat> predicate)
    {
        _conditionalFormats.RemoveAll(predicate);
    }

    /// <summary>
    /// The method consolidate the same conditional formats, which are located in adjacent ranges.
    /// </summary>
    internal void Consolidate()
    {
        var formats = _conditionalFormats
            .Where(cf => cf.Ranges.Any())
            .ToList();
        _conditionalFormats.Clear();

        while (formats.Count > 0)
        {
            var item = formats[0];

            if (!CFTypesExcludedFromConsolidation.Contains(item.ConditionalFormatType))
            {
                var similarFormats = ConsolidateItem(item, formats);
                similarFormats.ForEach(cf => formats.Remove(cf));
            }

            _conditionalFormats.Add(item);
            formats.Remove(item);
        }
    }

    private static List<IXLConditionalFormat> ConsolidateItem(IXLConditionalFormat item, List<IXLConditionalFormat> formats)
    {
        var rangesToJoin = new XLRanges();
        item.Ranges.ForEach(rangesToJoin.Add);
        var firstRange = item.Ranges.First();
        var skippedRanges = new XLRanges();
        Func<IXLConditionalFormat, bool> IsSameFormat = f =>
            f != item && f.Ranges.First().Worksheet.Position == firstRange.Worksheet.Position &&
            XLConditionalFormat.NoRangeComparer.Equals(f, item);

        var baseAddress = new XLAddress(
            item.Ranges.Select(r => r.RangeAddress.FirstAddress.RowNumber).Min(),
            item.Ranges.Select(r => r.RangeAddress.FirstAddress.ColumnNumber).Min(),
            false, false);
        var baseCell = (XLCell)firstRange.Worksheet.Cell(baseAddress);

        var similarFormats = FindSimilarFormats(formats, rangesToJoin, skippedRanges, IsSameFormat);

        var consRanges = rangesToJoin.Consolidate();
        item.Ranges.RemoveAll();
        consRanges.ForEach(r => item.Ranges.Add(r));

        var targetCell = (XLCell)item.Ranges.First().FirstCell();
        ((XLConditionalFormat)item).AdjustFormulas(baseCell, targetCell);

        return similarFormats;
    }

    private static List<IXLConditionalFormat> FindSimilarFormats(
        List<IXLConditionalFormat> formats,
        XLRanges rangesToJoin,
        XLRanges skippedRanges,
        Func<IXLConditionalFormat, bool> isSameFormat)
    {
        List<IXLConditionalFormat> similarFormats = [];
        int i = 1;
        bool stop;
        do
        {
            stop = (i >= formats.Count);

            if (!stop)
            {
                var nextFormat = formats[i];

                var intersectsSkipped =
                    skippedRanges.Any(left => nextFormat.Ranges.GetIntersectedRanges(left.RangeAddress).Any());

                var isSame = isSameFormat(nextFormat);

                if (isSame && !intersectsSkipped)
                {
                    similarFormats.Add(nextFormat);
                    nextFormat.Ranges.ForEach(rangesToJoin.Add);
                }
                else if (rangesToJoin.Any(left => nextFormat.Ranges.GetIntersectedRanges(left.RangeAddress).Any()) ||
                         intersectsSkipped)
                {
                    stop = true;
                }

                if (!isSame)
                    nextFormat.Ranges.ForEach(skippedRanges.Add);
            }

            i++;
        } while (!stop);

        return similarFormats;
    }

    public void RemoveAll()
    {
        _conditionalFormats.Clear();
    }

    /// <summary>
    /// Reorders the according to original priority. Done during load process
    /// </summary>
    public void ReorderAccordingToOriginalPriority()
    {
        var reorderedFormats = _conditionalFormats.OrderBy(cf => ((XLConditionalFormat)cf).Priority).ToList();
        _conditionalFormats.Clear();
        _conditionalFormats.AddRange(reorderedFormats);
    }
}
