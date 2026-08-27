using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel.ConditionalFormats;

/// <summary>
/// A container for conditional formatting of a <see cref="XLWorksheet"/>. It contains
/// a collection of <see cref="XLConditionalFormat"/>. Doesn't contain pivot table formats,
/// they are in pivot table <see cref="XLPivotTable.ConditionalFormats"/>,
/// </summary>
internal sealed class XLConditionalFormats : IXLConditionalFormats, ISheetListener
{
    private readonly List<IXLConditionalFormat> _conditionalFormats = [];

    private readonly XLWorksheet _worksheet;

    internal XLConditionalFormats(XLWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

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

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public void Remove(Predicate<IXLConditionalFormat> predicate)
    {
        _conditionalFormats.RemoveAll(predicate);
    }

    #region ISheetListener

    /// <summary>
    /// Moves every rule's coverage over the edit, and removes a rule whose coverage transforms to
    /// nothing.
    /// </summary>
    /// <remarks>
    /// Coverage is the value-typed <see cref="XLAreaList"/>, not live repository ranges, so the
    /// transform is pure and can never alias or double-shift across overlapping coverage
    /// (ClosedXML issue #2850). Mirrors <see cref="XLDataValidations"/>, which shifts its sqref
    /// coverage the same way.
    /// </remarks>
    void ISheetListener.OnInsertAreaAndShiftDown(in SheetEdit edit) => ShiftCoverage<RowAxis>(in edit);

    void ISheetListener.OnInsertAreaAndShiftRight(in SheetEdit edit) => ShiftCoverage<ColumnAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftUp(in SheetEdit edit) => ShiftCoverage<RowAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftLeft(in SheetEdit edit) => ShiftCoverage<ColumnAxis>(in edit);

    private void ShiftCoverage<TAxis>(in SheetEdit edit)
        where TAxis : struct, IGridAxis
    {
        if (edit.Sheet != _worksheet || _conditionalFormats.Count == 0)
            return;

        // CoverageArea, not edit.Area: coverage derives the |Shift| lines the edit moves from the
        // range and the shift rather than trusting the shifter's area. See SheetEdit.
        var axis = default(TAxis);
        var affected = edit.CoverageArea<TAxis>();
        foreach (var cf in _conditionalFormats.OfType<XLConditionalFormat>().ToList())
        {
            var newAreas = edit.Shift > 0
                ? axis.InsertAndShift(cf.Areas, affected)
                : axis.DeleteAndShift(cf.Areas, affected);

            if (newAreas.Count == 0)
                Remove(f => f == cf);
            else
                cf.SetAreas(newAreas);
        }
    }

    #endregion ISheetListener

    /// <summary>
    /// The method consolidates the same conditional formats, which are located in adjacent ranges.
    /// </summary>
    internal void Consolidate()
    {
        var formats = _conditionalFormats
            .Where(cf => cf.Ranges.Count > 0)
            .ToList();
        _conditionalFormats.Clear();

        while (formats.Count > 0)
        {
            var item = formats[0];

            if (!CFTypesExcludedFromConsolidation.Contains(item.ConditionalFormatType))
            {
                var similarFormats = ConsolidateItem(item, formats);
                formats.RemoveAll(similarFormats.Contains);
            }

            _conditionalFormats.Add(item);
            formats.Remove(item);
        }
    }

    private static List<IXLConditionalFormat> ConsolidateItem(IXLConditionalFormat item,
        List<IXLConditionalFormat> formats)
    {
        var rangesToJoin = new XLRanges();
        item.Ranges.ForEach(rangesToJoin.Add);
        var firstRange = item.Ranges.First();
        var skippedRanges = new XLRanges();

        var baseAddress = new XLAddress(
            item.Ranges.Select(r => r.RangeAddress.FirstAddress.RowNumber).Min(),
            item.Ranges.Select(r => r.RangeAddress.FirstAddress.ColumnNumber).Min(),
            false, false);
        var baseCell = (XLCell)firstRange.Worksheet.Cell(baseAddress);

        var similarFormats = FindSimilarFormats(formats, rangesToJoin, skippedRanges, IsSameFormat);

        var consAreas = XLAreaList.FromRanges(rangesToJoin).GetConsolidated();
        ((XLConditionalFormat)item).SetAreas(consAreas);

        var targetCell = (XLCell)item.Ranges.First().FirstCell();
        ((XLConditionalFormat)item).AdjustFormulas(baseCell, targetCell);

        return similarFormats;

        bool IsSameFormat(IXLConditionalFormat f) => f != item &&
                                                     f.Ranges.First().Worksheet.Position ==
                                                     firstRange.Worksheet.Position &&
                                                     XLConditionalFormat.NoRangeComparer.Equals(f, item);
    }

    private static List<IXLConditionalFormat> FindSimilarFormats(
        List<IXLConditionalFormat> formats,
        XLRanges rangesToJoin,
        XLRanges skippedRanges,
        Func<IXLConditionalFormat, bool> isSameFormat)
    {
        List<IXLConditionalFormat> similarFormats = [];
        var i = 1;
        bool stop;
        do
        {
            stop = i >= formats.Count;

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
    /// Reorders the conditional formats according to the original priority. Done during the load process.
    /// </summary>
    public void ReorderAccordingToOriginalPriority()
    {
        var reorderedFormats = _conditionalFormats.OrderBy(cf => ((XLConditionalFormat)cf).Priority).ToList();
        _conditionalFormats.Clear();
        _conditionalFormats.AddRange(reorderedFormats);
    }
}
