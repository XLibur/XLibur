using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XLibur.Extensions;

namespace XLibur.Excel;

internal sealed class XLRangeRows : XLStylizedBase, IXLRangeRows, IXLStylized
{
    private readonly List<XLRangeRow> _ranges = new List<XLRangeRow>();

    public XLRangeRows() : base(XLStyle.Default.Value)
    {
    }

    #region IXLRangeRows Members

    public IXLRangeRows Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        _ranges.ForEach(c => c.Clear(clearOptions));
        return this;
    }

    public void Delete()
    {
        _ranges.OrderByDescending(r => r.RowNumber()).ForEach(r => r.Delete());
        _ranges.Clear();
    }

    public void Add(IXLRangeRow rowRange)
    {
        _ranges.Add((XLRangeRow)rowRange);
    }

    public IEnumerator<IXLRangeRow> GetEnumerator()
    {
        return _ranges.Cast<IXLRangeRow>()
            .OrderBy(r => r.Worksheet.Position)
            .ThenBy(r => r.RowNumber())
            .GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public IXLCells Cells() => XLCells.FromRanges(_ranges, false, XLCellsUsedOptions.AllContents);

    public IXLCells CellsUsed() => CellsUsed(XLCellsUsedOptions.AllContents);

    public IXLCells CellsUsed(XLCellsUsedOptions options) => XLCells.FromRanges(_ranges, true, options);

    #endregion IXLRangeRows Members

    #region IXLStylized Members

    protected override IEnumerable<XLStylizedBase> Children
    {
        get
        {
            foreach (var range in _ranges)
                yield return range;
        }
    }


    public override IXLRanges RangesUsed
    {
        get
        {
            var retVal = new XLRanges();
            this.ForEach(c => retVal.Add(c.AsRange()));
            return retVal;
        }
    }

    #endregion IXLStylized Members

    public void Select()
    {
        foreach (var range in this)
            range.Select();
    }
}
