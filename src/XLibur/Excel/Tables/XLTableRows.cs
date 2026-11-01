using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XLibur.Extensions;

namespace XLibur.Excel;

internal sealed class XLTableRows : XLStylizedBase, IXLTableRows, IXLStylized
{
    private readonly List<XLTableRow> _ranges = new List<XLTableRow>();

    public XLTableRows(IXLStyle defaultStyle) : base(((XLStyle)defaultStyle).Value)
    {
    }

    #region IXLStylized Members

    protected override IEnumerable<XLStylizedBase> Children
    {
        get
        {
            foreach (var range in _ranges)
            {
                yield return range;
            }
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

    #region IXLTableRows Members

    public IXLTableRows Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        _ranges.ForEach(r => r.Clear(clearOptions));
        return this;
    }

    public void Delete()
    {
        _ranges.OrderByDescending(r => r.RowNumber()).ForEach(r => r.Delete());
        _ranges.Clear();
    }

    public void Add(IXLTableRow tableRow)
    {
        _ranges.Add((XLTableRow)tableRow);
    }

    public IEnumerator<IXLTableRow> GetEnumerator()
    {
        var retList = new List<IXLTableRow>();
        _ranges.ForEach(retList.Add);
        return retList.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public IXLCells Cells() => XLCells.FromRanges(_ranges, false, XLCellsUsedOptions.AllContents);

    public IXLCells CellsUsed() => CellsUsed(XLCellsUsedOptions.AllContents);

    public IXLCells CellsUsed(bool includeFormats)
    {
        return CellsUsed(includeFormats
            ? XLCellsUsedOptions.All
            : XLCellsUsedOptions.AllContents);
    }

    public IXLCells CellsUsed(XLCellsUsedOptions options) => XLCells.FromRanges(_ranges, true, options);

    #endregion IXLTableRows Members

    public void Select()
    {
        foreach (var range in this)
            range.Select();
    }
}
