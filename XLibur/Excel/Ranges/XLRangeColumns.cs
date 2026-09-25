using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XLibur.Extensions;

namespace XLibur.Excel;

internal sealed class XLRangeColumns : XLStylizedBase, IXLRangeColumns, IXLStylized
{
    private readonly List<XLRangeColumn> _ranges = [];

    public XLRangeColumns() : base(XLWorkbook.DefaultStyleValue)
    {
    }

    #region IXLRangeColumns Members

    public IXLRangeColumns Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        _ranges.ForEach(c => c.Clear(clearOptions));
        return this;
    }

    public void Delete()
    {
        _ranges.OrderByDescending(c => c.ColumnNumber()).ForEach(r => r.Delete());
        _ranges.Clear();
    }

    public void Add(IXLRangeColumn columnRange)
    {
        _ranges.Add((XLRangeColumn)columnRange);
    }

    public IEnumerator<IXLRangeColumn> GetEnumerator()
    {
        return _ranges.Cast<IXLRangeColumn>()
            .OrderBy(r => r.Worksheet.Position)
            .ThenBy(r => r.ColumnNumber())
            .GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public IXLCells Cells() => XLCells.FromRanges(_ranges, false, XLCellsUsedOptions.AllContents);

    public IXLCells CellsUsed() => CellsUsed(XLCellsUsedOptions.AllContents);

    public IXLCells CellsUsed(XLCellsUsedOptions options) => XLCells.FromRanges(_ranges, true, options);

    #endregion IXLRangeColumns Members

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
