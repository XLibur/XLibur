namespace ClosedXML.Excel;

internal sealed class XLCustomFilteredColumn : IXLCustomFilteredColumn
{
    private readonly XLFilterColumn _filterColumn;
    private readonly XLConnector _connector;

    public XLCustomFilteredColumn(XLFilterColumn filterColumn, XLConnector connector)
    {
        _filterColumn = filterColumn;
        _connector = connector;
    }

    public void EqualTo(XLCellValue value, bool reapply)
    {
        ApplyCustomFilter(value, XLFilterOperator.Equal, reapply);
    }

    public void NotEqualTo(XLCellValue value, bool reapply)
    {
        ApplyCustomFilter(value, XLFilterOperator.NotEqual, reapply);
    }

    public void GreaterThan(XLCellValue value, bool reapply)
    {
        ApplyCustomFilter(value, XLFilterOperator.GreaterThan, reapply);
    }

    public void LessThan(XLCellValue value, bool reapply)
    {
        ApplyCustomFilter(value, XLFilterOperator.LessThan, reapply);
    }

    public void EqualOrGreaterThan(XLCellValue value, bool reapply)
    {
        ApplyCustomFilter(value, XLFilterOperator.EqualOrGreaterThan, reapply);
    }

    public void EqualOrLessThan(XLCellValue value, bool reapply)
    {
        ApplyCustomFilter(value, XLFilterOperator.EqualOrLessThan, reapply);
    }

    public void BeginsWith(string value, bool reapply)
    {
        ApplyWildcardCustomFilter(value + "*", true, reapply);
    }

    public void NotBeginsWith(string value, bool reapply)
    {
        ApplyWildcardCustomFilter(value + "*", false, reapply);
    }

    public void EndsWith(string value, bool reapply)
    {
        ApplyWildcardCustomFilter("*" + value, true, reapply);
    }

    public void NotEndsWith(string value, bool reapply)
    {
        ApplyWildcardCustomFilter("*" + value, false, reapply);
    }

    public void Contains(string value, bool reapply)
    {
        ApplyWildcardCustomFilter("*" + value + "*", true, reapply);
    }

    public void NotContains(string value, bool reapply)
    {
        ApplyWildcardCustomFilter("*" + value + "*", false, reapply);
    }

    private void ApplyCustomFilter(XLCellValue value, XLFilterOperator op, bool reapply)
    {
        _filterColumn.AddFilter(XLFilter.CreateCustomFilter(value, op, _connector), reapply);
    }

    private void ApplyWildcardCustomFilter(string pattern, bool match, bool reapply)
    {
        _filterColumn.AddFilter(XLFilter.CreateCustomPatternFilter(pattern, match, _connector), reapply);
    }
}
