namespace XLibur.Excel;

public interface IXLCustomFilteredColumn
{
    void EqualTo(XLCellValue value, bool reapply = true);
    void NotEqualTo(XLCellValue value, bool reapply = true);
    void GreaterThan(XLCellValue value, bool reapply = true);
    void LessThan(XLCellValue value, bool reapply = true);
    void EqualOrGreaterThan(XLCellValue value, bool reapply = true);
    void EqualOrLessThan(XLCellValue value, bool reapply = true);
    void BeginsWith(string value, bool reapply = true);
    void NotBeginsWith(string value, bool reapply = true);
    void EndsWith(string value, bool reapply = true);
    void NotEndsWith(string value, bool reapply = true);
    void Contains(string value, bool reapply = true);
    void NotContains(string value, bool reapply = true);
}
