namespace XLibur.Excel;

internal record XLSortElement(
    int ElementNumber,
    XLSortOrder SortOrder,
    bool IgnoreBlanks,
    bool MatchCase) : IXLSortElement;
