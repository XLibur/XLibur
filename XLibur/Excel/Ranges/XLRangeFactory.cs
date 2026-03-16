using System;
using XLibur.Excel.Tables;

namespace XLibur.Excel;

internal sealed class XLRangeFactory
{
    #region Properties

    public XLWorksheet Worksheet { get; private set; }

    #endregion Properties

    #region Constructors

    public XLRangeFactory(XLWorksheet worksheet)
    {
        Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }

    #endregion Constructors

    #region Methods

    public XLRangeBase Create(XLRangeKey key) => key.RangeType switch
    {
        XLRangeType.Range => CreateRange(key.RangeAddress),
        XLRangeType.Column => CreateColumn(key.RangeAddress.FirstAddress.ColumnNumber),
        XLRangeType.Row => CreateColumn(key.RangeAddress.FirstAddress.RowNumber),
        XLRangeType.RangeColumn => CreateRangeColumn(key.RangeAddress),
        XLRangeType.RangeRow => CreateRangeRow(key.RangeAddress),
        XLRangeType.Table => CreateTable(key.RangeAddress),
        _ => throw new NotImplementedException(key.RangeType.ToString()),
    };

    public XLRange CreateRange(XLRangeAddress rangeAddress)
    {
        var xlRangeParameters = new XLRangeParameters(rangeAddress, Worksheet.Style);
        return new XLRange(xlRangeParameters);
    }

    public XLColumn CreateColumn(int columnNumber)
    {
        return new XLColumn(Worksheet, columnNumber);
    }

    public XLRow CreateRow(int rowNumber)
    {
        return new XLRow(Worksheet, rowNumber);
    }

    public XLRangeColumn CreateRangeColumn(XLRangeAddress rangeAddress)
    {
        var xlRangeParameters = new XLRangeParameters(rangeAddress, Worksheet.Style);
        return new XLRangeColumn(xlRangeParameters);
    }

    public XLRangeRow CreateRangeRow(XLRangeAddress rangeAddress)
    {
        var xlRangeParameters = new XLRangeParameters(rangeAddress, Worksheet.Style);
        return new XLRangeRow(xlRangeParameters);
    }

    public XLTable CreateTable(XLRangeAddress rangeAddress)
    {
        return new XLTable(new XLRangeParameters(rangeAddress, Worksheet.Style));
    }

    #endregion Methods
}
