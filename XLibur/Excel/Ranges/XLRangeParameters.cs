namespace ClosedXML.Excel;

internal sealed class XLRangeParameters
{
    #region Constructor

    public XLRangeParameters(XLRangeAddress rangeAddress, IXLStyle defaultStyle)
    {
        RangeAddress = rangeAddress;
        DefaultStyle = defaultStyle;
    }

    #endregion

    #region Properties

    public XLRangeAddress RangeAddress { get; private set; }

    public IXLStyle DefaultStyle { get; private set; }
    #endregion
}
