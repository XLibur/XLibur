namespace XLibur.Excel;

public interface IXLCFDataBarMax
{
    IXLConditionalFormat Maximum(XLCFContentType type, string value);
    IXLConditionalFormat Maximum(XLCFContentType type, double value);
    IXLConditionalFormat HighestValue();
}
