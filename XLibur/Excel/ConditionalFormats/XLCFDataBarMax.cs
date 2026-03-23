namespace XLibur.Excel.ConditionalFormats;

internal sealed class XLCFDataBarMax : IXLCFDataBarMax
{
    private readonly XLConditionalFormat _conditionalFormat;
    public XLCFDataBarMax(XLConditionalFormat conditionalFormat)
    {
        _conditionalFormat = conditionalFormat;
    }

    public IXLConditionalFormat Maximum(XLCFContentType type, string value)
    {
        _conditionalFormat.ContentTypes.Add(type);
        _conditionalFormat.Values.Add(new XLFormula { Value = value });
        return _conditionalFormat;
    }
    public IXLConditionalFormat Maximum(XLCFContentType type, double value)
    {
        return Maximum(type, value.ToInvariantString());
    }

    public IXLConditionalFormat HighestValue()
    {
        return Maximum(XLCFContentType.Maximum, "0");
    }
}
