#nullable disable


namespace XLibur.Excel;

internal class XLOutline : IXLOutline
{
    public XLOutline(IXLOutline outline)
    {
        if (outline != null)
        {
            SummaryHLocation = outline.SummaryHLocation;
            SummaryVLocation = outline.SummaryVLocation;
        }
    }
    public XLOutlineSummaryVLocation SummaryVLocation { get; set; }
    public XLOutlineSummaryHLocation SummaryHLocation { get; set; }
}
