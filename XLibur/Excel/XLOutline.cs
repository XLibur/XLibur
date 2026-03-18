using System.Diagnostics.CodeAnalysis;

namespace XLibur.Excel;

internal sealed class XLOutline : IXLOutline
{
    [SetsRequiredMembers]
    public XLOutline()
    {
        SummaryVLocation = XLOutlineSummaryVLocation.Top;
        SummaryHLocation = XLOutlineSummaryHLocation.Left;
    }
    [SetsRequiredMembers]
    public XLOutline(IXLOutline outline)
    {
        SummaryHLocation = outline.SummaryHLocation;
        SummaryVLocation = outline.SummaryVLocation;
    }

    public required XLOutlineSummaryVLocation SummaryVLocation { get; set; }

    public required XLOutlineSummaryHLocation SummaryHLocation { get; set; }
}
