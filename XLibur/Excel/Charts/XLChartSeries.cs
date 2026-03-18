namespace XLibur.Excel;

internal sealed class XLChartSeries : IXLChartSeries
{
    public string Name { get; set; } = string.Empty;
    public string? CategoryReferences { get; set; }
    public string ValueReferences { get; set; } = string.Empty;
    public uint Index { get; internal set; }
    public uint Order { get; internal set; }
}
