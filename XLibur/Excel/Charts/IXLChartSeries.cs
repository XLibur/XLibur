namespace XLibur.Excel;

public interface IXLChartSeries
{
    string Name { get; set; }
    string? CategoryReferences { get; set; }
    string ValueReferences { get; set; }
    uint Index { get; }
    uint Order { get; }
}
