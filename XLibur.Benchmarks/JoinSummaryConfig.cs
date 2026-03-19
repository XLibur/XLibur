using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Reports;
using BenchmarkDotNet.Running;

namespace XLibur.Benchmarks;

public class JoinSummaryConfig : ManualConfig
{
    public JoinSummaryConfig()
    {
        WithOption(ConfigOptions.JoinSummary, true);
        AddColumn(new LibraryNameColumn());
    }
}

public class LibraryNameColumn : IColumn
{
    public string Id => "Library";
    public string ColumnName => "Library";
    public bool IsDefault(Summary summary, BenchmarkCase b) => false;

    public string GetValue(Summary summary, BenchmarkCase b)
    {
        var typeName = b.Descriptor.Type.Name;
        return typeName switch
        {
            var n when n.Contains("ClosedXml") => "ClosedXML",
            var n when n.Contains("EpPlus") => "EPPlus",
            var n when n.Contains("XLibur") => "XLibur",
            _ => typeName
        };
    }

    public string GetValue(Summary s, BenchmarkCase b, SummaryStyle style) => GetValue(s, b);
    public bool IsAvailable(Summary s) => true;
    public bool AlwaysShow => true;
    public ColumnCategory Category => ColumnCategory.Job;
    public int PriorityInCategory => -10;
    public bool IsNumeric => false;
    public UnitType UnitType => UnitType.Dimensionless;
    public string Legend => "Library under test";
}
