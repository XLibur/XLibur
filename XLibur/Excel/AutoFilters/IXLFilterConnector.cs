namespace XLibur.Excel;

public interface IXLFilterConnector
{
    IXLCustomFilteredColumn And { get; }
    IXLCustomFilteredColumn Or { get; }
}
