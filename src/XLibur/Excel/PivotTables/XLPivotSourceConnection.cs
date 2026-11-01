using System;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// Source of data for a <see cref="XLPivotCache"/> that takes data from a connection
/// to external source of data (e.g. database or a workbook).
/// </summary>
internal sealed class XLPivotSourceConnection : IXLPivotSource
{
    public XLPivotSourceConnection(uint connectionId)
    {
        ConnectionId = connectionId;
    }

    public uint ConnectionId { get; }

    public XLPivotSourceKind Kind => XLPivotSourceKind.Connection;

    public bool Equals(IXLPivotSource? otherSource)
    {
        var other = otherSource as XLPivotSourceConnection;
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        return ConnectionId == other.ConnectionId;
    }

    public override bool Equals(object? obj)
    {
        return obj is IXLPivotSource other && Equals(other);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(ConnectionId).GetHashCode();
    }

    /// <summary>XLibur cannot read through a data connection, so there is never a sheet area to report.</summary>
    public bool TryGetSource(XLWorkbook workbook, out XLWorksheet? sheet, out Area? sheetArea)
    {
        sheet = null;
        sheetArea = null;
        return false;
    }
}
