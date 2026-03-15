using System;

namespace XLibur.Excel.CalcEngine.Exceptions;

/// <summary>
/// Exception that happens when formula in a cell depends on other cells,
/// but the supporting formulas are still dirty.
/// </summary>
#pragma warning disable S3871
internal sealed class GettingDataException : Exception
#pragma warning restore S3871
{
    public GettingDataException(XLBookPoint point)
    {
        Point = point;
    }

    public XLBookPoint Point { get; }
}
