using System;
using System.Diagnostics;
using System.Drawing;

namespace XLibur.Excel.Drawings;

[DebuggerDisplay("R{RowNumber}C{ColumnNumber} {Offset}")]
internal sealed class XLMarker
{
    // Using a range to store the location so that it gets added to the range repository
    // and hence will be adjusted when there are insertions / deletions
    private readonly IXLRange rangeCell;

    internal XLMarker(IXLCell cell)
        : this(cell.AsRange(), new Point(0, 0))
    {
    }

    internal XLMarker(IXLCell cell, Point offset)
        : this(cell.AsRange(), offset)
    {
    }

    private XLMarker(IXLRange rangeCell, Point offset)
    {
        if (rangeCell.RowCount() != 1 || rangeCell.ColumnCount() != 1)
            throw new ArgumentException("Range should contain only one cell.", nameof(rangeCell));

        this.rangeCell = rangeCell;
        Offset = offset;
    }

    public IXLCell Cell => rangeCell.FirstCell();

    public int ColumnNumber => rangeCell.RangeAddress.FirstAddress.ColumnNumber;

    public Point Offset { get; set; }

    public int RowNumber => rangeCell.RangeAddress.FirstAddress.RowNumber;
}