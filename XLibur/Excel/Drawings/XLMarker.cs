using System.Diagnostics;
using XLibur.Excel.Coordinates;
using PixelOffset = System.Drawing.Point;

namespace XLibur.Excel.Drawings;

/// <summary>
/// One corner of a picture's anchor: the cell it hangs off, and a pixel offset within that cell.
/// </summary>
/// <remarks>
/// Anchors are shifted by <see cref="DrawingAnchorListener"/>, registered in
/// <see cref="XLWorksheet.GetSheetListeners"/>. Before spec 33 the anchor was held as a one-cell
/// range rather than as a point, allocated for no reason except that the repository of live ranges
/// shifts a range and does not shift a point — and the field's own comment said so. It was
/// validated to be exactly one cell on the way in and unwrapped again by every reader, so it was
/// never a range in any sense but the one that made it move: the seam smuggled past, in a comment,
/// in shipped code. It is a <see cref="Coordinates.Point"/> now, and it moves because a listener
/// tells it to.
/// <para>
/// Two types called <c>Point</c> meet here: <see cref="Offset"/> is a <c>System.Drawing.Point</c> of
/// pixels, aliased to <c>PixelOffset</c> so it cannot be confused with the row-and-column
/// <see cref="Coordinates.Point"/> the anchor is.
/// </para>
/// </remarks>
[DebuggerDisplay("R{RowNumber}C{ColumnNumber} {Offset}")]
internal sealed class XLMarker
{
    private readonly XLWorksheet _worksheet;

    internal XLMarker(IXLCell cell)
        : this(cell, new PixelOffset(0, 0))
    {
    }

    internal XLMarker(IXLCell cell, PixelOffset offset)
    {
        _worksheet = (XLWorksheet)cell.Worksheet;
        Anchor = new Point(cell.Address.RowNumber, cell.Address.ColumnNumber);
        Offset = offset;
    }

    /// <summary>The anchored cell, as a row and a column. Moved by <see cref="DrawingAnchorListener"/>.</summary>
    internal Point Anchor { get; set; }

    public IXLCell Cell => _worksheet.Internals.CellsCollection.GetCell(Anchor);

    public int ColumnNumber => Anchor.Column;

    public PixelOffset Offset { get; set; }

    public int RowNumber => Anchor.Row;
}
