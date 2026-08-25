using XLibur.Excel.Drawings;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// Where a drawing sits on a sheet and how big it is, in the units the model holds it in: pixels for
/// the free-floating geometry, cell markers for the two anchored forms.
/// <see cref="DrawingAnchorFactory"/> converts to EMU and picks whichever of these the placement
/// actually uses.
/// </summary>
/// <remarks>
/// Every value is supplied whatever the placement, because the geometry describes the drawing rather
/// than the anchor. A free-floating drawing still has a size; an anchored one still has a pixel
/// width and height, which is what the <see cref="ToMarker"/> fallback is derived from.
/// </remarks>
internal sealed class DrawingAnchorGeometry
{
    /// <summary>
    /// The sheet the drawing sits on. It is here for two things the factory cannot work out on its
    /// own: the workbook resolution the pixel values were measured at, and the cell a missing marker
    /// falls back to.
    /// </summary>
    internal required IXLWorksheet Worksheet { get; init; }

    /// <summary>Distance from the left edge of the sheet, in pixels. Used by free-floating only.</summary>
    internal required int LeftPx { get; init; }

    /// <summary>Distance from the top edge of the sheet, in pixels. Used by free-floating only.</summary>
    internal required int TopPx { get; init; }

    /// <summary>The drawing's width in pixels.</summary>
    internal required int WidthPx { get; init; }

    /// <summary>The drawing's height in pixels.</summary>
    internal required int HeightPx { get; init; }

    /// <summary>
    /// The top-left anchor point, or <c>null</c> to take the factory's A1 fallback. Used by the
    /// <see cref="XLPicturePlacement.Move"/> and <see cref="XLPicturePlacement.MoveAndSize"/> forms.
    /// </summary>
    internal XLMarker? FromMarker { get; init; }

    /// <summary>
    /// The bottom-right anchor point, or <c>null</c> to take the factory's A1 fallback. Used by the
    /// <see cref="XLPicturePlacement.MoveAndSize"/> form only.
    /// </summary>
    internal XLMarker? ToMarker { get; init; }
}
