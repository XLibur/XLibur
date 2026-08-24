using System;
using System.Drawing;
using DocumentFormat.OpenXml;
using XLibur.Excel.Drawings;
using XLibur.Extensions;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// Builds the spreadsheet-drawing anchor that fixes a drawing to a sheet, in whichever of the three
/// forms the placement calls for.
/// </summary>
/// <remarks>
/// <para>
/// The anchor is the same regardless of what it holds — a picture, a shape, a text box, a chart
/// frame — so the content is handed in rather than built here. What varies between the forms is
/// only which of the geometry's parts is used: <see cref="XLPicturePlacement.FreeFloating"/> takes
/// the pixel position and size, <see cref="XLPicturePlacement.Move"/> takes the top-left marker and
/// the size, and <see cref="XLPicturePlacement.MoveAndSize"/> takes both markers and no size at all.
/// </para>
/// <para>
/// <b>The A1 fallbacks are part of the contract, not an implementation detail.</b> A caller that
/// leaves a marker unset does not get an exception and does not get a missing element — it gets a
/// marker at A1:
/// </para>
/// <list type="bullet">
/// <item>
/// A missing <see cref="DrawingAnchorGeometry.FromMarker"/> becomes A1 at offset zero, for both
/// anchored forms.
/// </item>
/// <item>
/// A missing <see cref="DrawingAnchorGeometry.ToMarker"/> becomes A1 offset by the drawing's own
/// pixel width and height, so a <see cref="XLPicturePlacement.MoveAndSize"/> drawing nobody
/// positioned still comes out its intended size rather than collapsed to nothing.
/// </item>
/// </list>
/// <para>
/// This is reachable through ordinary use, not just by misuse: a picture is created
/// <see cref="XLPicturePlacement.MoveAndSize"/> with neither marker set, so a picture nobody moved
/// takes both fallbacks at once.
/// </para>
/// <para>
/// <see cref="XLPicturePlacement"/> is named for pictures because pictures were the first drawings
/// XLibur wrote, but what it describes — how a drawing reacts when the rows and columns underneath
/// it are resized — belongs to every kind of drawing. A second enum saying the same three things
/// would be exactly the duplication this layer exists to prevent, so other drawing kinds should
/// reuse this one rather than declare their own.
/// </para>
/// </remarks>
internal static class DrawingAnchorFactory
{
    /// <summary>
    /// Creates the anchor for the given placement, wrapping <paramref name="content"/> and closing
    /// with the <c>xdr:clientData</c> the schema requires.
    /// </summary>
    /// <param name="placement">Which of the three anchor forms to build.</param>
    /// <param name="geometry">Where the drawing sits and how big it is. See the A1 fallbacks above.</param>
    /// <param name="content">
    /// The drawing itself — an <c>xdr:pic</c>, <c>xdr:sp</c> or any other member of the anchor's
    /// content choice group. It is placed into the anchor as-is, so it must not already have a
    /// parent.
    /// </param>
    internal static OpenXmlCompositeElement Create(
        XLPicturePlacement placement,
        DrawingAnchorGeometry geometry,
        OpenXmlElement content)
    {
        var workbook = geometry.Worksheet.Workbook;
        var dpiX = workbook.DpiX;
        var dpiY = workbook.DpiY;

        switch (placement)
        {
            case XLPicturePlacement.FreeFloating:
                return new Xdr.AbsoluteAnchor(
                    new Xdr.Position
                    {
                        X = DrawingUnits.PixelsToEmu(geometry.LeftPx, dpiX),
                        Y = DrawingUnits.PixelsToEmu(geometry.TopPx, dpiY)
                    },
                    Extent(geometry, dpiX, dpiY),
                    content,
                    new Xdr.ClientData()
                );

            case XLPicturePlacement.MoveAndSize:
            {
                // Resolved in this order because building a fallback marker registers a range with
                // the workbook, and the inline code this replaced registered the from marker first.
                var from = CreateMarker<Xdr.FromMarker>(FromMarkerOf(geometry), dpiX, dpiY);
                var to = CreateMarker<Xdr.ToMarker>(ToMarkerOf(geometry), dpiX, dpiY);

                return new Xdr.TwoCellAnchor(from, to, content, new Xdr.ClientData());
            }

            case XLPicturePlacement.Move:
                return new Xdr.OneCellAnchor(
                    CreateMarker<Xdr.FromMarker>(FromMarkerOf(geometry), dpiX, dpiY),
                    Extent(geometry, dpiX, dpiY),
                    content,
                    new Xdr.ClientData()
                );

            default:
                throw new ArgumentOutOfRangeException(
                    nameof(placement), placement, "Unsupported picture placement.");
        }
    }

    private static Xdr.Extent Extent(DrawingAnchorGeometry geometry, double dpiX, double dpiY) => new()
    {
        Cx = DrawingUnits.PixelsToEmu(geometry.WidthPx, dpiX),
        Cy = DrawingUnits.PixelsToEmu(geometry.HeightPx, dpiY)
    };

    /// <summary>The top-left anchor point, falling back to A1 at offset zero.</summary>
    private static XLMarker FromMarkerOf(DrawingAnchorGeometry geometry) =>
        geometry.FromMarker ?? new XLMarker(geometry.Worksheet.Cell("A1"));

    /// <summary>
    /// The bottom-right anchor point, falling back to A1 offset by the drawing's own size — which is
    /// what keeps an unpositioned <see cref="XLPicturePlacement.MoveAndSize"/> drawing its right
    /// size rather than collapsing it onto its own top-left corner.
    /// </summary>
    private static XLMarker ToMarkerOf(DrawingAnchorGeometry geometry) =>
        geometry.ToMarker ?? new XLMarker(
            geometry.Worksheet.Cell("A1"), new Point(geometry.WidthPx, geometry.HeightPx));

    /// <summary>
    /// Fills in one marker. <c>xdr:from</c> and <c>xdr:to</c> are the same schema type differing only
    /// in name, so they are built once: a second copy of the off-by-one and the two unit conversions
    /// is exactly the drift this layer exists to prevent.
    /// </summary>
    /// <remarks>
    /// The column and row are written zero-based, while the model counts from one.
    /// </remarks>
    private static TMarker CreateMarker<TMarker>(XLMarker marker, double dpiX, double dpiY)
        where TMarker : Xdr.MarkerType, new() => new()
    {
        ColumnId = new Xdr.ColumnId((marker.ColumnNumber - 1).ToInvariantString()),
        RowId = new Xdr.RowId((marker.RowNumber - 1).ToInvariantString()),
        ColumnOffset = new Xdr.ColumnOffset(DrawingUnits.PixelsToEmu(marker.Offset.X, dpiX).ToInvariantString()),
        RowOffset = new Xdr.RowOffset(DrawingUnits.PixelsToEmu(marker.Offset.Y, dpiY).ToInvariantString())
    };
}
