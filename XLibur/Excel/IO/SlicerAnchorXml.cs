using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// The drawing side of a slicer: the graphic frame Excel draws it through, and the anchor that
/// fixes that frame to the grid.
/// </summary>
/// <remarks>
/// <para>
/// A slicer is not drawn by its own part. <c>xl/slicers/slicerN.xml</c> says what the buttons filter
/// and what they look like; where the panel sits is a <c>xdr:graphicFrame</c> in the sheet's drawing
/// part, holding nothing but the slicer's name. Without it the workbook opens and the slicer is
/// invisible, which is why this is the sixth of the six pieces a created slicer needs.
/// </para>
/// <para>
/// The anchor itself is built by <see cref="DrawingAnchorFactory"/> and nowhere else. Slicers get
/// <see cref="XLPicturePlacement.Move"/>: a one-cell anchor takes a top-left marker and an explicit
/// size, which is exactly what a slicer has, and it means the panel moves with the rows and columns
/// above it without being stretched by them. That matches the <c>editAs="oneCell"</c> Excel writes
/// for the pivot slicer in the round-trip fixture.
/// </para>
/// <para>
/// <b>The factory's A1 fallback is never taken here, deliberately.</b> Its remarks are explicit that
/// a missing marker is not an error and not an omission — it silently becomes A1. For a picture that
/// is a reasonable default; for a slicer it would drop the panel on top of the first cell of the
/// sheet, covering the data it is meant to filter. So a created slicer is given a marker before the
/// factory ever sees it — see <see cref="XLSlicers"/> — and the fallback stays unreachable.
/// </para>
/// </remarks>
internal static class SlicerAnchorXml
{
    /// <summary>The graphic data URI Excel uses for a slicer frame, for both kinds of slicer.</summary>
    private const string SlicerGraphicUri = "http://schemas.microsoft.com/office/drawing/2010/slicer";

    private static readonly DrawingFrameSpec Spec =
        new(SlicerGraphicUri, "sle", "slicer", SlicerGraphicUri);

    internal static void Append(Xdr.WorksheetDrawing worksheetDrawing, XLSlicer xlSlicer)
    {
        var frame = DrawingFrameXml.BuildFrame(worksheetDrawing, Spec, xlSlicer.Name);

        var anchor = DrawingAnchorFactory.Create(
            XLPicturePlacement.Move,
            new DrawingAnchorGeometry
            {
                Worksheet = xlSlicer.Worksheet,
                LeftPx = 0,
                TopPx = 0,
                WidthPx = xlSlicer.WidthPx,
                HeightPx = xlSlicer.HeightPx,

                // Never null by the time this runs; see the remarks on the A1 fallback.
                FromMarker = xlSlicer.FromMarker,
                ToMarker = xlSlicer.ToMarker,
            },
            frame);

        worksheetDrawing.Append(anchor);
    }

    internal static void Move(DrawingsPart drawingsPart, XLSlicer xlSlicer)
    {
        if (xlSlicer.FromMarker is not { } target)
            return;

        DrawingFrameXml.MoveAnchor(drawingsPart, Spec, xlSlicer.Name, target);
    }

    internal static void Remove(DrawingsPart? drawingsPart, XLSlicer xlSlicer) =>
        DrawingFrameXml.RemoveAnchor(drawingsPart, Spec, xlSlicer.Name);

    internal static void ReadPositions(DrawingsPart? drawingsPart, XLSlicers slicers)
    {
        var worksheetDrawing = drawingsPart?.WorksheetDrawing;
        if (worksheetDrawing is null)
            return;

        foreach (var slicer in slicers.Items)
        {
            var anchor = DrawingFrameXml.FindAnchor(worksheetDrawing, Spec, slicer.Name);
            if (anchor is null)
                continue;

            var (from, to) = DrawingFrameXml.ReadMarkers(anchor, (XLWorksheet)slicer.Worksheet);
            if (from is not null)
                slicer.FromMarker = from;

            if (to is not null)
                slicer.ToMarker = to;
        }
    }
}
