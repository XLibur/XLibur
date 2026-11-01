using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// The drawing side of a timeline: the graphic frame Excel draws it through, and the anchor that
/// fixes that frame to the grid.
/// </summary>
/// <remarks>
/// A timeline is not drawn by its own part. <c>xl/timelines/timelineN.xml</c> says what the band
/// scrubs and what it looks like; where the band sits is a <c>xdr:graphicFrame</c> in the sheet's
/// drawing part holding a <c>tsle:timeslicer</c> element that carries nothing but the timeline's
/// name.
/// </remarks>
internal static class TimelineAnchorXml
{
    /// <summary>The graphic data URI Excel uses for a timeline frame.</summary>
    private const string TimelineGraphicUri = "http://schemas.microsoft.com/office/drawing/2012/timeslicer";

    internal static readonly DrawingFrameSpec Spec =
        new(TimelineGraphicUri, "tsle", "timeslicer", TimelineGraphicUri);

    /// <summary>
    /// Reads the anchor of every timeline on the sheet, so that a loaded timeline reports where it
    /// is.
    /// </summary>
    /// <remarks>
    /// The frame names the timeline, and the timeline name is unique within the workbook, which is
    /// what pairs the two up. All three anchor forms are read: Excel writes a two-cell anchor,
    /// XLibur writes a one-cell one, and a file from elsewhere may carry either or an absolute one.
    /// </remarks>
    internal static void ReadPositions(DrawingsPart? drawingsPart, XLTimelines timelines)
    {
        var worksheetDrawing = drawingsPart?.WorksheetDrawing;
        if (worksheetDrawing is null)
            return;

        foreach (var timeline in timelines.Items)
        {
            var anchor = DrawingFrameXml.FindAnchor(worksheetDrawing, Spec, timeline.Name);
            if (anchor is null)
                continue;

            var (from, to) = DrawingFrameXml.ReadMarkers(anchor, (XLWorksheet)timeline.Worksheet);
            if (from is not null)
                timeline.FromMarker = from;

            if (to is not null)
                timeline.ToMarker = to;
        }
    }

    /// <summary>
    /// Appends the anchored graphic frame for a newly created timeline to the sheet's drawing.
    /// </summary>
    /// <remarks>
    /// Timelines take <see cref="XLPicturePlacement.Move"/>: a one-cell anchor takes a top-left
    /// marker and an explicit size, which is exactly what a timeline has, and it means the band
    /// moves with the rows and columns above it without being stretched by them. The factory's A1
    /// fallback is never taken — a created timeline is always given a marker first.
    /// </remarks>
    internal static void Append(Xdr.WorksheetDrawing worksheetDrawing, XLTimeline xlTimeline)
    {
        var frame = DrawingFrameXml.BuildFrame(worksheetDrawing, Spec, xlTimeline.Name);

        var anchor = DrawingAnchorFactory.Create(
            XLPicturePlacement.Move,
            new DrawingAnchorGeometry
            {
                Worksheet = xlTimeline.Worksheet,
                LeftPx = 0,
                TopPx = 0,
                WidthPx = xlTimeline.WidthPx,
                HeightPx = xlTimeline.HeightPx,
                FromMarker = xlTimeline.FromMarker,
                ToMarker = xlTimeline.ToMarker,
            },
            frame);

        worksheetDrawing.Append(anchor);
    }

    /// <summary>
    /// Moves the frame of a loaded timeline, shifting both corners by the same number of rows and
    /// columns so the band keeps the size it had.
    /// </summary>
    internal static void Move(DrawingsPart drawingsPart, XLTimeline xlTimeline)
    {
        if (xlTimeline.FromMarker is not { } target)
            return;

        DrawingFrameXml.MoveAnchor(drawingsPart, Spec, xlTimeline.Name, target);
    }

    /// <summary>Takes the anchored frame of a removed timeline out of the sheet's drawing.</summary>
    internal static void Remove(DrawingsPart? drawingsPart, XLTimeline xlTimeline) =>
        DrawingFrameXml.RemoveAnchor(drawingsPart, Spec, xlTimeline.Name);
}
