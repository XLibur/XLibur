using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.IO.DrawingML;

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
}
