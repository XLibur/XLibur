using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using XLibur.Excel.IO.DrawingML;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes the worksheet half of a timeline: the timelines part holding its definition and the
/// <c>extLst</c> reference that makes the worksheet point at it.
/// </summary>
/// <remarks>
/// A surviving timeline part is worthless if the sheet stops referencing it, and the worksheet part
/// is rebuilt from the model on every save while the timeline part is not.
/// </remarks>
internal static class TimelineWriter
{
    /// <summary>The worksheet extension holding the list of timelines on the sheet.</summary>
    private const string TimelineExtensionUri = "{7E03D99C-DC04-49d9-9315-930204A7B6E9}";

    private const string X15Main2010SsNs = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

    internal static void WriteTimelines(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        WorksheetPart worksheetPart,
        SaveContext context)
    {
        var timelines = xlWorksheet.TimelinesInternal;

        RemoveDeletedTimelines(worksheet, cm, worksheetPart, timelines);

        foreach (var timeline in timelines.Items)
        {
            if (!timeline.IsNew)
            {
                // A timeline that already exists in the package is never regenerated — that is what
                // keeps the parts of its XML XLibur does not model intact. Only the properties the
                // caller actually changed are patched into the existing part.
                TimelinePatcher.PatchTimeline(worksheetPart, timeline);
                continue;
            }

            WriteNewTimeline(worksheet, cm, worksheetPart, timeline, context);
            timeline.IsNew = false;
        }
    }

    private static void WriteNewTimeline(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        WorksheetPart worksheetPart,
        XLTimeline xlTimeline,
        SaveContext context)
    {
        var relId = context.RelIdGenerator.GetNext(RelType.Workbook);
        var part = worksheetPart.AddNewPart<TimeLinePart>(relId);
        xlTimeline.PartRelId = relId;

        var root = new X15.Timelines();
        root.AddNamespaceDeclaration("x", Main2006SsNs);

        var timeline = new X15.Timeline
        {
            Name = xlTimeline.Name,
            Cache = xlTimeline.Cache.Name,
            Caption = xlTimeline.Caption,
        };

        // Attributes at their schema default are left off, which is what Excel writes and what keeps
        // a generated part comparable with a hand-made one.
        if (!xlTimeline.ShowHeader)
            timeline.ShowHeader = false;

        if (!xlTimeline.ShowSelectionLabel)
            timeline.ShowSelectionLabel = false;

        if (!xlTimeline.ShowTimeLevel)
            timeline.ShowTimeLevel = false;

        if (!xlTimeline.ShowHorizontalScrollbar)
            timeline.ShowHorizontalScrollbar = false;

        if (xlTimeline.Style is { } style)
            timeline.Style = style;

        if (xlTimeline.LevelRaw != 0)
        {
            timeline.Level = xlTimeline.LevelRaw;

            // Excel writes selectionLevel alongside level and keeps the two in step on a timeline
            // that has never been scrubbed.
            timeline.SelectionLevel = xlTimeline.LevelRaw;
        }

        root.AppendChild(timeline);
        part.Timelines = root;

        EnsureTimelineReference(worksheet, cm, relId);
        WriteAnchor(worksheet, cm, worksheetPart, xlTimeline, context);
    }

    /// <summary>
    /// Draws the timeline: the graphic frame in the sheet's drawing part, and the sheet's reference
    /// to that part.
    /// </summary>
    /// <remarks>
    /// The sixth of the six pieces a created timeline needs. Without it the workbook opens and the
    /// timeline is simply not there, because <c>xl/timelines/timelineN.xml</c> says what a timeline
    /// scrubs but never where it sits.
    /// </remarks>
    private static void WriteAnchor(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        WorksheetPart worksheetPart,
        XLTimeline xlTimeline,
        SaveContext context)
    {
        var drawingsPart = DrawingPartScaffold.EnsureDrawingsPart(worksheetPart, context);
        var worksheetDrawing = drawingsPart.WorksheetDrawing!;
        DrawingPartScaffold.EnsureNamespaces(worksheetDrawing);

        TimelineAnchorXml.Append(worksheetDrawing, xlTimeline);

        DrawingPartScaffold.EnsureDrawingElement(worksheet, cm, worksheetPart, drawingsPart);
    }

    private static void EnsureTimelineReference(
        Worksheet worksheet, XLWorksheetContentManager cm, string relId)
    {
        var list = SheetExtensionRefs.EnsureList<X15.TimelineReferences>(
            worksheet, cm, TimelineExtensionUri, "x15", X15Main2010SsNs);

        if (!list.Elements<X15.TimelineReference>().Any(r => r.Id?.Value == relId))
            list.AppendChild(new X15.TimelineReference { Id = relId });
    }

    /// <summary>
    /// Unpicks the worksheet half of every timeline removed since the workbook was loaded.
    /// </summary>
    /// <remarks>
    /// Because a created timeline always gets a part of its own and a loaded one always has one,
    /// removing a timeline removes the whole part rather than one element inside it. The extension
    /// goes once its list is empty, and the extension list once it is. The anchored frame goes from
    /// the sheet's drawing as well, since that is what Excel actually draws a timeline through.
    /// </remarks>
    private static void RemoveDeletedTimelines(
        Worksheet worksheet, XLWorksheetContentManager cm, WorksheetPart worksheetPart, XLTimelines timelines)
    {
        if (timelines.Removed.Count == 0)
            return;

        foreach (var removed in timelines.Removed)
        {
            // The frame lives in the drawing rather than in the timelines part, so it has to go too —
            // otherwise the sheet still asks Excel to draw something the package no longer defines.
            TimelineAnchorXml.Remove(worksheetPart.DrawingsPart, removed);

            if (removed.PartRelId is not { } relId
                || !worksheetPart.Parts.Any(p => p.RelationshipId == relId)
                || worksheetPart.GetPartById(relId) is not TimeLinePart part)
            {
                continue;
            }

            SheetExtensionRefs.RemoveRefs<X15.TimelineReferences>(
                worksheet, cm, r => r is X15.TimelineReference reference && reference.Id?.Value == relId);

            worksheetPart.DeletePart(part);
        }
    }
}
