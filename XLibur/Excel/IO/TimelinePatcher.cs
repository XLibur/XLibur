using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Applies model changes to a timeline that already exists in the package.
/// </summary>
/// <remarks>
/// <para>
/// XLibur never regenerates the XML of a timeline it read from a file. That is what carries a
/// timeline's <c>selectionLevel</c>, its <c>scrollPosition</c>, the extension list Excel hangs off
/// it and any attribute a future Excel invents through a load and save untouched.
/// </para>
/// <para>
/// The price of that guarantee is that an edit has to be patched into the element the reader saw.
/// This does exactly that, and only for the properties the caller actually assigned (see
/// <see cref="XLTimeline.AssignedFormat"/>): a timeline nobody edited is not written to at all, and
/// the part is not even opened for it.
/// </para>
/// </remarks>
internal static class TimelinePatcher
{
    internal static void PatchTimeline(WorksheetPart worksheetPart, XLTimeline xlTimeline)
    {
        if (xlTimeline.AssignedFormat == XLTimelineFormat.None)
            return;

        // Moving a timeline touches the drawing part rather than the timelines part, so the two are
        // resolved separately: assigning only a position must not open the timelines part, and
        // assigning only a caption must not open the drawing.
        if (xlTimeline.AssignedFormat.HasFlag(XLTimelineFormat.Position)
            && worksheetPart.DrawingsPart is { } drawingsPart)
        {
            TimelineAnchorXml.Move(drawingsPart, xlTimeline);
        }

        if ((xlTimeline.AssignedFormat & ~XLTimelineFormat.Position) == XLTimelineFormat.None)
            return;

        var part = ResolvePart(worksheetPart, xlTimeline);
        if (part?.Timelines is not { } timelines)
            return;

        var timeline = timelines
            .Elements<X15.Timeline>()
            .FirstOrDefault(t => string.Equals(t.Name?.Value, xlTimeline.Name, StringComparison.Ordinal));
        if (timeline is null)
            return;

        Apply(timeline, xlTimeline);
    }

    private static void Apply(X15.Timeline timeline, XLTimeline xlTimeline)
    {
        var assigned = xlTimeline.AssignedFormat;

        // Each optional attribute is cleared with a typed null rather than a bare one. Assigning
        // `null` goes through the implicit conversion from string or bool and produces a value
        // wrapping null — serialised as `caption=""`, not as an absent attribute. The cast is what
        // actually removes it.
        if (assigned.HasFlag(XLTimelineFormat.Caption))
        {
            // Excel omits the caption when it matches the name and shows the name instead, so
            // setting the caption back to the name removes the attribute rather than restating it.
            timeline.Caption = string.Equals(xlTimeline.Caption, xlTimeline.Name, StringComparison.Ordinal)
                ? (StringValue?)null
                : xlTimeline.Caption;
        }

        // The four booleans default to true; writing a value that is already the default is legal
        // but noisy, and it is not what Excel does.
        if (assigned.HasFlag(XLTimelineFormat.ShowHeader))
            timeline.ShowHeader = xlTimeline.ShowHeader ? (BooleanValue?)null : false;

        if (assigned.HasFlag(XLTimelineFormat.ShowSelectionLabel))
            timeline.ShowSelectionLabel = xlTimeline.ShowSelectionLabel ? (BooleanValue?)null : false;

        if (assigned.HasFlag(XLTimelineFormat.ShowTimeLevel))
            timeline.ShowTimeLevel = xlTimeline.ShowTimeLevel ? (BooleanValue?)null : false;

        if (assigned.HasFlag(XLTimelineFormat.ShowHorizontalScrollbar))
        {
            timeline.ShowHorizontalScrollbar =
                xlTimeline.ShowHorizontalScrollbar ? (BooleanValue?)null : false;
        }

        if (assigned.HasFlag(XLTimelineFormat.Style))
            timeline.Style = xlTimeline.Style is { } style ? style : (StringValue?)null;

        // level defaults to 0, so a timeline set back to Years drops the attribute.
        if (assigned.HasFlag(XLTimelineFormat.Level))
            timeline.Level = xlTimeline.LevelRaw == 0 ? (UInt32Value?)null : xlTimeline.LevelRaw;
    }

    /// <summary>
    /// The timelines part a loaded timeline was read from.
    /// </summary>
    /// <remarks>
    /// Opening the part here is what finally attaches its DOM, and it happens only for a timeline
    /// with a pending change. Everything <see cref="TimelineReader"/> does is deliberately detached
    /// so that this is the single point at which a timeline part stops being copied through verbatim.
    /// </remarks>
    private static TimeLinePart? ResolvePart(WorksheetPart worksheetPart, XLTimeline xlTimeline)
    {
        if (xlTimeline.PartRelId is null)
            return null;

        // GetPartById throws for an unknown id, which is reachable when the timeline came from a
        // package this one was not saved from.
        if (!worksheetPart.Parts.Any(p => p.RelationshipId == xlTimeline.PartRelId))
            return null;

        return worksheetPart.GetPartById(xlTimeline.PartRelId) as TimeLinePart;
    }
}
