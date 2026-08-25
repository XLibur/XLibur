using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XLibur.Excel.IO;

/// <summary>
/// Applies model changes to a slicer that already exists in the package.
/// </summary>
/// <remarks>
/// <para>
/// XLibur never regenerates the XML of a slicer it read from a file. That is what carries a
/// slicer's <c>xr10:uid</c>, its <c>startItem</c>, its <c>level</c>, the extension list Excel hangs
/// off it and any attribute a future Excel invents through a load and save untouched — see
/// <c>docs/round-trip-fidelity.md</c>. Regenerating instead is the failure spec 17 records for
/// pictures, where the save path "destroys existing styling".
/// </para>
/// <para>
/// The price of that guarantee is that an edit has to be patched into the element the reader saw.
/// This class does exactly that, and only for the properties the caller actually assigned (see
/// <see cref="XLSlicer.AssignedFormat"/>): a slicer nobody edited is not written to at all, and the
/// part is not even opened for it.
/// </para>
/// <para>
/// This is <see cref="ChartPatcher"/> applied to a much smaller element. The one difference worth
/// noting is that a slicer is found by name rather than by position: <c>slicer/@name</c> is unique
/// within the part, so there is no need for the index-matching scan a chart's series list requires.
/// </para>
/// </remarks>
internal static class SlicerPatcher
{
    /// <summary>
    /// Writes the pending property changes of a loaded slicer back into its slicers part.
    /// </summary>
    internal static void PatchSlicer(WorksheetPart worksheetPart, XLSlicer xlSlicer, double emuPerPoint)
    {
        if (xlSlicer.AssignedFormat == XLSlicerFormat.None)
            return;

        var part = ResolvePart(worksheetPart, xlSlicer);
        if (part?.Slicers is not { } slicers)
            return;

        var slicer = slicers
            .Elements<X14.Slicer>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, xlSlicer.Name, StringComparison.Ordinal));
        if (slicer is null)
            return;

        Apply(slicer, xlSlicer, emuPerPoint);
    }

    private static void Apply(X14.Slicer slicer, XLSlicer xlSlicer, double emuPerPoint)
    {
        var assigned = xlSlicer.AssignedFormat;

        // Each optional attribute is cleared with a typed null rather than a bare one. Assigning
        // `null` to one of these properties goes through the implicit conversion from string or
        // bool, which produces a value wrapping null — serialised as `caption=""`, not as an absent
        // attribute. The cast is what actually removes it.
        if (assigned.HasFlag(XLSlicerFormat.Caption))
        {
            // Excel omits the caption when it matches the name and shows the name instead, so
            // setting the caption back to the name removes the attribute rather than restating it.
            slicer.Caption = string.Equals(xlSlicer.Caption, xlSlicer.Name, StringComparison.Ordinal)
                ? (StringValue?)null
                : xlSlicer.Caption;
        }

        // showCaption and columnCount default to true and 1. Writing a value that is already the
        // default as an explicit attribute is legal but noisy, and it is not what Excel does.
        if (assigned.HasFlag(XLSlicerFormat.ShowCaption))
            slicer.ShowCaption = xlSlicer.ShowCaption ? (BooleanValue?)null : false;

        if (assigned.HasFlag(XLSlicerFormat.Style))
            slicer.Style = xlSlicer.Style is { } style ? style : (StringValue?)null;

        if (assigned.HasFlag(XLSlicerFormat.ColumnCount))
            slicer.ColumnCount = xlSlicer.ColumnCount == 1 ? (UInt32Value?)null : xlSlicer.ColumnCount;

        if (assigned.HasFlag(XLSlicerFormat.RowHeight))
        {
            // Unlike the others, rowHeight is a required attribute, so clearing it writes Excel's
            // default rather than removing it — a slicer without one fails schema validation.
            var rowHeightPt = xlSlicer.RowHeightPt ?? XLSlicer.DefaultRowHeightPt;
            slicer.RowHeight = (uint)Math.Round(rowHeightPt * emuPerPoint, MidpointRounding.AwayFromZero);
        }
    }

    /// <summary>
    /// The slicers part a loaded slicer was read from.
    /// </summary>
    /// <remarks>
    /// Opening the part here is what finally attaches its DOM, and it happens only for a slicer
    /// with a pending change. Everything <see cref="SlicerReader"/> does is deliberately detached
    /// so that this is the single point at which a slicer part stops being copied through verbatim.
    /// </remarks>
    private static SlicersPart? ResolvePart(WorksheetPart worksheetPart, XLSlicer xlSlicer)
    {
        if (xlSlicer.PartRelId is null)
            return null;

        // GetPartById throws for an unknown id, which is reachable when the slicer came from a
        // package this one was not saved from.
        if (!worksheetPart.Parts.Any(p => p.RelationshipId == xlSlicer.PartRelId))
            return null;

        return worksheetPart.GetPartById(xlSlicer.PartRelId) as SlicersPart;
    }
}
