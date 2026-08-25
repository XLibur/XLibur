using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.Drawings;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// The graphic frame Excel draws a named control through, and the anchor that fixes it to the grid.
/// </summary>
/// <remarks>
/// Shared by slicers and timelines. Neither control is drawn by its own part: the part says what the
/// control filters, and a <c>xdr:graphicFrame</c> in the sheet's drawing says where it sits. Without
/// the frame the workbook opens and the control is simply invisible.
/// </remarks>
internal static class DrawingFrameXml
{
    /// <summary>
    /// The frame Excel recognises as the named control: a frame holding nothing but its name, under
    /// the spec's graphic-data URI.
    /// </summary>
    /// <remarks>
    /// Excel wraps this frame in <c>mc:AlternateContent</c>, whose <c>mc:Fallback</c> holds a
    /// rectangle explaining the control to a version of Excel too old to draw one. XLibur writes the
    /// frame directly, for two reasons. The wrapper protects nothing — the only readers the fallback
    /// serves are ones that could not have drawn the control anyway — and
    /// <c>OpenXmlValidator</c> rejects <c>mc:AlternateContent</c> as the content of a
    /// <c>xdr:oneCellAnchor</c>, which is the anchor form both controls use.
    /// </remarks>
    internal static Xdr.GraphicFrame BuildFrame(
        Xdr.WorksheetDrawing worksheetDrawing, in DrawingFrameSpec spec, string name)
    {
        // The SDK has no typed class for sle:slicer or tsle:timeslicer, so it is built as an unknown
        // element — which is also how it comes back when a file carrying one is read.
        var child = new OpenXmlUnknownElement(spec.Prefix, spec.LocalName, spec.ChildNamespace);
        child.SetAttribute(new OpenXmlAttribute(string.Empty, "name", string.Empty, name));
        child.AddNamespaceDeclaration(spec.Prefix, spec.ChildNamespace);

        return new Xdr.GraphicFrame(
            new Xdr.NonVisualGraphicFrameProperties(
                new Xdr.NonVisualDrawingProperties { Id = NextFrameId(worksheetDrawing), Name = name },
                new Xdr.NonVisualGraphicFrameDrawingProperties()),

            // Zero, as Excel writes it: the anchor decides where the frame goes and these are
            // ignored. The element is required all the same.
            new Xdr.Transform(
                new A.Offset { X = 0, Y = 0 },
                new A.Extents { Cx = 0, Cy = 0 }),

            new A.Graphic(new A.GraphicData(child) { Uri = spec.GraphicUri }))
        {
            Macro = string.Empty,
        };
    }

    /// <summary>
    /// The anchor holding the frame for the named control, in whichever of the three forms it uses.
    /// </summary>
    internal static OpenXmlCompositeElement? FindAnchor(
        Xdr.WorksheetDrawing worksheetDrawing, in DrawingFrameSpec spec, string name)
    {
        var graphicUri = spec.GraphicUri;
        var localName = spec.LocalName;

        foreach (var anchor in worksheetDrawing.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            if (anchor is not (Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor))
                continue;

            // The frame may be a direct child or, as Excel writes it, inside mc:AlternateContent.
            foreach (var graphicData in anchor.Descendants<A.GraphicData>())
            {
                if (graphicData.Uri?.Value != graphicUri)
                    continue;

                if (NameOfControl(graphicData, localName) == name)
                    return anchor;
            }
        }

        return null;
    }

    /// <summary>
    /// Moves a frame's anchor, shifting both corners by the same number of rows and columns so the
    /// control keeps the size it had.
    /// </summary>
    /// <remarks>
    /// The anchor is edited rather than replaced. Excel's own frame carries an
    /// <c>mc:AlternateContent</c> wrapper, a fallback shape and an <c>a16:creationId</c>, none of
    /// which XLibur models — replacing the anchor to move a control three columns would throw all of
    /// that away.
    /// </remarks>
    internal static void MoveAnchor(
        DrawingsPart drawingsPart, in DrawingFrameSpec spec, string name, XLMarker target)
    {
        var worksheetDrawing = drawingsPart.WorksheetDrawing;
        if (worksheetDrawing is null)
            return;

        var anchor = FindAnchor(worksheetDrawing, spec, name);
        var from = anchor?.GetFirstChild<Xdr.FromMarker>();
        if (anchor is null || from is null)
            return;

        // The delta is taken from what the file says rather than from the model's own old marker, so
        // a control moved twice before a save still lands where the caller last put it.
        var columnDelta = target.ColumnNumber - 1 - ReadInt(from.ColumnId);
        var rowDelta = target.RowNumber - 1 - ReadInt(from.RowId);

        WriteMarker(from, ReadInt(from.ColumnId) + columnDelta, ReadInt(from.RowId) + rowDelta);

        // A one-cell or absolute anchor has no bottom-right corner to keep in step.
        if (anchor.GetFirstChild<Xdr.ToMarker>() is { } to)
            WriteMarker(to, ReadInt(to.ColumnId) + columnDelta, ReadInt(to.RowId) + rowDelta);
    }

    /// <summary>
    /// Takes the anchored frame of a removed control out of the sheet's drawing.
    /// </summary>
    /// <remarks>
    /// The whole anchor goes, not just the frame inside it. An anchor is a position with one thing
    /// anchored at it, so an emptied one is a position for nothing — and where Excel wrapped the
    /// frame in <c>mc:AlternateContent</c>, the fallback shape would be left behind to be drawn in
    /// its place.
    /// </remarks>
    internal static void RemoveAnchor(DrawingsPart? drawingsPart, in DrawingFrameSpec spec, string name)
    {
        var worksheetDrawing = drawingsPart?.WorksheetDrawing;
        if (worksheetDrawing is null)
            return;

        FindAnchor(worksheetDrawing, spec, name)?.Remove();
    }

    /// <summary>
    /// The two corner markers of an anchor. Either may be absent: a one-cell anchor has no
    /// bottom-right corner, and an absolute anchor has neither.
    /// </summary>
    internal static (XLMarker? From, XLMarker? To) ReadMarkers(
        OpenXmlCompositeElement anchor, XLWorksheet worksheet)
    {
        var from = anchor.GetFirstChild<Xdr.FromMarker>() is { } f ? ReadMarker(worksheet, f) : null;
        var to = anchor.GetFirstChild<Xdr.ToMarker>() is { } t ? ReadMarker(worksheet, t) : null;
        return (from, to);
    }

    /// <summary>
    /// A drawing id no other drawing on the sheet is using. Ids are unique within the drawing part,
    /// not within the workbook.
    /// </summary>
    private static uint NextFrameId(Xdr.WorksheetDrawing worksheetDrawing)
    {
        var used = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().ToList();
        return used.Count > 0 ? used.Max(p => p.Id?.Value ?? 0U) + 1 : 1U;
    }

    /// <summary>
    /// The control name off a graphic frame's single child, which the SDK has no typed class for and
    /// so deserialises as an unknown element.
    /// </summary>
    private static string? NameOfControl(A.GraphicData graphicData, string localName)
    {
        foreach (var child in graphicData.ChildElements)
        {
            if (child.LocalName != localName)
                continue;

            var name = child.GetAttribute("name", string.Empty).Value;
            if (!string.IsNullOrEmpty(name))
                return name;
        }

        return null;
    }

    private static XLMarker ReadMarker(XLWorksheet worksheet, Xdr.MarkerType marker)
    {
        // Markers are written zero-based; the model counts from one.
        var column = ReadInt(marker.ColumnId) + 1;
        var row = ReadInt(marker.RowId) + 1;

        var cell = worksheet.Cell(
            row < 1 ? 1 : row,
            column < 1 ? 1 : column);

        return new XLMarker(cell, new System.Drawing.Point(
            EmuToPixels(ReadLong(marker.ColumnOffset), worksheet.Workbook.DpiX),
            EmuToPixels(ReadLong(marker.RowOffset), worksheet.Workbook.DpiY)));
    }

    private static void WriteMarker(Xdr.MarkerType marker, int column, int row)
    {
        marker.ColumnId = new Xdr.ColumnId((column < 0 ? 0 : column).ToInvariantString());
        marker.RowId = new Xdr.RowId((row < 0 ? 0 : row).ToInvariantString());
    }

    private static int ReadInt(OpenXmlLeafTextElement? element) =>
        int.TryParse(element?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : 0;

    private static long ReadLong(OpenXmlLeafTextElement? element) =>
        long.TryParse(element?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : 0;

    /// <summary>
    /// The inverse of <see cref="DrawingUnits.PixelsToEmu"/>, for reporting an offset a file carries.
    /// </summary>
    private static int EmuToPixels(long emu, double resolution) =>
        emu == 0 ? 0 : (int)System.Math.Round(emu * resolution / 914400d);
}
