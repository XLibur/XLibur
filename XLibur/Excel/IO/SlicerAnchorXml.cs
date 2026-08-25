using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using A = DocumentFormat.OpenXml.Drawing;
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

    private const string Slicer2010Ns = "http://schemas.microsoft.com/office/drawing/2010/slicer";

    /// <summary>
    /// Appends the anchored graphic frame for a newly created slicer to the sheet's drawing.
    /// </summary>
    internal static void Append(Xdr.WorksheetDrawing worksheetDrawing, XLSlicer xlSlicer)
    {
        var frame = BuildGraphicFrame(worksheetDrawing, xlSlicer);

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

    /// <summary>
    /// Moves the frame of a loaded slicer, shifting both corners by the same number of rows and
    /// columns so the panel keeps the size it had.
    /// </summary>
    /// <remarks>
    /// The anchor is edited rather than replaced. Excel's own frame carries an
    /// <c>mc:AlternateContent</c> wrapper, a fallback shape and an <c>a16:creationId</c>, none of
    /// which XLibur models — replacing the anchor to move a slicer three columns would throw all of
    /// that away, which is the failure spec 17 records for pictures.
    /// </remarks>
    internal static void Move(DrawingsPart drawingsPart, XLSlicer xlSlicer)
    {
        var worksheetDrawing = drawingsPart.WorksheetDrawing;
        if (worksheetDrawing is null || xlSlicer.FromMarker is not { } target)
            return;

        var anchor = FindAnchor(worksheetDrawing, xlSlicer.Name);
        var from = anchor?.GetFirstChild<Xdr.FromMarker>();
        if (anchor is null || from is null)
            return;

        // The delta is taken from what the file says rather than from the model's own old marker,
        // so a slicer moved twice before a save still lands where the caller last put it.
        var columnDelta = target.ColumnNumber - 1 - ReadInt(from.ColumnId);
        var rowDelta = target.RowNumber - 1 - ReadInt(from.RowId);

        WriteMarker(from, ReadInt(from.ColumnId) + columnDelta, ReadInt(from.RowId) + rowDelta);

        // A one-cell or absolute anchor has no bottom-right corner to keep in step.
        if (anchor.GetFirstChild<Xdr.ToMarker>() is { } to)
            WriteMarker(to, ReadInt(to.ColumnId) + columnDelta, ReadInt(to.RowId) + rowDelta);
    }

    /// <summary>
    /// Reads the anchor of every slicer on the sheet, so that a loaded slicer reports where it is.
    /// </summary>
    /// <remarks>
    /// The frame names the slicer, and the slicer name is unique within the workbook, which is what
    /// pairs the two up. All three anchor forms are read: Excel writes a two-cell anchor, XLibur
    /// writes a one-cell one, and a file from elsewhere may carry either or an absolute one.
    /// </remarks>
    internal static void ReadPositions(DrawingsPart? drawingsPart, XLSlicers slicers)
    {
        var worksheetDrawing = drawingsPart?.WorksheetDrawing;
        if (worksheetDrawing is null)
            return;

        foreach (var slicer in slicers.Items)
        {
            var anchor = FindAnchor(worksheetDrawing, slicer.Name);
            if (anchor is null)
                continue;

            var worksheet = (XLWorksheet)slicer.Worksheet;

            if (anchor.GetFirstChild<Xdr.FromMarker>() is { } from)
                slicer.FromMarker = ReadMarker(worksheet, from);

            if (anchor.GetFirstChild<Xdr.ToMarker>() is { } to)
                slicer.ToMarker = ReadMarker(worksheet, to);
        }
    }

    // ── The frame ───────────────────────────────────────────────────────

    /// <summary>
    /// The graphic frame Excel recognises as a slicer: a frame holding nothing but the slicer's
    /// name, under the slicer graphic-data URI.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Excel wraps this frame in <c>mc:AlternateContent</c>, whose <c>mc:Choice</c> requires the
    /// 2010 drawing feature for a pivot slicer or the 2012 one for a table slicer, and whose
    /// <c>mc:Fallback</c> holds a rectangle explaining the slicer to a version of Excel too old to
    /// draw one. XLibur writes the frame directly instead, for two reasons.
    /// </para>
    /// <para>
    /// The first is that the wrapper protects nothing. Slicers did not exist before Excel 2010 and
    /// table slicers before Excel 2013, so the only readers the fallback is for are the ones that
    /// could not have shown the slicer anyway. Every Excel that can render a slicer resolves it from
    /// <c>a:graphicData/@uri</c> without help.
    /// </para>
    /// <para>
    /// The second is that <c>OpenXmlValidator</c> rejects <c>mc:AlternateContent</c> as the content
    /// of a <c>xdr:oneCellAnchor</c> — it accepts it inside a <c>xdr:twoCellAnchor</c>, which is
    /// what the round-trip fixture uses, so this is a quirk of the SDK's schema rather than of the
    /// format. Writing the frame directly satisfies the validator and says the same thing.
    /// </para>
    /// </remarks>
    private static OpenXmlCompositeElement BuildGraphicFrame(
        Xdr.WorksheetDrawing worksheetDrawing, XLSlicer xlSlicer)
    {
        // The SDK has no typed class for sle:slicer, so it is built as an unknown element — which
        // is also how it comes back when a file carrying one is read.
        var slicerElement = new OpenXmlUnknownElement("sle", "slicer", Slicer2010Ns);
        slicerElement.SetAttribute(new OpenXmlAttribute(string.Empty, "name", string.Empty, xlSlicer.Name));
        slicerElement.AddNamespaceDeclaration("sle", Slicer2010Ns);

        return new Xdr.GraphicFrame(
            new Xdr.NonVisualGraphicFrameProperties(
                new Xdr.NonVisualDrawingProperties { Id = NextFrameId(worksheetDrawing), Name = xlSlicer.Name },
                new Xdr.NonVisualGraphicFrameDrawingProperties()),

            // Zero, as Excel writes it: the anchor decides where the frame goes and these are
            // ignored. The element is required all the same.
            new Xdr.Transform(
                new A.Offset { X = 0, Y = 0 },
                new A.Extents { Cx = 0, Cy = 0 }),

            new A.Graphic(new A.GraphicData(slicerElement) { Uri = SlicerGraphicUri }))
        {
            Macro = string.Empty,
        };
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

    // ── Finding and converting ──────────────────────────────────────────

    /// <summary>
    /// The anchor holding the frame for the named slicer, in whichever of the three forms it uses.
    /// </summary>
    private static OpenXmlCompositeElement? FindAnchor(Xdr.WorksheetDrawing worksheetDrawing, string slicerName)
    {
        foreach (var anchor in worksheetDrawing.ChildElements.OfType<OpenXmlCompositeElement>())
        {
            if (anchor is not (Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor))
                continue;

            // The frame may be a direct child or, as Excel writes it, inside mc:AlternateContent.
            foreach (var graphicData in anchor.Descendants<A.GraphicData>())
            {
                if (graphicData.Uri?.Value != SlicerGraphicUri)
                    continue;

                if (NameOfSlicer(graphicData) == slicerName)
                    return anchor;
            }
        }

        return null;
    }

    /// <summary>
    /// The slicer name off a graphic frame's <c>sle:slicer</c> child, which the SDK has no typed
    /// class for and so deserialises as an unknown element.
    /// </summary>
    private static string? NameOfSlicer(A.GraphicData graphicData)
    {
        foreach (var child in graphicData.ChildElements)
        {
            if (child.LocalName != "slicer")
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
        int.TryParse(element?.Text, System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out var value)
            ? value
            : 0;

    private static long ReadLong(OpenXmlLeafTextElement? element) =>
        long.TryParse(element?.Text, System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out var value)
            ? value
            : 0;

    /// <summary>
    /// The inverse of <see cref="DrawingUnits.PixelsToEmu"/>, for reporting an offset a file carries.
    /// </summary>
    private static int EmuToPixels(long emu, double resolution) =>
        emu == 0 ? 0 : (int)System.Math.Round(emu * resolution / 914400d);
}
