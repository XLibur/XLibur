using System;
using System.Linq;
using DocumentFormat.OpenXml;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads the shapes of a drawing into <see cref="IXLWorksheet.Shapes"/> (spec 15): each
/// <c>xdr:sp</c> an anchor holds directly, with a preset geometry.
/// </summary>
/// <remarks>
/// <para>
/// A shape records the drawing id it was read with, not its element: the load package is closed
/// when the load returns, and the save finds the shape again by id in the package it opens. A shape
/// whose id another loaded shape already has is given a fresh one (#859, <see cref="DrawingIdRepairs"/>).
/// </para>
/// <para>
/// Only the simple forms the model holds are read — a solid or no fill, a line's solid colour, width
/// and preset dash, the text by paragraph and run — and anything else reads as <c>null</c>, which
/// the save leaves as the file has it. A shape with a custom geometry, one in a group, a connector
/// and a graphic frame are not loaded, and stay in the file as they are.
/// </para>
/// </remarks>
internal static class ShapeReader
{
    /// <summary>Loads <paramref name="element"/>, an <c>xdr:sp</c> of <paramref name="anchor"/>, if the model can hold it.</summary>
    internal static void Load(OpenXmlCompositeElement anchor, Xdr.Shape element, Xdr.WorksheetDrawing drawing,
        XLWorksheet ws, Lazy<AnchorGrid> grid)
    {
        if (element.NonVisualShapeProperties?.NonVisualDrawingProperties is not { Id.Value: var fileId } properties ||
            element.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset is not { HasValue: true } preset)
            return;

        var shapes = (XLShapes)ws.Shapes;
        // Id 0 is no id: the model holds a shape not yet saved that way, and the save could not find
        // the shape again by it.
        var id = fileId == 0 || shapes.IsIdUsed((int)fileId)
            ? ws.DrawingIdRepairs.Repair(drawing, DrawingIdRepairs.Kind.Shape, element, fileId)
            : (int)fileId;

        var isTextBox = element.NonVisualShapeProperties.NonVisualShapeDrawingProperties?.TextBox?.Value == true;
        var shape = shapes.AddLoaded(preset.InnerText!, isTextBox, properties.Name?.Value ?? string.Empty, id);

        LoadGeometry(shape, anchor, element, ws, grid);
        LoadFormat(shape, element.ShapeProperties!);
        if (element.TextBody is { } body)
            LoadText(shape.Body, body);

        shape.MarkSaved(id);
    }

    private static void LoadGeometry(XLShape shape, OpenXmlCompositeElement anchor, Xdr.Shape element, XLWorksheet ws,
        Lazy<AnchorGrid> grid)
    {
        var dpiX = ws.Workbook.DpiX;
        var dpiY = ws.Workbook.DpiY;
        switch (anchor)
        {
            case Xdr.AbsoluteAnchor absolute:
                shape.SetLoadedPlacement(XLPicturePlacement.FreeFloating);
                shape.Markers[XLMarkerPosition.TopLeft] = new XLMarker(ws.Cell(1, 1), new System.Drawing.Point(
                    ToPixels(absolute.Position?.X?.Value, dpiX), ToPixels(absolute.Position?.Y?.Value, dpiY)));
                shape.SetLoadedSize(ToPixels(absolute.Extent?.Cx?.Value, dpiX), ToPixels(absolute.Extent?.Cy?.Value, dpiY));
                break;

            case Xdr.OneCellAnchor oneCell when oneCell.FromMarker is { } from:
                shape.SetLoadedPlacement(XLPicturePlacement.Move);
                shape.Markers[XLMarkerPosition.TopLeft] = DrawingPartReader.LoadMarker(ws, from);
                shape.SetLoadedSize(ToPixels(oneCell.Extent?.Cx?.Value, dpiX), ToPixels(oneCell.Extent?.Cy?.Value, dpiY));
                break;

            case Xdr.TwoCellAnchor twoCell when twoCell is { FromMarker: { } from, ToMarker: { } to }:
            {
                shape.SetLoadedPlacement(XLPicturePlacement.MoveAndSize);
                shape.Markers[XLMarkerPosition.TopLeft] = DrawingPartReader.LoadMarker(ws, from);
                shape.Markers[XLMarkerPosition.BottomRight] = DrawingPartReader.LoadMarker(ws, to);

                // Excel sizes a two-cell anchor from its markers; a box with nothing visible keeps
                // the size a:xfrm gives.
                var box = AnchorBox.Measure(grid.Value, from, to);
                var extents = element.ShapeProperties?.Transform2D?.Extents;
                shape.SetLoadedSize(
                    box.Width > 0 ? box.Width : ToPixels(extents?.Cx?.Value, dpiX),
                    box.Height > 0 ? box.Height : ToPixels(extents?.Cy?.Value, dpiY));
                break;
            }
        }
    }

    private static int ToPixels(long? emu, double dpi) => DrawingPartReader.ConvertFromEnglishMetricUnits(emu ?? 0, dpi);

    private static void LoadFormat(XLShape shape, Xdr.ShapeProperties shapeProperties)
    {
        var noFill = shapeProperties.GetFirstChild<A.NoFill>() is not null;
        var fillColor = DrawingColorReader.ReadSolidFillColor(shapeProperties);

        XLColor? lineColor = null;
        double? lineWidth = null;
        XLLineDashStyle? lineDash = null;
        var noLine = false;
        if (shapeProperties.GetFirstChild<A.Outline>() is { } outline)
        {
            noLine = outline.GetFirstChild<A.NoFill>() is not null;
            lineColor = DrawingColorReader.ReadSolidFillColor(outline);
            lineWidth = outline.Width?.Value is { } width ? DrawingUnits.EmuToPoints(width) : null;
            lineDash = outline.GetFirstChild<A.PresetDash>()?.Val is { HasValue: true } dash
                ? ShapePropertiesWriter.FromPresetDash(dash.Value)
                : null;
        }

        shape.SetLoadedFormat(fillColor, noFill, lineColor, lineWidth, lineDash, noLine);
    }

    private static void LoadText(XLTextBody body, Xdr.TextBody element)
    {
        if (element.BodyProperties is { } properties)
        {
            var anchor = properties.Anchor?.Value;
            var vertical = anchor == A.TextAnchoringTypeValues.Center ? XLTextVerticalAlignment.Center
                : anchor == A.TextAnchoringTypeValues.Bottom ? XLTextVerticalAlignment.Bottom
                : XLTextVerticalAlignment.Top;
            body.SetLoaded(vertical, properties.Wrap?.Value != A.TextWrappingValues.None);
        }

        var paragraphs = element.Elements<A.Paragraph>().ToList();

        // Excel writes an empty body as one paragraph with no text; the model holds it as none.
        if (paragraphs.Count == 1 && !paragraphs[0].ChildElements.Any(TextBodyPrimitives.IsRunLevel))
        {
            body.HasPlaceholderParagraph = true;
            return;
        }

        for (var i = 0; i < paragraphs.Count; i++)
            LoadParagraph(body.AddLoadedParagraph(i), paragraphs[i]);
    }

    private static void LoadParagraph(XLTextParagraph paragraph, A.Paragraph element)
    {
        paragraph.SetLoadedAlignment(ToAlignment(element.ParagraphProperties?.Alignment?.Value));

        var runIndex = 0;
        foreach (var child in element.ChildElements)
        {
            switch (child)
            {
                case A.Run run:
                    LoadFont(paragraph.AddLoadedRun(run.Text?.Text ?? string.Empty, runIndex++), run.RunProperties);
                    break;
                case A.Break:
                    paragraph.AddLoadedText("\v");
                    break;
                case A.Field field:
                    paragraph.AddLoadedText(field.Text?.Text ?? string.Empty);
                    break;
            }
        }
    }

    private static void LoadFont(XLTextRun run, A.RunProperties? properties)
    {
        if (properties is null)
            return;

        var underline = properties.Underline?.Value;
        var strike = properties.Strike?.Value;
        run.SetLoadedFont(
            properties.Bold?.Value,
            properties.Italic?.Value,
            underline == A.TextUnderlineValues.None ? XLShapeTextUnderline.None
                : underline == A.TextUnderlineValues.Single ? XLShapeTextUnderline.Single
                : underline == A.TextUnderlineValues.Double ? XLShapeTextUnderline.Double
                : null,
            strike == A.TextStrikeValues.SingleStrike ? true : strike == A.TextStrikeValues.NoStrike ? false : null,
            properties.FontSize?.Value is { } size ? size / 100.0 : null,
            DrawingColorReader.ReadSolidFillColor(properties),
            properties.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
    }

    private static XLTextAlignment ToAlignment(A.TextAlignmentTypeValues? alignment) =>
        alignment == A.TextAlignmentTypeValues.Center ? XLTextAlignment.Center
        : alignment == A.TextAlignmentTypeValues.Right ? XLTextAlignment.Right
        : alignment == A.TextAlignmentTypeValues.Justified ? XLTextAlignment.Justified
        : alignment == A.TextAlignmentTypeValues.Distributed ? XLTextAlignment.Distributed
        : XLTextAlignment.Left;
}
