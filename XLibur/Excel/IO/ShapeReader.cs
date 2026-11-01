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
            element.ShapeProperties is not { } shapeProperties ||
            shapeProperties.GetFirstChild<A.PresetGeometry>()?.Preset is not { HasValue: true } preset)
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
        LoadFormat(shape, shapeProperties);
        if (element.TextBody is { } body)
            LoadText(shape.Body, body);

        shape.MarkSaved(id, element.OuterXml);
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
                LoadTwoCellGeometry(shape, twoCell, from, to, element, ws, grid.Value);
                break;
        }
    }

    /// <summary>
    /// A two-cell anchor's placement is its <c>editAs</c>, as a picture's is: <c>oneCell</c> moves with
    /// its top-left cell at its own size, <c>absolute</c> stays where it is, and anything else moves
    /// and sizes with its cells. A shape given a new size or position is then written in the anchor
    /// form its placement builds.
    /// </summary>
    private static void LoadTwoCellGeometry(XLShape shape, Xdr.TwoCellAnchor anchor, Xdr.MarkerType from,
        Xdr.MarkerType to, Xdr.Shape element, XLWorksheet ws, AnchorGrid grid)
    {
        // Excel sizes a two-cell anchor from its markers; a box with nothing visible keeps the size
        // a:xfrm gives.
        var box = AnchorBox.Measure(grid, from, to);
        var extents = element.ShapeProperties?.Transform2D?.Extents;
        shape.SetLoadedSize(
            box.Width > 0 ? box.Width : ToPixels(extents?.Cx?.Value, ws.Workbook.DpiX),
            box.Height > 0 ? box.Height : ToPixels(extents?.Cy?.Value, ws.Workbook.DpiY));

        var editAs = anchor.EditAs?.Value;
        if (editAs == Xdr.EditAsValues.Absolute)
        {
            shape.SetLoadedPlacement(XLPicturePlacement.FreeFloating);
            shape.Markers[XLMarkerPosition.TopLeft] = new XLMarker(ws.Cell(1, 1), new System.Drawing.Point(box.Left, box.Top));
            return;
        }

        shape.Markers[XLMarkerPosition.TopLeft] = DrawingPartReader.LoadMarker(ws, from);
        if (editAs == Xdr.EditAsValues.OneCell)
        {
            shape.SetLoadedPlacement(XLPicturePlacement.Move);
            return;
        }

        shape.SetLoadedPlacement(XLPicturePlacement.MoveAndSize);
        shape.Markers[XLMarkerPosition.BottomRight] = DrawingPartReader.LoadMarker(ws, to);
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
            var vertical = properties.Anchor?.Value is { } anchor
                ? Find(VerticalAlignments, anchor) ?? XLTextVerticalAlignment.Top
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

        run.SetLoadedFont(
            properties.Bold?.Value,
            properties.Italic?.Value,
            properties.Underline?.Value is { } underline ? Find(Underlines, underline) : null,
            properties.Strike?.Value is { } strike ? Find(Strikes, strike) : null,
            properties.FontSize?.Value is { } size ? size / 100.0 : null,
            DrawingColorReader.ReadSolidFillColor(properties),
            properties.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
    }

    private static XLTextAlignment ToAlignment(A.TextAlignmentTypeValues? alignment) =>
        alignment is { } value ? Find(Alignments, value) ?? XLTextAlignment.Left : XLTextAlignment.Left;

    // The SDK's enum values are value types whose equality is their == operator, so they are looked
    // up in short tables rather than switched on, as DrawingColorReader's scheme colours are. A value
    // the model has no member for reads as null: not modelled, and left as the file has it.
    private static readonly (A.TextAlignmentTypeValues Xml, XLTextAlignment Model)[] Alignments =
    [
        (A.TextAlignmentTypeValues.Left, XLTextAlignment.Left),
        (A.TextAlignmentTypeValues.Center, XLTextAlignment.Center),
        (A.TextAlignmentTypeValues.Right, XLTextAlignment.Right),
        (A.TextAlignmentTypeValues.Justified, XLTextAlignment.Justified),
        (A.TextAlignmentTypeValues.Distributed, XLTextAlignment.Distributed),
    ];

    private static readonly (A.TextAnchoringTypeValues Xml, XLTextVerticalAlignment Model)[] VerticalAlignments =
    [
        (A.TextAnchoringTypeValues.Top, XLTextVerticalAlignment.Top),
        (A.TextAnchoringTypeValues.Center, XLTextVerticalAlignment.Center),
        (A.TextAnchoringTypeValues.Bottom, XLTextVerticalAlignment.Bottom),
    ];

    private static readonly (A.TextUnderlineValues Xml, XLShapeTextUnderline Model)[] Underlines =
    [
        (A.TextUnderlineValues.None, XLShapeTextUnderline.None),
        (A.TextUnderlineValues.Single, XLShapeTextUnderline.Single),
        (A.TextUnderlineValues.Double, XLShapeTextUnderline.Double),
    ];

    private static readonly (A.TextStrikeValues Xml, bool Model)[] Strikes =
    [
        (A.TextStrikeValues.SingleStrike, true),
        (A.TextStrikeValues.NoStrike, false),
    ];

    private static TModel? Find<TXml, TModel>((TXml Xml, TModel Model)[] table, TXml value)
        where TXml : struct, System.IEquatable<TXml>
        where TModel : struct
    {
        foreach (var (xml, model) in table)
        {
            if (xml.Equals(value))
                return model;
        }

        return null;
    }
}
