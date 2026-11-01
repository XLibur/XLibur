using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using static XLibur.Excel.XLWorkbook;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes a sheet's shapes into its drawing part (spec 15): a new shape into a fresh anchor, a
/// loaded one by <see cref="ShapePatcher"/> into the <c>xdr:sp</c> it was read from, and a deleted
/// one out with its anchor.
/// </summary>
/// <remarks>
/// A new shape starts from what Excel itself writes for a shape it adds, captured from the fixture
/// <c>Other/Shapes/ExcelShapes.xlsx</c> (Excel 365 over COM): for a text box a white, theme-lined
/// rectangle with a dark theme font; for a preset shape an <c>xdr:style</c> that takes its fill, line
/// and font from the theme's first accent. Whatever was assigned is then written over that, through
/// the same writes a loaded shape gets, so there is one way to write a fill, not two.
/// </remarks>
internal static class ShapeWriter
{
    /// <summary>
    /// Removes the anchor of every shape deleted since the load or the last save, from a drawing the
    /// save package has, when <paramref name="ids"/> is given.
    /// </summary>
    /// <remarks>
    /// The deletions are forgotten only once the save has succeeded: a save that fails leaves the
    /// next one copying the old source, which still holds the shapes.
    /// </remarks>
    internal static void RemoveDeleted(XLShapes shapes, DrawingIds? ids, SaveContext context)
    {
        var deleted = shapes.Deleted.ToList();
        if (ids is not null)
        {
            foreach (var id in deleted)
                ids.FindShape(id)?.Anchor.Remove();
        }

        context.KeepAfterSave.Add(() =>
        {
            foreach (var id in deleted)
                shapes.Deleted.Remove(id);
        });
    }

    /// <summary>
    /// Writes every shape that changed, or every shape when <paramref name="writeAll"/> is set, as
    /// for a drawing the save package does not have yet. Each shape takes the drawing id it was
    /// written under once the save has succeeded.
    /// </summary>
    internal static void WriteShapes(DrawingsPart drawingsPart, XLShapes shapes, SaveContext context, DrawingIds ids,
        bool writeAll)
    {
        var drawing = drawingsPart.WorksheetDrawing!;
        DrawingPartScaffold.EnsureNamespaces(drawing);

        foreach (var shape in shapes)
        {
            if (!writeAll && !shape.HasUnsavedChanges)
                continue;

            uint id;
            if (shape.IsInFile && !writeAll)
            {
                // A shape the file holds is patched or left alone, never written a second time: one
                // the drawing no longer has under its id keeps what the drawing has.
                if (ids.FindShape(shape.Id) is not { } existing)
                    continue;

                ShapePatcher.Patch(existing.Anchor, existing.Shape, shape);
                id = (uint)shape.Id;
            }
            else
            {
                id = ids.NextDrawingId();
                drawing.Append(DrawingAnchorFactory.Create(shape.Placement, GeometryOf(shape), NewShapeElement(shape, id)));
            }

            var saved = shape;
            context.KeepAfterSave.Add(() => saved.MarkSaved((int)id));
        }
    }

    /// <summary>Where the model puts the shape. See <see cref="DrawingAnchorFactory"/> for the A1 fallbacks.</summary>
    internal static DrawingAnchorGeometry GeometryOf(XLShape shape) => new()
    {
        Worksheet = shape.Worksheet,
        LeftPx = shape.Left,
        TopPx = shape.Top,
        WidthPx = shape.Width,
        HeightPx = shape.Height,
        FromMarker = shape.Markers[XLMarkerPosition.TopLeft],
        ToMarker = shape.Markers[XLMarkerPosition.BottomRight],
    };

    /// <summary>The <c>xdr:sp</c> of a new shape: Excel's creation defaults, then the assigned formatting and the text.</summary>
    private static Xdr.Shape NewShapeElement(XLShape shape, uint id)
    {
        var wb = shape.Worksheet.Workbook;
        var shapeProperties = new Xdr.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = 0, Y = 0 },
                new A.Extents
                {
                    Cx = DrawingUnits.PixelsToEmu(shape.Width, wb.DpiX),
                    Cy = DrawingUnits.PixelsToEmu(shape.Height, wb.DpiY),
                }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = new A.ShapeTypeValues(shape.PresetGeometry) });

        if (shape.IsTextBox)
        {
            shapeProperties.Append(
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Light1 }),
                new A.Outline(new A.SolidFill(new A.SchemeColor(new A.Shade { Val = 50000 }) { Val = A.SchemeColorValues.Light1 }))
                {
                    Width = 9525,
                    CompoundLineType = A.CompoundLineValues.Single,
                });
        }

        ShapeFormatWriter.Write(shapeProperties, shape);

        var body = new Xdr.TextBody(
            new A.BodyProperties
            {
                VerticalOverflow = A.TextVerticalOverflowValues.Clip,
                HorizontalOverflow = A.TextHorizontalOverflowValues.Clip,
                Vertical = shape.IsTextBox ? A.TextVerticalValues.Horizontal : null,
                RightToLeftColumns = false,
                Anchor = A.TextAnchoringTypeValues.Top,
            },
            new A.ListStyle());
        ShapeTextWriter.Write(body, shape.Body, alignNewParagraphs: !shape.IsTextBox);

        return new Xdr.Shape(
            new Xdr.NonVisualShapeProperties(
                new Xdr.NonVisualDrawingProperties { Id = id, Name = shape.Name },
                new Xdr.NonVisualShapeDrawingProperties { TextBox = shape.IsTextBox ? true : null }),
            shapeProperties,
            shape.IsTextBox ? TextBoxStyle() : PresetStyle(shape.PresetGeometry),
            body)
        {
            Macro = string.Empty,
            TextLink = string.Empty,
        };
    }

    /// <summary>The style Excel gives a text box: nothing from the theme but its minor font in dark 1.</summary>
    private static Xdr.ShapeStyle TextBoxStyle() => new(
        new A.LineReference(Black()) { Index = 0 },
        new A.FillReference(Black()) { Index = 0 },
        new A.EffectReference(Black()) { Index = 0 },
        new A.FontReference(new A.SchemeColor { Val = A.SchemeColorValues.Dark1 }) { Index = A.FontCollectionIndexValues.Minor });

    /// <summary>
    /// The style Excel gives a preset shape: the theme's first accent for the fill and a darker
    /// shade of it for the line, with light text. An open shape, a line or an arc, has no fill and
    /// dark text instead.
    /// </summary>
    private static Xdr.ShapeStyle PresetStyle(string preset)
    {
        if (IsOpen(preset))
        {
            return new Xdr.ShapeStyle(
                new A.LineReference(Accent1()) { Index = 2 },
                new A.FillReference(Accent1()) { Index = 0 },
                new A.EffectReference(Accent1()) { Index = 1 },
                new A.FontReference(new A.SchemeColor { Val = A.SchemeColorValues.Text1 }) { Index = A.FontCollectionIndexValues.Minor });
        }

        return new Xdr.ShapeStyle(
            new A.LineReference(new A.SchemeColor(new A.Shade { Val = 15000 }) { Val = A.SchemeColorValues.Accent1 }) { Index = 2 },
            new A.FillReference(Accent1()) { Index = 1 },
            new A.EffectReference(Accent1()) { Index = 0 },
            new A.FontReference(new A.SchemeColor { Val = A.SchemeColorValues.Light1 }) { Index = A.FontCollectionIndexValues.Minor });
    }

    private static readonly HashSet<string> OpenPresets = ["line", "arc", "bracketPair", "bracePair", "leftBracket",
        "rightBracket", "leftBrace", "rightBrace", "lineInv"];

    private static bool IsOpen(string preset) => OpenPresets.Contains(preset);

    private static A.SchemeColor Accent1() => new() { Val = A.SchemeColorValues.Accent1 };

    private static A.RgbColorModelPercentage Black() => new() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };
}
