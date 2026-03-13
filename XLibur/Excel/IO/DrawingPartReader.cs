using XLibur.Extensions;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using XLibur.Excel.Drawings;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.Drawing;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads drawings (pictures), VML comment shapes, and drawing style properties from worksheet parts.
/// </summary>
internal static class DrawingPartReader
{
    /// <summary>
    /// List of all VML length units and their conversion. Key is a name, value is a conversion
    /// function to EMU. See <a href="https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vml-units">documentation</a>.
    /// </summary>
    /// <remarks>
    /// OI-29500 says <em>Office also uses EMUs throughout VML as a valid unit system</em>.
    /// Relative units conversions are guesstimated by how Excel 2022 behaves for inset
    /// attribute of <c>TextBox</c> element of a note/comment. Generally speaking, Excel
    /// converts relative values to physical length (e.g. <c>px</c> to <c>pt</c>) and saves
    /// them as such. The <c>ex</c>/<c>em</c> units are not interpreted as described in the
    /// doc, but as 1/90th or an inch. The <c>%</c> seems to be always 0.
    /// </remarks>
    private static readonly Dictionary<string, Func<double, double, Emu?>> VmlLengthUnits = new()
    {
        { "in", (value, _) => Emu.From(value, AbsLengthUnit.Inch) },
        { "cm", (value, _) => Emu.From(value, AbsLengthUnit.Centimeter) },
        { "mm", (value, _) => Emu.From(value, AbsLengthUnit.Millimeter) },
        { "pt", (value, _) => Emu.From(value, AbsLengthUnit.Point) },
        { "pc", (value, _) => Emu.From(value, AbsLengthUnit.Pica) },
        { "emu", (value, _) => Emu.From(value, AbsLengthUnit.Emu) },
        { "px", (value, dpi) => Emu.From(value / dpi, AbsLengthUnit.Inch) },
        { "em", (value, _) => Emu.From(value * 72.0 / 90.0, AbsLengthUnit.Point) },
        { "ex", (value, _) => Emu.From(value * 72.0 / 90.0, AbsLengthUnit.Point) },
        { "%", (_, _) => Emu.ZeroPt },
    };

    private static readonly Dictionary<string, double> KnownPtUnits = new()
    {
        { "pt", 1.0 },
        { "in", 72.0 },
        { "mm", 72.0 / 25.4 }
    };

    internal static void LoadDrawings(WorksheetPart wsPart, XLWorksheet ws)
    {
        if (wsPart.DrawingsPart != null)
        {
            var drawingsPart = wsPart.DrawingsPart;

            foreach (var anchor in drawingsPart.WorksheetDrawing!.ChildElements)
            {
                var imgId = XLWorkbook.GetImageRelIdFromAnchor(anchor);

                //If imgId is null, we're probably dealing with a TextBox (or another shape) instead of a picture
                if (imgId == null) continue;

                // Skip external image references (e.g. URLs) — they have no embedded part.
                if (!drawingsPart.TryGetPartById(imgId, out var imagePart))
                    continue;
                using var stream = imagePart.GetStream();
                using var ms = new MemoryStream();
                stream.CopyTo(ms);
                var vsdp = XLWorkbook.GetPropertiesFromAnchor(anchor);
                var pictureName = vsdp!.Name?.Value;
                var pictureId = Convert.ToInt32(vsdp.Id!.Value);

                // Empty name is valid per ECMA-376 (xsd:string, no minLength). Excel can produce such files.
                XLPicture picture;
                if (string.IsNullOrWhiteSpace(pictureName))
                {
                    picture = (XLPicture)ws.AddPicture(ms);
                    picture.Id = pictureId;
                }
                else
                {
                    picture = (XLPicture)ws.AddPicture(ms, pictureName, pictureId);
                }
                picture!.RelId = imgId;

                var spPr = anchor.Descendants<Xdr.ShapeProperties>().First();
                picture.Placement = XLPicturePlacement.FreeFloating;

                if (spPr.Transform2D?.Extents?.Cx?.HasValue ?? false)
                    picture.Width = ConvertFromEnglishMetricUnits(spPr.Transform2D.Extents.Cx, ws.Workbook.DpiX);

                if (spPr.Transform2D?.Extents?.Cy?.HasValue ?? false)
                    picture.Height = ConvertFromEnglishMetricUnits(spPr.Transform2D.Extents.Cy, ws.Workbook.DpiY);

                if (anchor is Xdr.AbsoluteAnchor absoluteAnchor)
                {
                    picture.MoveTo(
                        ConvertFromEnglishMetricUnits(absoluteAnchor.Position!.X!.Value, ws.Workbook.DpiX),
                        ConvertFromEnglishMetricUnits(absoluteAnchor.Position!.Y!.Value, ws.Workbook.DpiY)
                    );
                }
                else if (anchor is Xdr.OneCellAnchor oneCellAnchor)
                {
                    var from = LoadMarker(ws, oneCellAnchor.FromMarker!);
                    picture.MoveTo(from.Cell, from.Offset);
                }
                else if (anchor is Xdr.TwoCellAnchor twoCellAnchor)
                {
                    var from = LoadMarker(ws, twoCellAnchor.FromMarker!);
                    var to = LoadMarker(ws, twoCellAnchor.ToMarker!);

                    if (twoCellAnchor.EditAs == null || !twoCellAnchor.EditAs.HasValue ||
                        twoCellAnchor.EditAs.Value == Xdr.EditAsValues.TwoCell)
                    {
                        picture.MoveTo(from.Cell, from.Offset, to.Cell, to.Offset);
                    }
                    else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.Absolute)
                    {
                        var shapeProperties = twoCellAnchor.Descendants<Xdr.ShapeProperties>().FirstOrDefault();
                        if (shapeProperties != null)
                        {
                            picture.MoveTo(
                                ConvertFromEnglishMetricUnits(spPr.Transform2D!.Offset!.X!, ws.Workbook.DpiX),
                                ConvertFromEnglishMetricUnits(spPr.Transform2D!.Offset!.Y!, ws.Workbook.DpiY)
                            );
                        }
                    }
                    else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.OneCell)
                    {
                        picture.MoveTo(from.Cell, from.Offset);
                    }
                }
            }
        }
    }

    internal static int ConvertFromEnglishMetricUnits(long emu, double resolution)
    {
        return Convert.ToInt32(emu * resolution / 914400);
    }

    internal static XLMarker LoadMarker(XLWorksheet ws, Xdr.MarkerType marker)
    {
        var row = Math.Min(XLHelper.MaxRowNumber, Math.Max(1, Convert.ToInt32(marker.RowId!.InnerText) + 1));
        var column = Math.Min(XLHelper.MaxColumnNumber, Math.Max(1, Convert.ToInt32(marker.ColumnId!.InnerText) + 1));
        return new XLMarker(
            ws.Cell(row, column),
            new Point(
                ConvertFromEnglishMetricUnits(Convert.ToInt32(marker.ColumnOffset!.InnerText), ws.Workbook.DpiX),
                ConvertFromEnglishMetricUnits(Convert.ToInt32(marker.RowOffset!.InnerText), ws.Workbook.DpiY)
            )
        );
    }

    internal static IList<XElement> GetCommentShapes(WorksheetPart worksheetPart)
    {
        // Cannot get this to return Vml.Shape elements
        foreach (var vmlPart in worksheetPart.VmlDrawingParts)
        {
            using var stream = vmlPart.GetStream(FileMode.Open);
            var xdoc = XDocumentExtensions.Load(stream);
            if (xdoc == null)
                continue;

            var root = xdoc.Root?.Element("xml") ?? xdoc.Root;

            var shapes = root?.Elements(XName.Get("shape", "urn:schemas-microsoft-com:vml"))
                .Where(e => new[]
                {
                    "#" + XLConstants.Comment.ShapeTypeId,
                    "#" + XLConstants.Comment.AlternateShapeTypeId
                }.Contains(e.Attribute("type")?.Value))
                .ToList();

            if (shapes != null)
                return shapes;
        }

        throw new ArgumentException("Could not load comments file");
    }

    internal static void LoadColorsAndLines<T>(IXLDrawing<T> drawing, XElement shape)
    {
        var strokeColor = shape.Attribute("strokecolor");
        if (strokeColor != null) drawing.Style.ColorsAndLines.LineColor = XLColor.FromVmlColor(strokeColor.Value);

        var strokeWeight = shape.Attribute("strokeweight");
        if (strokeWeight != null && TryGetPtValue(strokeWeight.Value, out var lineWeight))
            drawing.Style.ColorsAndLines.LineWeight = lineWeight;

        var fillColor = shape.Attribute("fillcolor");
        if (fillColor != null) drawing.Style.ColorsAndLines.FillColor = XLColor.FromVmlColor(fillColor.Value);

        var fill = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "fill");
        if (fill != null)
        {
            var opacity = fill.Attribute("opacity");
            if (opacity != null)
            {
                var opacityVal = opacity.Value;
                if (opacityVal.EndsWith("f"))
                    drawing.Style.ColorsAndLines.FillTransparency =
                        double.Parse(opacityVal.Substring(0, opacityVal.Length - 1), CultureInfo.InvariantCulture) /
                        65536.0;
                else
                    drawing.Style.ColorsAndLines.FillTransparency =
                        double.Parse(opacityVal, CultureInfo.InvariantCulture);
            }
        }

        var stroke = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "stroke");
        if (stroke != null)
        {
            var opacity = stroke.Attribute("opacity");
            if (opacity != null)
            {
                var opacityVal = opacity.Value;
                if (opacityVal.EndsWith("f"))
                    drawing.Style.ColorsAndLines.LineTransparency =
                        double.Parse(opacityVal.Substring(0, opacityVal.Length - 1), CultureInfo.InvariantCulture) /
                        65536.0;
                else
                    drawing.Style.ColorsAndLines.LineTransparency =
                        double.Parse(opacityVal, CultureInfo.InvariantCulture);
            }

            var dashStyle = stroke.Attribute("dashstyle");
            if (dashStyle != null)
            {
                var dashStyleVal = dashStyle.Value.ToLower();
                if (dashStyleVal is "1 1" or "shortdot")
                {
                    var endCap = stroke.Attribute("endcap");
                    drawing.Style.ColorsAndLines.LineDash =
                        endCap is { Value: "round" } ? XLDashStyle.RoundDot : XLDashStyle.SquareDot;
                }
                else
                {
                    drawing.Style.ColorsAndLines.LineDash = dashStyleVal switch
                    {
                        "dash" => XLDashStyle.Dash,
                        "dashdot" => XLDashStyle.DashDot,
                        "longdash" => XLDashStyle.LongDash,
                        "longdashdot" => XLDashStyle.LongDashDot,
                        "longdashdotdot" => XLDashStyle.LongDashDotDot,
                        _ => drawing.Style.ColorsAndLines.LineDash
                    };
                }
            }

            var lineStyle = stroke.Attribute("linestyle");
            if (lineStyle != null)
            {
                drawing.Style.ColorsAndLines.LineStyle = lineStyle.Value.ToLower() switch
                {
                    "single" => XLLineStyle.Single,
                    "thickbetweenthin" => XLLineStyle.ThickBetweenThin,
                    "thickthin" => XLLineStyle.ThickThin,
                    "thinthick" => XLLineStyle.ThinThick,
                    "thinthin" => XLLineStyle.ThinThin,
                    _ => drawing.Style.ColorsAndLines.LineStyle,
                };
            }
        }
    }

    internal static void LoadTextBox<T>(IXLDrawing<T> xlDrawing, XElement textBox, double dpiX, double dpiY)
    {
        var attStyle = textBox.Attribute("style");
        if (attStyle != null) LoadTextBoxStyle(xlDrawing, attStyle);

        var attInset = textBox.Attribute("inset");
        if (attInset != null) LoadTextBoxInset(xlDrawing, attInset, dpiX, dpiY);
    }

    internal static void LoadTextBoxInset<T>(IXLDrawing<T> xlDrawing, XAttribute attInset, double dpiX, double dpiY)
    {
        var split = attInset.Value.Split(',');
        xlDrawing.Style.Margins.Left = GetInsetInInches(split[0], dpiX);
        xlDrawing.Style.Margins.Top = GetInsetInInches(split[1], dpiY);
        xlDrawing.Style.Margins.Right = GetInsetInInches(split[2], dpiX);
        xlDrawing.Style.Margins.Bottom = GetInsetInInches(split[3], dpiY);
    }

    internal static double GetInsetInInches(string value, double dpi)
    {
        var unit = value.Trim();
        foreach (var (unitName, conversion) in VmlLengthUnits)
        {
            if (unit.EndsWith(unitName) && double.TryParse(unit[..^unitName.Length], NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var unitValue))
            {
                var insetEmu = conversion(unitValue, dpi) ?? Emu.ZeroPt;
                return insetEmu.To(AbsLengthUnit.Inch);
            }
        }

        // Excel treats no/unexpected unit as 0
        return 0;
    }

    internal static void LoadTextBoxStyle<T>(IXLDrawing<T> xlDrawing, XAttribute attStyle)
    {
        var style = attStyle.Value;
        var attributes = style.Split(';');
        foreach (var pair in attributes)
        {
            var split = pair.Split(':');
            if (split.Length != 2) continue;

            var attribute = split[0].Trim().ToLower();
            var attrValue = split[1].Trim();
            var isVertical = false;
            switch (attribute)
            {
                case "mso-fit-shape-to-text": xlDrawing.Style.Size.SetAutomaticSize(attrValue.Equals("t")); break;
                case "mso-layout-flow-alt":
                    if (attrValue.Equals("bottom-to-top"))
                        xlDrawing.Style.Alignment.SetOrientation(XLDrawingTextOrientation.BottomToTop);
                    else if (attrValue.Equals("top-to-bottom"))
                        xlDrawing.Style.Alignment.SetOrientation(XLDrawingTextOrientation.Vertical);
                    break;

                case "layout-flow": isVertical = attrValue.Equals("vertical"); break;
                case "mso-direction-alt":
                    if (attrValue == "auto") xlDrawing.Style.Alignment.Direction = XLDrawingTextDirection.Context;
                    break;
                case "direction":
                    if (attrValue == "RTL") xlDrawing.Style.Alignment.Direction = XLDrawingTextDirection.RightToLeft;
                    break;
            }

            if (isVertical && xlDrawing.Style.Alignment.Orientation == XLDrawingTextOrientation.LeftToRight)
                xlDrawing.Style.Alignment.Orientation = XLDrawingTextOrientation.TopToBottom;
        }
    }

    internal static void LoadClientData<T>(IXLDrawing<T> drawing, XElement clientData)
    {
        var anchor = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Anchor");
        if (anchor != null) LoadClientDataAnchor(drawing, anchor);

        LoadDrawingPositioning(drawing, clientData);
        LoadDrawingProtection(drawing, clientData);

        var visible = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Visible");
        drawing.Visible = visible != null &&
                          (string.IsNullOrEmpty(visible.Value) ||
                           visible.Value.StartsWith("t", StringComparison.OrdinalIgnoreCase));

        LoadDrawingHAlignment(drawing, clientData);
        LoadDrawingVAlignment(drawing, clientData);
    }

    internal static void LoadDrawingHAlignment<T>(IXLDrawing<T> drawing, XElement clientData)
    {
        var textHAlign = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "TextHAlign");
        if (textHAlign != null)
            drawing.Style.Alignment.Horizontal =
                (XLDrawingHorizontalAlignment)Enum.Parse(typeof(XLDrawingHorizontalAlignment),
                    textHAlign.Value.ToProper());
    }

    internal static void LoadDrawingVAlignment<T>(IXLDrawing<T> drawing, XElement clientData)
    {
        var textVAlign = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "TextVAlign");
        if (textVAlign != null)
            drawing.Style.Alignment.Vertical =
                (XLDrawingVerticalAlignment)Enum.Parse(typeof(XLDrawingVerticalAlignment), textVAlign.Value.ToProper());
    }

    internal static void LoadDrawingProtection<T>(IXLDrawing<T> drawing, XElement clientData)
    {
        var lockedElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Locked");
        var lockTextElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "LockText");
        var locked = lockedElement != null && lockedElement.Value.ToLower() == "true";
        var lockText = lockTextElement != null && lockTextElement.Value.ToLower() == "true";
        drawing.Style.Protection.Locked = locked;
        drawing.Style.Protection.LockText = lockText;
    }

    internal static void LoadDrawingPositioning<T>(IXLDrawing<T> drawing, XElement clientData)
    {
        var moveWithCellsElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "MoveWithCells");
        var sizeWithCellsElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "SizeWithCells");
        var moveWithCells = !(moveWithCellsElement != null && moveWithCellsElement.Value.ToLower() == "true");
        var sizeWithCells = !(sizeWithCellsElement != null && sizeWithCellsElement.Value.ToLower() == "true");
        if (moveWithCells && !sizeWithCells)
            drawing.Style.Properties.Positioning = XLDrawingAnchor.MoveWithCells;
        else if (moveWithCells && sizeWithCells)
            drawing.Style.Properties.Positioning = XLDrawingAnchor.MoveAndSizeWithCells;
        else
            drawing.Style.Properties.Positioning = XLDrawingAnchor.Absolute;
    }

    internal static void LoadClientDataAnchor<T>(IXLDrawing<T> drawing, XElement anchor)
    {
        var location = anchor.Value.Split(',');
        drawing.Position.Column = int.Parse(location[0]) + 1;
        drawing.Position.ColumnOffset = double.Parse(location[1], CultureInfo.InvariantCulture) / 7.5;
        drawing.Position.Row = int.Parse(location[2]) + 1;
        drawing.Position.RowOffset = double.Parse(location[3], CultureInfo.InvariantCulture);
    }

    internal static void LoadShapeProperties<T>(IXLDrawing<T> xlDrawing, XElement shape)
    {
        if (shape.Attribute("style") == null)
            return;

        foreach (var attributePair in shape.Attribute("style")!.Value.Split(';'))
        {
            var split = attributePair.Split(':');
            if (split.Length != 2) continue;

            var attribute = split[0].Trim().ToLower();
            var attrValue = split[1].Trim();

            switch (attribute)
            {
                case "visibility":
                    xlDrawing.Visible = string.Equals("visible", attrValue, StringComparison.OrdinalIgnoreCase); break;
                case "width":
                    if (TryGetPtValue(attrValue, out var ptWidth))
                    {
                        xlDrawing.Style.Size.Width = ptWidth / 7.5;
                    }

                    break;

                case "height":
                    if (TryGetPtValue(attrValue, out var ptHeight))
                    {
                        xlDrawing.Style.Size.Height = ptHeight;
                    }

                    break;

                case "z-index":
                    if (int.TryParse(attrValue, out var zOrder))
                    {
                        xlDrawing.ZOrder = zOrder;
                    }

                    break;
            }
        }
    }

    internal static bool TryGetPtValue(string value, out double result)
    {
        var knownUnit = KnownPtUnits.FirstOrDefault(ku => value.Contains(ku.Key));

        if (knownUnit.Key == null)
            return double.TryParse(value, out result);

        value = value.Replace(knownUnit.Key, string.Empty);

        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
        {
            result *= knownUnit.Value;
            return true;
        }

        result = 0d;
        return false;
    }

    internal static string GetTableColumnName(string name)
    {
        return name.Replace("_x000a_", Environment.NewLine).Replace("_x005f_x000a_", "_x000a_");
    }
}
