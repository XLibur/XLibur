using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using XLibur.Extensions;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads drawings (pictures, and shapes through <see cref="ShapeReader"/>), VML comment shapes, and
/// drawing style properties from worksheet parts.
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

    internal static void LoadDrawings(WorksheetPart wsPart, XLWorksheet ws, LoadContext context)
    {
        if (wsPart.DrawingsPart == null)
            return;

        var drawingsPart = wsPart.DrawingsPart;

        // One grid for every two-cell anchor on the sheet, built only if the sheet has one.
        var columnWidths = context.SheetColumnWidths;
        var grid = new Lazy<AnchorGrid>(
            () => new AnchorGrid(ws, columnWidths, ws.Workbook.MaxDigitWidth),
            LazyThreadSafetyMode.None);

        var layout = GridLayout.OnFirstUse(ws);

        // The names of the shapes that do not load as pictures stay taken (issue #796).
        var drawing = drawingsPart.WorksheetDrawing!;
        ((XLPictures)ws.Pictures).LoadDrawing(ShapeNames(drawing), () =>
        {
            foreach (var anchor in drawing.ChildElements)
                LoadDrawingAnchor(anchor, drawingsPart, ws, grid, layout);
        });
    }

    /// <summary>The name of each shape of the drawing, at any depth, once per shape.</summary>
    /// <remarks>
    /// The branches of an <c>mc:AlternateContent</c> are one drawing in two forms, and Excel gives
    /// them one id and one name: a 3D model's frame and its fallback picture. Such a drawing is one
    /// shape, so its name is given once, and the picture loaded from the fallback accounts for it.
    /// </remarks>
    private static IEnumerable<string> ShapeNames(Xdr.WorksheetDrawing drawing)
    {
        var forms = new HashSet<(AlternateContent Wrapper, uint Id)>();
        foreach (var properties in drawing.Descendants<Xdr.NonVisualDrawingProperties>())
        {
            if (properties.Name?.Value is not { Length: > 0 } name)
                continue;

            if (properties.Ancestors<AlternateContent>().FirstOrDefault() is { } wrapper &&
                properties.Id?.Value is { } id &&
                !forms.Add((wrapper, id)))
                continue;

            yield return name;
        }
    }

    private static void LoadDrawingAnchor(OpenXmlElement anchor, DrawingsPart drawingsPart, XLWorksheet ws,
        Lazy<AnchorGrid> grid, Lazy<GridLayout> layout)
    {
        // Pictures nested inside a top-level group shape (xdr:grpSp) are laid out in the
        // group's child coordinate space. Load each of them as an editable picture that
        // remembers its group transform, so on save we can update the xdr:pic in place and
        // leave the rest of the group (sibling pictures, connectors, shapes) intact.
        var group = anchor.Elements<Xdr.GroupShape>().FirstOrDefault();
        if (group != null)
        {
            LoadGroupedPictures(group, drawingsPart, ws);
            return;
        }

        if (anchor is OpenXmlCompositeElement composite && composite.GetFirstChild<Xdr.Shape>() is { } shape)
        {
            ShapeReader.Load(composite, shape, drawingsPart.WorksheetDrawing!, ws, grid);
            return;
        }

        // A picture deeper in the anchor than its wrapper's branches, as in a group inside the
        // wrapper, is left as the file has it (#787).
        var first = anchor.Descendants<Xdr.Picture>().FirstOrDefault();
        if (first is not null && !DrawingAnchorContent.IsAnchorPicture(first, anchor))
            return;

        var imgId = XLWorkbook.GetImageRelIdFromAnchor(anchor);

        //If imgId is null, we're probably dealing with a TextBox (or another shape) instead of a picture
        if (imgId == null) return;

        // Skip external image references (e.g. URLs) — they have no embedded part.
        if (!drawingsPart.TryGetPartById(imgId, out var imagePart))
            return;

        var picture = LoadPictureFromAnchor(anchor, first, drawingsPart.WorksheetDrawing!, imgId, imagePart, ws);
        if (first is not null)
            LoadPictureStyle(picture, first);

        LoadPictureTransform(picture, anchor, ws);
        LoadPicturePlacement(picture, anchor, ws, grid, layout);
        picture.MarkSaved();
    }

    internal static int ConvertFromEnglishMetricUnits(long emu, double resolution)
    {
        return Convert.ToInt32(emu * resolution / 914400);
    }

    /// <summary>
    /// Load every picture contained in a group shape, at any nesting depth. Each picture is added to
    /// the worksheet with its sheet-space size (its child-space extent scaled by the composed group
    /// transform) and tagged with its <see cref="XLPictureGroup"/> so the writer can update it in
    /// place. Non-picture children (connectors, shapes) are left untouched and preserved verbatim.
    /// </summary>
    private static void LoadGroupedPictures(Xdr.GroupShape group, DrawingsPart drawingsPart, XLWorksheet ws)
    {
        // The parent coordinate space of the top-level group is the sheet itself: identity transform.
        LoadGroupRecursive(group, drawingsPart, ws, parentScaleX: 1.0, parentScaleY: 1.0,
            parentOffsetX: 0.0, parentOffsetY: 0.0);
    }

    /// <summary>
    /// Recursively load the pictures of <paramref name="group"/> and its nested groups. The
    /// <c>parent…</c> arguments are the affine transform mapping this group's <em>parent</em> child
    /// coordinate space to the sheet; this group's own transform is composed in so a picture's sheet
    /// geometry is <c>offset + child · scale</c>, where scale is the product of every enclosing
    /// group's <c>ext/chExt</c> ratio.
    /// </summary>
    private static void LoadGroupRecursive(Xdr.GroupShape group, DrawingsPart drawingsPart, XLWorksheet ws,
        double parentScaleX, double parentScaleY, double parentOffsetX, double parentOffsetY)
    {
        var xfrm = group.GroupShapeProperties?.TransformGroup;
        if (xfrm == null)
            return; // No transform — cannot map child coordinates; leave this group preserved as-is.

        var gOffX = xfrm.Offset?.X?.Value ?? 0;
        var gOffY = xfrm.Offset?.Y?.Value ?? 0;
        var gExtCx = xfrm.Extents?.Cx?.Value ?? 0;
        var gExtCy = xfrm.Extents?.Cy?.Value ?? 0;
        var chOffX = xfrm.ChildOffset?.X?.Value ?? 0;
        var chOffY = xfrm.ChildOffset?.Y?.Value ?? 0;
        var chExtCx = xfrm.ChildExtents?.Cx?.Value ?? 0;
        var chExtCy = xfrm.ChildExtents?.Cy?.Value ?? 0;

        // This group's own transform maps its child space to its parent's child space:
        //   parentSpace = (off - chOff·levelScale) + child·levelScale
        var levelScaleX = chExtCx == 0 ? 1.0 : (double)gExtCx / chExtCx;
        var levelScaleY = chExtCy == 0 ? 1.0 : (double)gExtCy / chExtCy;
        var levelOffsetX = gOffX - chOffX * levelScaleX;
        var levelOffsetY = gOffY - chOffY * levelScaleY;

        // Compose with the parent transform to get child→sheet: scale multiplies, offset accumulates.
        var scaleX = parentScaleX * levelScaleX;
        var scaleY = parentScaleY * levelScaleY;
        var offsetX = parentOffsetX + levelOffsetX * parentScaleX;
        var offsetY = parentOffsetY + levelOffsetY * parentScaleY;

        var groupId = group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
        var groupKey = ((XLPictures)ws.Pictures).AllocateGroupKey();

        foreach (var pic in group.Elements<Xdr.Picture>())
            LoadGroupedPicture(pic, drawingsPart, ws, scaleX, scaleY, offsetX, offsetY, groupId, groupKey);

        foreach (var nested in group.Elements<Xdr.GroupShape>())
            LoadGroupRecursive(nested, drawingsPart, ws, scaleX, scaleY, offsetX, offsetY);
    }

#pragma warning disable S107 // One parameter per drawing anchor attribute
    private static void LoadGroupedPicture(Xdr.Picture pic, DrawingsPart drawingsPart, XLWorksheet ws,
        double scaleX, double scaleY, double offsetX, double offsetY, uint? groupId, long groupKey)
#pragma warning restore S107
    {
        var embed = pic.BlipFill?.Blip?.Embed?.Value;
        if (embed == null)
            return;

        // External image references (e.g. URLs) have no embedded part.
        if (!drawingsPart.TryGetPartById(embed, out var imagePart))
            return;

        var transform = pic.ShapeProperties?.Transform2D;
        var picOffX = transform?.Offset?.X?.Value ?? 0;
        var picOffY = transform?.Offset?.Y?.Value ?? 0;
        var sheetEmuCx = (long)Math.Round((transform?.Extents?.Cx?.Value ?? 0) * scaleX);
        var sheetEmuCy = (long)Math.Round((transform?.Extents?.Cy?.Value ?? 0) * scaleY);
        var sheetEmuX = (long)Math.Round(offsetX + picOffX * scaleX);
        var sheetEmuY = (long)Math.Round(offsetY + picOffY * scaleY);

        var xlPicture = AddGroupedPicture(pic, drawingsPart.WorksheetDrawing!, embed, imagePart, ws);

        // FreeFloating lets Width/Height/Left/Top be edited; placement does not drive the anchor for
        // grouped pictures because they are written back in place.
        xlPicture.Placement = XLPicturePlacement.FreeFloating;
        var widthPx = ConvertFromEnglishMetricUnits(sheetEmuCx, ws.Workbook.DpiX);
        var heightPx = ConvertFromEnglishMetricUnits(sheetEmuCy, ws.Workbook.DpiY);
        xlPicture.Width = widthPx;
        xlPicture.Height = heightPx;

        // Expose the picture's true sheet position via an A1-relative marker, so Left/Top read and
        // set sheet-space coordinates (consistent with FreeFloating placement).
        var leftPx = ConvertFromEnglishMetricUnits(sheetEmuX, ws.Workbook.DpiX);
        var topPx = ConvertFromEnglishMetricUnits(sheetEmuY, ws.Workbook.DpiY);
        xlPicture.Markers[XLMarkerPosition.TopLeft] = new XLMarker(ws.Cell(1, 1), new System.Drawing.Point(leftPx, topPx));

        xlPicture.GroupInfo = new XLPictureGroup
        {
            ScaleX = scaleX,
            ScaleY = scaleY,
            OffsetX = offsetX,
            OffsetY = offsetY,
            LoadedWidthPx = widthPx,
            LoadedHeightPx = heightPx,
            LoadedLeftPx = leftPx,
            LoadedTopPx = topPx,
            GroupId = groupId,
            GroupKey = groupKey,
        };
        LoadPictureStyle(xlPicture, pic);
        xlPicture.MarkSaved();
    }

    private static XLPicture AddGroupedPicture(Xdr.Picture pic, Xdr.WorksheetDrawing drawing, string embed,
        OpenXmlPart imagePart, XLWorksheet ws)
    {
        using var stream = imagePart.GetStream();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);

        var nvProps = pic.NonVisualPictureProperties?.NonVisualDrawingProperties;
        var pictureName = nvProps?.Name?.Value;
        var pictureId = PictureIdOf(ws, drawing, pic, Convert.ToInt32(nvProps?.Id?.Value ?? 0));

        XLPicture picture;
        if (string.IsNullOrWhiteSpace(pictureName))
        {
            picture = (XLPicture)ws.AddPicture(ms);
            if (pictureId > 0)
                picture.Id = pictureId;
        }
        else
        {
            picture = (XLPicture)ws.AddPicture(ms, pictureName, pictureId);
        }

        picture.RelId = embed;
        return picture;
    }

    /// <summary>
    /// Reads the picture's styling, and keeps its <c>xdr:pic</c>, which a copy of the picture starts
    /// from so the copy keeps what XLibur does not model.
    /// </summary>
    private static void LoadPictureStyle(XLPicture picture, Xdr.Picture element)
    {
        PictureStyleReader.Read(element, picture.Style);
        picture.LoadedPictureXml = element.OuterXml;
    }

    internal static XLMarker LoadMarker(XLWorksheet ws, Xdr.MarkerType marker)
    {
        var row = Math.Min(XLHelper.MaxRowNumber, Math.Max(1, Convert.ToInt32(marker.RowId!.InnerText) + 1));
        var column = Math.Min(XLHelper.MaxColumnNumber, Math.Max(1, Convert.ToInt32(marker.ColumnId!.InnerText) + 1));
        return new XLMarker(
            ws.Cell(row, column),
            new System.Drawing.Point(
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
        LoadShapeColors(drawing, shape);

        var fill = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "fill");
        if (fill != null)
            LoadFillOpacity(drawing, fill);

        var stroke = shape.Elements().FirstOrDefault(e => e.Name.LocalName == "stroke");
        if (stroke != null)
            LoadStrokeProperties(drawing, stroke);
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

    private static void LoadTextBoxStyle<T>(IXLDrawing<T> xlDrawing, XAttribute attStyle)
    {
        var style = attStyle.Value;
        var attributes = style.Split(';');
        foreach (var pair in attributes)
            LoadTextBoxStyleProperty(xlDrawing, pair);
    }

    private static void LoadTextBoxStyleProperty<T>(IXLDrawing<T> xlDrawing, string pair)
    {
        var split = pair.Split(':');
        if (split.Length != 2) return;

        var attribute = split[0].Trim().ToLower();
        var attrValue = split[1].Trim();
        var isVertical = false;
        switch (attribute)
        {
            case "mso-fit-shape-to-text": xlDrawing.Style.Size.SetAutomaticSize(attrValue.Equals("t")); break;
            case "mso-layout-flow-alt":
                LoadLayoutFlowAlt(xlDrawing, attrValue);
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
                Enum.Parse<XLDrawingHorizontalAlignment>(textHAlign.Value.ToProper());
    }

    internal static void LoadDrawingVAlignment<T>(IXLDrawing<T> drawing, XElement clientData)
    {
        var textVAlign = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "TextVAlign");
        if (textVAlign != null)
            drawing.Style.Alignment.Vertical =
                Enum.Parse<XLDrawingVerticalAlignment>(textVAlign.Value.ToProper());
    }

    internal static void LoadDrawingProtection<T>(IXLDrawing<T> drawing, XElement clientData)
    {
        var lockedElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "Locked");
        var lockTextElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "LockText");
        var locked = lockedElement != null && lockedElement.Value.Equals("true", StringComparison.OrdinalIgnoreCase);
        var lockText = lockTextElement != null && lockTextElement.Value.Equals("true", StringComparison.OrdinalIgnoreCase);
        drawing.Style.Protection.Locked = locked;
        drawing.Style.Protection.LockText = lockText;
    }

    internal static void LoadDrawingPositioning<T>(IXLDrawing<T> drawing, XElement clientData)
    {
        var moveWithCellsElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "MoveWithCells");
        var sizeWithCellsElement = clientData.Elements().FirstOrDefault(e => e.Name.LocalName == "SizeWithCells");
        var moveWithCells = !(moveWithCellsElement != null && moveWithCellsElement.Value.Equals("true", StringComparison.OrdinalIgnoreCase));
        var sizeWithCells = !(sizeWithCellsElement != null && sizeWithCellsElement.Value.Equals("true", StringComparison.OrdinalIgnoreCase));
        if (!moveWithCells)
            drawing.Style.Properties.Positioning = XLDrawingAnchor.Absolute;
        else if (sizeWithCells)
            drawing.Style.Properties.Positioning = XLDrawingAnchor.MoveAndSizeWithCells;
        else
            drawing.Style.Properties.Positioning = XLDrawingAnchor.MoveWithCells;
    }

    private static void LoadClientDataAnchor<T>(IXLDrawing<T> drawing, XElement anchor)
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

            LoadShapeProperty(xlDrawing, attribute, attrValue);
        }
    }

    private static void LoadShapeProperty<T>(IXLDrawing<T> xlDrawing, string attribute, string attrValue)
    {
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

    private static XLPicture LoadPictureFromAnchor(
        OpenXmlElement anchor, Xdr.Picture? element, Xdr.WorksheetDrawing drawing, string imgId,
        OpenXmlPart imagePart, XLWorksheet ws)
    {
        using var stream = imagePart.GetStream();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        var vsdp = XLWorkbook.GetPropertiesFromAnchor(anchor);
        var pictureName = vsdp!.Name?.Value;
        var pictureId = element is null
            ? Convert.ToInt32(vsdp.Id!.Value)
            : PictureIdOf(ws, drawing, element, Convert.ToInt32(vsdp.Id!.Value));

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

        picture.RelId = imgId;
        return picture;
    }

    /// <summary>
    /// The id a loaded picture takes: the file's, or a fresh one when another picture of the sheet
    /// already has it (#859), id 0 included. The first save writes the fresh one into the drawing.
    /// </summary>
    private static int PictureIdOf(XLWorksheet ws, Xdr.WorksheetDrawing drawing, Xdr.Picture element, int fileId) =>
        ((XLPictures)ws.Pictures).IsIdUsed(fileId)
            ? ws.DrawingIdRepairs.Repair(drawing, DrawingIdRepairs.Kind.Picture, element, (uint)fileId)
            : fileId;

    private static void LoadPictureTransform(XLPicture picture,
        OpenXmlElement anchor, XLWorksheet ws)
    {
        var spPr = anchor.Descendants<Xdr.ShapeProperties>().First();
        picture.Placement = XLPicturePlacement.FreeFloating;

        if (spPr.Transform2D?.Extents?.Cx?.HasValue ?? false)
            picture.Width = ConvertFromEnglishMetricUnits(spPr.Transform2D.Extents.Cx, ws.Workbook.DpiX);

        if (spPr.Transform2D?.Extents?.Cy?.HasValue ?? false)
            picture.Height = ConvertFromEnglishMetricUnits(spPr.Transform2D.Extents.Cy, ws.Workbook.DpiY);
    }

    private static void LoadPicturePlacement(XLPicture picture,
        OpenXmlElement anchor, XLWorksheet ws, Lazy<AnchorGrid> grid, Lazy<GridLayout> layout)
    {
        if (anchor is Xdr.AbsoluteAnchor absoluteAnchor)
        {
            picture.MoveTo(
                ConvertFromEnglishMetricUnits(absoluteAnchor.Position!.X!.Value, ws.Workbook.DpiX),
                ConvertFromEnglishMetricUnits(absoluteAnchor.Position.Y!.Value, ws.Workbook.DpiY)
            );
        }
        else if (anchor is Xdr.OneCellAnchor oneCellAnchor)
        {
            var from = LoadMarker(ws, oneCellAnchor.FromMarker!);
            picture.MoveTo(from.Cell, from.Offset);
        }
        else if (anchor is Xdr.TwoCellAnchor twoCellAnchor)
        {
            LoadTwoCellAnchorPlacement(picture, twoCellAnchor, ws, grid.Value, layout);
        }
    }

    private static void LoadTwoCellAnchorPlacement(XLPicture picture, Xdr.TwoCellAnchor twoCellAnchor, XLWorksheet ws,
        AnchorGrid grid, Lazy<GridLayout> layout)
    {
        var from = LoadMarker(ws, twoCellAnchor.FromMarker!);
        var to = LoadMarker(ws, twoCellAnchor.ToMarker!);

        // Excel sizes every two-cell anchor from its markers, whatever editAs says, and ignores
        // a:xfrm, which other writers leave at offset 0,0 and the native size (#782). The size is
        // set while the picture is still free-floating, before a placement below forbids it. A box
        // with no visible width or height (all of it hidden, or to before from) keeps the xfrm size
        // read before, which a later edit can still write back.
        var box = AnchorBox.Measure(grid, twoCellAnchor.FromMarker!, twoCellAnchor.ToMarker!);
        if (box.Width > 0)
            picture.Width = box.Width;

        if (box.Height > 0)
            picture.Height = box.Height;

        if (twoCellAnchor.EditAs == null || !twoCellAnchor.EditAs.HasValue ||
            twoCellAnchor.EditAs.Value == Xdr.EditAsValues.TwoCell)
        {
            // MoveAndSize keeps both markers, so the model says this anchor back by itself.
            picture.MoveTo(from.Cell, from.Offset, to.Cell, to.Offset);
            return;
        }
        else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.Absolute)
        {
            picture.MoveTo(box.Left, box.Top);
        }
        else if (twoCellAnchor.EditAs.Value == Xdr.EditAsValues.OneCell)
        {
            picture.MoveTo(from.Cell, from.Offset);
        }
        else
        {
            return;
        }

        // Neither placement can say this anchor back: Excel sizes it from the markers dropped
        // above, not from the xfrm the model was given. Set last, because placing the picture
        // drops it.
        picture.KeptAnchor = new XLKeptAnchor(layout.Value, picture.Placement, from, to);
    }

    private static void LoadShapeColors<T>(IXLDrawing<T> drawing, XElement shape)
    {
        var strokeColor = shape.Attribute("strokecolor");
        if (strokeColor != null) drawing.Style.ColorsAndLines.LineColor = XLColor.FromVmlColor(strokeColor.Value);

        var strokeWeight = shape.Attribute("strokeweight");
        if (strokeWeight != null && TryGetPtValue(strokeWeight.Value, out var lineWeight))
            drawing.Style.ColorsAndLines.LineWeight = lineWeight;

        var fillColor = shape.Attribute("fillcolor");
        if (fillColor != null) drawing.Style.ColorsAndLines.FillColor = XLColor.FromVmlColor(fillColor.Value);
    }

    private static void LoadFillOpacity<T>(IXLDrawing<T> drawing, XElement fill)
    {
        var opacity = fill.Attribute("opacity");
        if (opacity != null)
        {
            var opacityVal = opacity.Value;
            if (opacityVal.EndsWith('f'))
                drawing.Style.ColorsAndLines.FillTransparency =
                    double.Parse(opacityVal.AsSpan(0, opacityVal.Length - 1), provider: CultureInfo.InvariantCulture) /
                    65536.0;
            else
                drawing.Style.ColorsAndLines.FillTransparency =
                    double.Parse(opacityVal, CultureInfo.InvariantCulture);
        }
    }

    private static void LoadStrokeProperties<T>(IXLDrawing<T> drawing, XElement stroke)
    {
        LoadStrokeOpacity(drawing, stroke);
        LoadStrokeDashStyle(drawing, stroke);
        LoadStrokeLineStyle(drawing, stroke);
    }

    private static void LoadStrokeOpacity<T>(IXLDrawing<T> drawing, XElement stroke)
    {
        var opacity = stroke.Attribute("opacity");
        if (opacity != null)
        {
            var opacityVal = opacity.Value;
            if (opacityVal.EndsWith('f'))
                drawing.Style.ColorsAndLines.LineTransparency =
                    double.Parse(opacityVal.AsSpan(0, opacityVal.Length - 1), provider: CultureInfo.InvariantCulture) /
                    65536.0;
            else
                drawing.Style.ColorsAndLines.LineTransparency =
                    double.Parse(opacityVal, CultureInfo.InvariantCulture);
        }
    }

    private static void LoadStrokeDashStyle<T>(IXLDrawing<T> drawing, XElement stroke)
    {
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
    }

    private static void LoadStrokeLineStyle<T>(IXLDrawing<T> drawing, XElement stroke)
    {
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

    private static void LoadLayoutFlowAlt<T>(IXLDrawing<T> xlDrawing, string attrValue)
    {
        if (attrValue.Equals("bottom-to-top"))
            xlDrawing.Style.Alignment.SetOrientation(XLDrawingTextOrientation.BottomToTop);
        else if (attrValue.Equals("top-to-bottom"))
            xlDrawing.Style.Alignment.SetOrientation(XLDrawingTextOrientation.Vertical);
    }
}
