namespace XLibur.Excel.Drawings;

/// <summary>
/// A drawing shape on a worksheet: a preset geometry that can carry text (DrawingML <c>xdr:sp</c>).
/// </summary>
/// <remarks>
/// <para>
/// A formatting property that is <c>null</c> was not assigned: the save leaves whatever the file has
/// for it, which may be a gradient, a picture fill or a colour from the shape's theme style. Only the
/// properties assigned since the shape was loaded or last saved are written. <see cref="NoFill"/> and
/// <see cref="NoLine"/> are the way to remove a fill or a line.
/// </para>
/// <para>
/// A shape loaded from a file is written back by editing its own XML, so what XLibur does not model —
/// rotation, effects, 3-D, hyperlinks, adjust handles — survives a save.
/// </para>
/// </remarks>
public interface IXLShape
{
    /// <summary>The drawing id, <c>xdr:cNvPr/@id</c>. <c>0</c> for a shape no save has written yet.</summary>
    int Id { get; }

    /// <summary>The name, <c>xdr:cNvPr/@name</c>.</summary>
    string Name { get; set; }

    /// <summary>The preset type, or <see cref="XLShapeType.Custom"/> for a geometry the enum does not name.</summary>
    XLShapeType ShapeType { get; }

    /// <summary>The DrawingML preset geometry, <c>a:prstGeom/@prst</c>, such as <c>"rect"</c>.</summary>
    string PresetGeometry { get; }

    IXLWorksheet Worksheet { get; }

    /// <summary>The distance from the left of the top-left cell, or of the sheet when free-floating, in pixels.</summary>
    /// <exception cref="System.ArgumentException">On set, when <see cref="Placement"/> is not <see cref="XLPicturePlacement.FreeFloating"/>.</exception>
    int Left { get; set; }

    /// <summary>The distance from the top of the top-left cell, or of the sheet when free-floating, in pixels.</summary>
    /// <exception cref="System.ArgumentException">On set, when <see cref="Placement"/> is not <see cref="XLPicturePlacement.FreeFloating"/>.</exception>
    int Top { get; set; }

    /// <summary>The width in pixels.</summary>
    /// <exception cref="System.ArgumentException">On set, when <see cref="Placement"/> is <see cref="XLPicturePlacement.MoveAndSize"/>, which sizes the shape by its cells.</exception>
    int Width { get; set; }

    /// <summary>The height in pixels.</summary>
    /// <exception cref="System.ArgumentException">On set, when <see cref="Placement"/> is <see cref="XLPicturePlacement.MoveAndSize"/>, which sizes the shape by its cells.</exception>
    int Height { get; set; }

    /// <summary>
    /// How the shape follows the cells under it. The type is named for pictures, the first drawings
    /// XLibur wrote, but it says the same three things about any drawing. A loaded shape given a new
    /// placement moves into an anchor of the new form with everything it holds.
    /// </summary>
    XLPicturePlacement Placement { get; set; }

    /// <summary>The cell the top-left corner hangs off, or A1 for a free-floating shape.</summary>
    IXLCell TopLeftCell { get; }

    /// <summary>The cell the bottom-right corner hangs off, for <see cref="XLPicturePlacement.MoveAndSize"/>.</summary>
    /// <exception cref="System.InvalidOperationException">The placement has no bottom-right cell.</exception>
    IXLCell BottomRightCell { get; }

    /// <summary>The fill colour, or <c>null</c> when none was read or assigned. Setting a colour clears <see cref="NoFill"/>.</summary>
    XLColor? FillColor { get; set; }

    /// <summary>Whether the shape has no fill, <c>a:noFill</c>. Setting it clears <see cref="FillColor"/>.</summary>
    bool NoFill { get; set; }

    /// <summary>The line colour, or <c>null</c> when none was read or assigned. Setting a colour clears <see cref="NoLine"/>.</summary>
    XLColor? LineColor { get; set; }

    /// <summary>The line width in points, or <c>null</c> when none was read or assigned.</summary>
    double? LineWidthPt { get; set; }

    /// <summary>The line dash, or <c>null</c> when none was read or assigned.</summary>
    XLLineDashStyle? LineDash { get; set; }

    /// <summary>Whether the shape has no line, <c>a:ln/a:noFill</c>. Setting it clears <see cref="LineColor"/>.</summary>
    bool NoLine { get; set; }

    /// <summary>The text the shape carries, by paragraph.</summary>
    IXLTextBody TextBody { get; }

    IXLShape MoveTo(int left, int top);

    IXLShape MoveTo(IXLCell cell);

    IXLShape MoveTo(IXLCell cell, int xOffset, int yOffset);

    IXLShape MoveTo(IXLCell fromCell, IXLCell toCell);

    IXLShape WithSize(int width, int height);

    IXLShape WithPlacement(XLPicturePlacement value);

    /// <summary>Removes the shape from the worksheet.</summary>
    void Delete();
}

/// <summary>A shape whose <c>cNvSpPr/@txBox</c> is 1: Excel's text box, a rectangle by default.</summary>
public interface IXLTextBox : IXLShape
{
    /// <summary>
    /// The whole text, paragraphs joined by <c>\n</c>. Setting it replaces every paragraph; each new
    /// paragraph takes the properties of the first paragraph the body had, and each new run the
    /// formatting of its first run. A control character other than a tab or <c>\n</c> throws
    /// <see cref="System.ArgumentException"/>.
    /// </summary>
    string Text { get; set; }
}
