using System.Collections.Generic;

namespace XLibur.Excel.Drawings;

/// <summary>
/// The drawing shapes of a worksheet: preset geometries such as rectangles, callouts and arrows, and
/// text boxes (DrawingML <c>xdr:sp</c>).
/// </summary>
/// <remarks>
/// Names follow <see cref="IXLPictures"/>: lookups ignore case, and names need not be unique, because
/// files Excel writes can give two shapes one name. <see cref="Shape"/> returns the first match in
/// drawing order; to tell such shapes apart, enumerate and compare <see cref="IXLShape.Id"/>.
/// </remarks>
public interface IXLShapes : IEnumerable<IXLShape>
{
    /// <summary>The number of shapes on the worksheet.</summary>
    int Count { get; }

    /// <summary>Adds a text box holding <paramref name="text"/>, with its top-left corner at <paramref name="topLeft"/>.</summary>
    /// <param name="text">The text, one paragraph per line. A line break is <c>\n</c>.</param>
    /// <param name="topLeft">The cell the text box hangs off. It moves with the cell, at its own size.</param>
    IXLTextBox AddTextBox(string text, IXLCell topLeft);

    /// <summary>Adds a free-floating text box holding <paramref name="text"/>.</summary>
    /// <param name="text">The text, one paragraph per line. A line break is <c>\n</c>.</param>
    /// <param name="left">The distance from the left edge of the sheet, in pixels.</param>
    /// <param name="top">The distance from the top edge of the sheet, in pixels.</param>
    IXLTextBox AddTextBox(string text, int left, int top);

    /// <summary>Adds a shape of a preset type, with its top-left corner at <paramref name="topLeft"/>.</summary>
    /// <remarks><see cref="XLShapeType.TextBox"/> adds a text box; use <see cref="AddPreset"/> for a geometry the enum does not name.</remarks>
    /// <exception cref="System.ArgumentException"><paramref name="type"/> is <see cref="XLShapeType.Custom"/>.</exception>
    IXLShape Add(XLShapeType type, IXLCell topLeft);

    /// <summary>Adds a free-floating shape of a preset type.</summary>
    /// <param name="type">The preset type.</param>
    /// <param name="left">The distance from the left edge of the sheet, in pixels.</param>
    /// <param name="top">The distance from the top edge of the sheet, in pixels.</param>
    /// <exception cref="System.ArgumentException"><paramref name="type"/> is <see cref="XLShapeType.Custom"/>.</exception>
    IXLShape Add(XLShapeType type, int left, int top);

    /// <summary>
    /// Adds a shape of any DrawingML preset geometry, by its <c>a:prstGeom/@prst</c> name, such as
    /// <c>"moon"</c> or <c>"heart"</c>. The shape reads back as <see cref="XLShapeType.Custom"/>
    /// unless the enum names the geometry.
    /// </summary>
    /// <exception cref="System.ArgumentException"><paramref name="presetGeometry"/> is not a DrawingML preset geometry.</exception>
    IXLShape AddPreset(string presetGeometry, IXLCell topLeft);

    /// <summary>The first shape named <paramref name="name"/>, in any case.</summary>
    /// <exception cref="System.ArgumentOutOfRangeException">No shape has the name.</exception>
    IXLShape Shape(string name);

    /// <summary>Finds the first shape named <paramref name="name"/>, in any case.</summary>
    bool TryGetShape(string name, out IXLShape? shape);

    /// <summary>Whether a shape is named <paramref name="name"/>, in any case.</summary>
    bool Contains(string name);

    /// <summary>Deletes every shape named <paramref name="name"/>, in any case.</summary>
    /// <exception cref="System.ArgumentOutOfRangeException">No shape has the name.</exception>
    void Delete(string name);

    /// <summary>Deletes <paramref name="shape"/>, and no other shape that shares its name.</summary>
    void Delete(IXLShape shape);
}
