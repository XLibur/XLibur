using System.Linq;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes the fill and line assigned to a shape into its <c>xdr:spPr</c>, and nothing else: a
/// property nobody assigned is left as the file or the creation defaults have it (spec 15).
/// </summary>
/// <remarks>
/// This decides whether to write each property; how is <see cref="ShapePropertiesWriter"/>'s, the
/// layer pictures and charts write through. The same writes serve a new shape and a loaded one.
/// </remarks>
internal static class ShapeFormatWriter
{
    internal static void Write(Xdr.ShapeProperties shapeProperties, XLShape shape)
    {
        var assigned = shape.Assigned;
        if (assigned == XLShapeProperties.None)
            return;

        if (Has(assigned, XLShapeProperties.NoFill) && shape.NoFill)
            ShapePropertiesWriter.SetNoFill(shapeProperties);
        else if (Has(assigned, XLShapeProperties.FillColor))
            ShapePropertiesWriter.SetFill(shapeProperties, shape.FillColor);
        else if (Has(assigned, XLShapeProperties.NoFill))
            ShapePropertiesWriter.RemoveNoFill(shapeProperties);

        WriteLine(shapeProperties, shape, assigned);
    }

    private static void WriteLine(Xdr.ShapeProperties shapeProperties, XLShape shape, XLShapeProperties assigned)
    {
        const XLShapeProperties line = XLShapeProperties.LineColor | XLShapeProperties.LineWidth |
                                       XLShapeProperties.LineDash | XLShapeProperties.NoLine;
        if (!Has(assigned, line))
            return;

        // An outline is added only for a value to put in it; one that only loses values is edited
        // if it is there.
        var wantsOutline = (Has(assigned, XLShapeProperties.LineColor) && shape.LineColor is not null) ||
                           (Has(assigned, XLShapeProperties.LineWidth) && shape.LineWidthPt is not null) ||
                           (Has(assigned, XLShapeProperties.LineDash) && shape.LineDash is not null) ||
                           (Has(assigned, XLShapeProperties.NoLine) && shape.NoLine);
        var outline = wantsOutline
            ? ShapePropertiesWriter.EnsureOutline(shapeProperties)
            : shapeProperties.Elements<A.Outline>().FirstOrDefault();
        if (outline is null)
            return;

        if (Has(assigned, XLShapeProperties.NoLine) && shape.NoLine)
            ShapePropertiesWriter.SetOutlineNoFill(outline);
        else if (Has(assigned, XLShapeProperties.LineColor))
            ShapePropertiesWriter.SetOutlineColor(outline, shape.LineColor);
        else if (Has(assigned, XLShapeProperties.NoLine))
            ShapePropertiesWriter.RemoveOutlineNoFill(outline);

        if (Has(assigned, XLShapeProperties.LineWidth))
            ShapePropertiesWriter.SetOutlineWidth(outline, shape.LineWidthPt);

        if (Has(assigned, XLShapeProperties.LineDash))
            ShapePropertiesWriter.SetOutlineDash(outline, shape.LineDash);
    }

    private static bool Has(XLShapeProperties assigned, XLShapeProperties any) => (assigned & any) != 0;
}
