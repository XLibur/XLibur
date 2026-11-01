using System.Linq;
using DocumentFormat.OpenXml;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// Readies a drawing's XML, kept from a file, to be written into another drawing part: a copy of a
/// picture or shape, or one written into the package a template save makes.
/// </summary>
internal static class DrawingTemplate
{
    /// <summary>
    /// Removes from <paramref name="root"/> every element that names a relationship, other than
    /// <paramref name="keep"/>: a hyperlink, an SVG form of an image, a picture fill. Each names a
    /// relationship of the drawing part the XML came from, which the part it goes into does not
    /// have, so it is dropped rather than written dangling. An extension holding the reference goes
    /// whole, since an empty <c>a:ext</c> says nothing.
    /// </summary>
    internal static void DropRelationshipReferences(OpenXmlElement root, OpenXmlElement? keep = null)
    {
        foreach (var element in root.Descendants().ToList())
        {
            if (element == keep || element.Parent is null ||
                !element.GetAttributes().Any(a => a.NamespaceUri == RelationshipsNs))
                continue;

            var drawing = element.Parent.LocalName == "ext" ? element.Parent : element;
            drawing.Remove();
        }
    }
}
