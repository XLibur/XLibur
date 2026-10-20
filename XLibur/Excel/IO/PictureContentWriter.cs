using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel.Drawings;
using XLibur.Excel.IO.DrawingML;
using XLibur.Extensions;
using static XLibur.Excel.IO.OpenXmlConst;
using static XLibur.Excel.XLWorkbook;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// What a written picture shows: the name and image link the model owns, the image parts behind
/// those links, and the relationships a removed picture leaves behind.
/// </summary>
internal static class PictureContentWriter
{
    /// <summary>
    /// Writes the name and image link the model owns into a picture written back into its anchor.
    /// </summary>
    /// <remarks>
    /// The name goes to every drawing in the anchor with the picture's id. Excel gives each branch
    /// of <c>mc:AlternateContent</c> the same id and name, as one drawing in two forms, so a 3D
    /// model renamed through its fallback picture is renamed in the form Excel shows. The image
    /// link goes to every image in the wrapper that showed the picture's image: a 3D model's
    /// <c>am3d:raster</c> is the render its fallback picture shows, so the two stay on one part.
    /// The shape properties of the other branches are left as Excel wrote them.
    /// </remarks>
    internal static void WritePictureNameAndImage(OpenXmlElement anchor, Xdr.Picture picture, string name,
        string embedId)
    {
        if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value is { } id)
        {
            foreach (var properties in anchor.Descendants<Xdr.NonVisualDrawingProperties>())
            {
                if (properties.Id?.Value == id)
                    properties.Name = name;
            }
        }

        if (picture.BlipFill?.Blip is not { } blip)
            return;

        var previous = blip.Embed?.Value;
        blip.Embed = embedId;
        if (previous is null || previous == embedId || DrawingAnchorContent.WrapperOf(picture) is not { } wrapper)
            return;

        foreach (var image in wrapper.Descendants())
        {
            if (image.LocalName == "blip" && EmbedOf(image) == previous)
                image.SetAttribute(new OpenXmlAttribute("r", "embed", RelationshipsNs, embedId));
        }
    }

    /// <summary>The <c>r:embed</c> of an image element, whatever namespace the element is in.</summary>
    private static string? EmbedOf(OpenXmlElement image) =>
        image.GetAttributes().FirstOrDefault(a => a.LocalName == "embed" && a.NamespaceUri == RelationshipsNs).Value;

    /// <summary>
    /// The image part <paramref name="pic"/> is written with, holding its image. Each part is fed
    /// once per save, by the first picture that shows it; <paramref name="fedImages"/> holds, by
    /// relationship id, the image each part was fed this save. The picture's own link is left as it
    /// was: it takes the link the save wrote once the save has succeeded, as it takes its id.
    /// </summary>
    /// <remarks>
    /// Pictures that share a part, as Excel's copies of a picture do (#781), feed it once. A copy
    /// whose image differs from what its part was fed, because the image was changed after the load,
    /// shows a part that holds its own image instead: one this save already fed with the same image,
    /// or a new one. So a changed image is not overwritten by the copies that kept theirs, and does
    /// not reach them.
    /// </remarks>
    internal static ImagePart FeedImagePart(DrawingsPart drawingsPart, XLPicture pic, SaveContext context,
        Dictionary<string, MemoryStream> fedImages)
    {
        if (pic.RelId is { } relId && drawingsPart.GetPartOrNull<ImagePart>(relId) is { } loaded)
        {
            if (!fedImages.TryGetValue(relId, out var fed))
                return Feed(loaded, relId);

            if (SameImage(fed, pic.ImageStream))
                return loaded;

            foreach (var (otherRelId, otherImage) in fedImages)
            {
                if (SameImage(otherImage, pic.ImageStream))
                    return drawingsPart.GetPartOrNull<ImagePart>(otherRelId)!;
            }
        }

        var newRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
        return Feed(drawingsPart.AddImagePart(pic.Format.ToOpenXml(), newRelId), newRelId);

        ImagePart Feed(ImagePart part, string partRelId)
        {
            pic.ImageStream.Position = 0;
            part.FeedData(pic.ImageStream);
            fedImages[partRelId] = pic.ImageStream;
            return part;
        }
    }

    private static bool SameImage(MemoryStream left, MemoryStream right) =>
        ReferenceEquals(left, right) || Bytes(left).SequenceEqual(Bytes(right));

    private static ReadOnlySpan<byte> Bytes(MemoryStream stream) =>
        stream.TryGetBuffer(out var buffer) ? buffer.AsSpan() : stream.ToArray();

    /// <summary>
    /// Every relationship id <paramref name="element"/> refers to, by any attribute in the
    /// relationships namespace: the image of a picture's <c>a:blip</c>, and whatever Excel added
    /// that XLibur does not model, such as a 3D model's part and its render's <c>am3d:blip</c>.
    /// </summary>
    internal static IEnumerable<string> RelationshipIds(OpenXmlElement element) =>
        element.Descendants().Prepend(element)
            .SelectMany(e => e.GetAttributes())
            .Where(a => a.NamespaceUri == RelationshipsNs && !string.IsNullOrEmpty(a.Value))
            .Select(a => a.Value!);

    /// <summary>
    /// Drops each of the drawing part's relationships in <paramref name="released"/> that nothing left
    /// in <paramref name="worksheetDrawing"/> refers to, once the drawings that referred to it have
    /// been removed.
    /// </summary>
    /// <remarks>
    /// Others may still refer to it: Excel shares one image part, and one 3D model part, between
    /// copies. Only this drawing's relationship is dropped; the SDK deletes the part it targets only
    /// when no other part, such as another sheet's drawing, still reaches it.
    /// </remarks>
    internal static void DeleteUnreferencedRelationships(DrawingsPart drawingsPart,
        Xdr.WorksheetDrawing worksheetDrawing, List<string> released)
    {
        if (released.Count == 0)
            return;

        var stillReferenced = RelationshipIds(worksheetDrawing).ToHashSet(StringComparer.Ordinal);
        foreach (var relId in released.Distinct(StringComparer.Ordinal))
        {
            if (stillReferenced.Contains(relId))
                continue;

            if (drawingsPart.HasPartWithId(relId))
                drawingsPart.DeletePart(relId);
            else if (HasReferenceRelationship(drawingsPart, relId))
                drawingsPart.DeleteReferenceRelationship(relId);
        }
    }

    /// <summary>Whether the part has a relationship <paramref name="relId"/> to no part: a hyperlink, an external target or media.</summary>
    private static bool HasReferenceRelationship(DrawingsPart drawingsPart, string relId) =>
        drawingsPart.ExternalRelationships.Any(r => r.Id == relId) ||
        drawingsPart.HyperlinkRelationships.Any(r => r.Id == relId) ||
        drawingsPart.DataPartReferenceRelationships.Any(r => r.Id == relId);
}
