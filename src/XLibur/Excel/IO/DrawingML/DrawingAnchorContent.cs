using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// Where a picture sits inside its anchor, as the loader reads it and the save writes it back.
/// </summary>
/// <remarks>
/// Excel puts a drawing that also has a newer form inside <c>mc:AlternateContent</c>: a 3D model is
/// an <c>xdr:graphicFrame</c> in the <c>mc:Choice</c> and an <c>xdr:pic</c> of its render in the
/// <c>mc:Fallback</c>, both directly in the anchor's wrapper (issue #787).
/// </remarks>
internal static class DrawingAnchorContent
{
    /// <summary>
    /// Whether <paramref name="picture"/> is the drawing <paramref name="anchor"/> places: its child,
    /// or a branch of the <c>mc:AlternateContent</c> that is its child. The loader reads such a
    /// picture as the anchor's own, sized and placed by the anchor, and the save writes it back.
    /// </summary>
    /// <remarks>
    /// A picture any deeper, as in a group inside the wrapper, is not: the anchor places the group,
    /// and the picture is sized in the group's own coordinates. XLibur leaves it as the file has it,
    /// as it does a 3D model inside a group, which is how Excel 16.0 writes a grouped model:
    /// <c>xdr:grpSp</c> holding the <c>mc:AlternateContent</c>.
    /// </remarks>
    internal static bool IsAnchorPicture(Xdr.Picture picture, OpenXmlElement anchor) =>
        picture.Parent == anchor || WrapperOf(picture)?.Parent == anchor;

    /// <summary>
    /// The <c>mc:AlternateContent</c> that <paramref name="picture"/> is a branch of, or <c>null</c>
    /// when the picture is not directly in an <c>mc:Choice</c> or <c>mc:Fallback</c>.
    /// </summary>
    internal static AlternateContent? WrapperOf(Xdr.Picture picture) =>
        picture.Parent is AlternateContentChoice or AlternateContentFallback
            ? picture.Parent.Parent as AlternateContent
            : null;

    /// <summary>
    /// Every form of the drawing <paramref name="picture"/> shows: the drawing in each branch of its
    /// <c>mc:AlternateContent</c>, such as a 3D model's frame and its render, or the picture alone
    /// when it is not a branch.
    /// </summary>
    internal static IEnumerable<OpenXmlElement> FormsOf(Xdr.Picture picture) =>
        WrapperOf(picture) is { } wrapper
            ? wrapper.ChildElements.SelectMany(branch => branch.ChildElements)
            : [picture];
}
