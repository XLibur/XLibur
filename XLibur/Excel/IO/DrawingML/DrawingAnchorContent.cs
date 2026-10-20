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
    /// The <c>mc:AlternateContent</c> that <paramref name="picture"/> is a branch of, or <c>null</c>
    /// when the picture is not directly in an <c>mc:Choice</c> or <c>mc:Fallback</c>.
    /// </summary>
    internal static AlternateContent? WrapperOf(Xdr.Picture picture) =>
        picture.Parent is AlternateContentChoice or AlternateContentFallback
            ? picture.Parent.Parent as AlternateContent
            : null;
}
