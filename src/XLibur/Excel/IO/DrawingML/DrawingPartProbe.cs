using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace XLibur.Excel.IO.DrawingML;

/// <summary>
/// Answers questions about what a drawing part contains without attaching its DOM.
/// </summary>
/// <remarks>
/// <para>
/// <b>Reading <see cref="DrawingsPart.WorksheetDrawing"/> is not free and not read-only.</b> The
/// property materialises the SDK's typed tree and attaches it to the part, and the SDK then writes
/// that tree back over the part's original bytes when the package is saved — whether or not
/// anything was changed. The part then round-trips as the SDK's serialisation rather than the
/// producer's: self-closing tags gain a space before <c>/&gt;</c>, the XML declaration's
/// <c>encoding</c> is lower-cased, and namespace declarations are hoisted to the root.
/// </para>
/// <para>
/// That matters for a sheet whose drawing holds nothing XLibur models — a slicer or a timeline
/// frame and nothing else. Such a part should come through untouched under the preservation rule in
/// <c>docs/round-trip-fidelity.md</c>, and it did not, because the save path reached for
/// <c>WorksheetDrawing</c> purely to ask whether the drawing was empty. Asking the question changed
/// the answer.
/// </para>
/// <para>
/// This streams the part through an <see cref="OpenXmlPartReader"/> instead, which leaves
/// <c>part.RootElement</c> unmaterialised — the same technique <c>SlicerReader</c> uses, for the
/// same reason.
/// </para>
/// </remarks>
internal static class DrawingPartProbe
{
    /// <summary>
    /// Whether the drawing has any element at all beneath its root.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The replacement for <c>WorksheetDrawing?.HasChildren</c> where the only question is whether
    /// the part would be saved empty.
    /// </para>
    /// <para>
    /// <b>A part whose DOM is already attached is answered from the DOM, not from the stream.</b>
    /// That is not a fallback, it is a correctness requirement: earlier in the same save,
    /// <c>PictureWriter</c> may have deleted the drawing's last picture out of the attached tree, and
    /// the bytes still on disk would then say the drawing has children when it no longer does —
    /// leaving an empty drawing part in the package. There is also nothing left to protect by
    /// streaming at that point, since the SDK will write the attached tree back either way.
    /// </para>
    /// </remarks>
    internal static bool HasAnyChild(DrawingsPart drawingsPart)
    {
        if (drawingsPart.IsRootElementLoaded)
            return drawingsPart.WorksheetDrawing?.HasChildren ?? false;

        using var reader = new OpenXmlPartReader(drawingsPart);

        // Creating the reader consumes the XML declaration only, so the first Read lands on the root
        // element and anything after it is a child.
        if (!reader.Read())
            return false;

        while (reader.Read())
        {
            if (reader.IsStartElement)
                return true;
        }

        return false;
    }
}
