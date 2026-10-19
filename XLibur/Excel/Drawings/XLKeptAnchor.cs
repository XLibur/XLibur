namespace XLibur.Excel.Drawings;

/// <summary>
/// The two-cell anchor a loaded picture came from, when the model cannot say it back: a
/// <c>xdr:twoCellAnchor</c> whose <c>editAs</c> is <c>oneCell</c> or <c>absolute</c> (issue #780).
/// </summary>
/// <remarks>
/// <para>
/// Excel places and sizes every two-cell anchor from its <c>from</c> and <c>to</c> markers, whatever
/// <c>editAs</c> says; <c>editAs</c> only decides what a later row or column edit does to the
/// picture. It is also Excel's default for an inserted picture. The model maps those two forms to
/// <see cref="XLPicturePlacement.Move"/> and <see cref="XLPicturePlacement.FreeFloating"/>, which
/// carry a pixel size and, for a free-floating picture, a pixel position. Both are read from the
/// picture's <c>a:xfrm</c>, which Excel does not use for a two-cell anchor and which is wrong
/// whenever a writer other than Excel produced it: XLibur and ClosedXML write offset 0,0 and the
/// picture's native size. Regenerating the anchor from the model therefore shrank a picture to its
/// native size, or moved it to the top-left corner of the sheet.
/// </para>
/// <para>
/// So the loaded markers are kept, and an unedited picture is written back by editing its loaded
/// anchor in place. Any edit of the picture's placement, size or position drops this, and the model
/// decides the anchor again. Renaming the picture or replacing its image does not, because neither
/// is in the anchor.
/// </para>
/// <para>
/// The markers still follow row and column edits, through <see cref="DrawingAnchorListener"/>, the
/// way Excel moves the picture: a <see cref="XLPicturePlacement.Move"/> picture moves both corners by
/// the top-left corner's shift and keeps its span, and a
/// <see cref="XLPicturePlacement.FreeFloating"/> one does not move. Only the column and the row of a
/// marker are ever written back; the offsets inside the cell stay as the file had them, so a
/// pixel round trip cannot drift them.
/// </para>
/// </remarks>
internal sealed class XLKeptAnchor(XLPicturePlacement placement, XLMarker from, XLMarker to)
{
    /// <summary>
    /// <see cref="XLPicturePlacement.Move"/> for <c>editAs="oneCell"</c>,
    /// <see cref="XLPicturePlacement.FreeFloating"/> for <c>editAs="absolute"</c>.
    /// </summary>
    internal XLPicturePlacement Placement { get; } = placement;

    /// <summary>The top-left corner, as <c>xdr:from</c> had it.</summary>
    internal XLMarker From { get; } = from;

    /// <summary>The bottom-right corner, as <c>xdr:to</c> had it.</summary>
    internal XLMarker To { get; } = to;
}
