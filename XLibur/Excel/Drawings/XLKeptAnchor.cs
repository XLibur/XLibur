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
/// The markers still follow row and column inserts and deletes, through
/// <see cref="DrawingAnchorListener"/>, the way Excel moves the picture: a
/// <see cref="XLPicturePlacement.Move"/> picture moves both corners by the top-left corner's shift
/// and keeps its span, and a <see cref="XLPicturePlacement.FreeFloating"/> one does not move. Only
/// the column and the row of a marker are ever written back; the offsets inside the cell stay as the
/// file had them, so a pixel round trip cannot drift them.
/// </para>
/// <para>
/// <b>Markers name cells, not lengths.</b> Kept markers keep the picture's size and position only
/// while the rows and columns they count across keep theirs: a taller row inside a one-cell
/// picture's span would stretch it, and a taller row above an absolute picture would move it, where
/// Excel keeps both fixed. The layout of those lines is recorded at load and compared at save
/// (<see cref="IsLayoutUnchanged"/>); when it differs, the kept markers are no longer the picture
/// Excel showed, and the model writes the anchor from its pixel geometry instead. The layout is
/// taken once for the sheet each time (<see cref="GridLayout"/>), not once for each picture, so a
/// sheet with many pictures and many sized rows does not pay for their product (issue #785).
/// </para>
/// </remarks>
internal sealed class XLKeptAnchor
{
    private readonly GridLayout _loadedLayout;
    private readonly (int First, int Last) _loadedRows;
    private readonly (int First, int Last) _loadedColumns;

    /// <summary>
    /// Keeps the markers loaded for a picture, and where they lie in <paramref name="loadedLayout"/>,
    /// the sheet's layout at load, which every kept anchor on the sheet shares.
    /// </summary>
    internal XLKeptAnchor(GridLayout loadedLayout, XLPicturePlacement placement, XLMarker from, XLMarker to)
    {
        _loadedLayout = loadedLayout;
        Placement = placement;
        From = from;
        To = to;
        _loadedRows = RowRun();
        _loadedColumns = ColumnRun();
    }

    /// <summary>
    /// <see cref="XLPicturePlacement.Move"/> for <c>editAs="oneCell"</c>,
    /// <see cref="XLPicturePlacement.FreeFloating"/> for <c>editAs="absolute"</c>.
    /// </summary>
    internal XLPicturePlacement Placement { get; }

    /// <summary>The top-left corner, as <c>xdr:from</c> had it.</summary>
    internal XLMarker From { get; }

    /// <summary>The bottom-right corner, as <c>xdr:to</c> had it.</summary>
    internal XLMarker To { get; }

    /// <summary>
    /// Whether the rows and columns the markers count across still have the sizes they had at load,
    /// read from where the markers are now in <paramref name="current"/>, the sheet's layout at save.
    /// </summary>
    /// <remarks>
    /// An insert that moves the markers moves the lines with them, so rows inserted above a one-cell
    /// picture leave this true. Rows of the default height inserted inside its span, or above an
    /// absolute picture, leave it true too, because the lines the markers now count across have the
    /// same sizes as before. A row that is resized, hidden or shown, or inserted at another height,
    /// makes it false.
    /// </remarks>
    internal bool IsLayoutUnchanged(GridLayout current)
    {
        var (firstRow, lastRow) = RowRun();
        var (firstColumn, lastColumn) = ColumnRun();
        return current.Rows.HasSameSizes(firstRow, lastRow, _loadedLayout.Rows, _loadedRows.First, _loadedRows.Last) &&
               current.Columns.HasSameSizes(firstColumn, lastColumn, _loadedLayout.Columns,
                   _loadedColumns.First, _loadedColumns.Last);
    }

    /// <summary>
    /// A one-cell picture's size is the lines from its top-left cell to its bottom-right cell; an
    /// absolute picture's position also takes every line before it, from the sheet's corner.
    /// </summary>
    private (int First, int Last) RowRun() => (CountsFromCorner ? 1 : From.RowNumber, To.RowNumber);

    private (int First, int Last) ColumnRun() => (CountsFromCorner ? 1 : From.ColumnNumber, To.ColumnNumber);

    private bool CountsFromCorner => Placement == XLPicturePlacement.FreeFloating;
}
