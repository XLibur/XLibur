namespace XLibur.Excel;

public enum XLSheetViewOptions { Normal, PageBreakPreview, PageLayout }

public interface IXLSheetView
{
    /// <summary>
    /// Gets or sets whether the split is frozen (<c>true</c>) or a draggable split bar
    /// (<c>false</c>, the default).
    /// </summary>
    /// <remarks>
    /// The <c>Freeze</c> methods set this; assigning <see cref="SplitRow"/> or
    /// <see cref="SplitColumn"/> on their own does not, so a split asked for that way is written as
    /// <c>&lt;pane state="split"&gt;</c> rather than as a freeze. A pane loaded as
    /// <c>frozenSplit</c> — one frozen out of an existing manual split — reads as frozen and saves
    /// as <c>frozen</c>; XLibur has never produced that state and does not start here.
    /// </remarks>
    bool FreezePanes { get; set; }

    /// <summary>
    /// Gets or sets the horizontal split position.
    /// </summary>
    /// <remarks>
    /// When <see cref="FreezePanes"/> is <c>true</c> this is the column after which the split takes
    /// place — a count of frozen columns. When it is <c>false</c> this is <c>ST_Pane</c>'s
    /// <c>xSplit</c>, the split bar's position in twentieths of a point, which XLibur carries
    /// verbatim rather than reinterpreting.
    /// </remarks>
    int SplitColumn { get; set; }

    /// <summary>
    /// Gets or sets the vertical split position.
    /// </summary>
    /// <remarks>
    /// When <see cref="FreezePanes"/> is <c>true</c> this is the row after which the split takes
    /// place — a count of frozen rows. When it is <c>false</c> this is <c>ST_Pane</c>'s
    /// <c>ySplit</c>, the split bar's position in twentieths of a point, which XLibur carries
    /// verbatim rather than reinterpreting.
    /// </remarks>
    int SplitRow { get; set; }

    /// <summary>
    /// Gets or sets the location of the top left visible cell
    /// </summary>
    /// <value>
    /// The scroll position's top left cell.
    /// </value>
    IXLAddress TopLeftCellAddress { get; set; }

    /// <summary>
    /// Gets or sets the top-left visible cell of the scrollable region below/right of a
    /// split (maps to <c>&lt;pane topLeftCell&gt;</c>).
    /// </summary>
    /// <remarks>
    /// When <c>null</c> (the default), XLibur anchors the pane to the first non-frozen cell
    /// (<c>split + 1</c>) for a freeze, and to <c>A1</c> for an unfrozen split, whose splits are
    /// not line counts — so a worksheet that has never set this value normalizes its scrollable
    /// region to the top on save. When set, the value is honored on save. Has no effect when the
    /// sheet has no split at all (no <c>&lt;pane&gt;</c> element is emitted).
    /// </remarks>
    /// <value>
    /// The scrollable region's top left cell, or <c>null</c> to use the default anchor.
    /// </value>
    IXLAddress? PaneTopLeftCellAddress { get; set; }

    XLSheetViewOptions View { get; set; }

    IXLWorksheet Worksheet { get; }

    /// <summary>
    /// Window zoom magnification for current view representing percent values. Horizontal and vertical scale together.
    /// </summary>
    /// <remarks>Representing percent values ranging from 10 to 400.</remarks>
    int ZoomScale { get; set; }

    /// <summary>
    /// Zoom magnification to use when in normal view. Horizontal and vertical scale together
    /// </summary>
    /// <remarks>Representing percent values ranging from 10 to 400.</remarks>
    int ZoomScaleNormal { get; set; }

    /// <summary>
    /// Zoom magnification to use when in page layout view. Horizontal and vertical scale together.
    /// </summary>
    /// <remarks>Representing percent values ranging from 10 to 400.</remarks>
    int ZoomScalePageLayoutView { get; set; }

    /// <summary>
    /// Zoom magnification to use when in page break preview. Horizontal and vertical scale together.
    /// </summary>
    /// <remarks>Representing percent values ranging from 10 to 400.</remarks>
    int ZoomScaleSheetLayoutView { get; set; }

    /// <summary>
    /// Freezes the specified rows and columns.
    /// </summary>
    /// <param name="rows">The rows to freeze.</param>
    /// <param name="columns">The columns to freeze.</param>
    void Freeze(int rows, int columns);

    /// <summary>
    /// Freezes the left X columns.
    /// </summary>
    /// <param name="columns">The columns to freeze.</param>
    void FreezeColumns(int columns);

    /// <summary>
    /// Freezes the top X rows.
    /// </summary>
    /// <param name="rows">The rows to freeze.</param>
    void FreezeRows(int rows);

    IXLSheetView SetView(XLSheetViewOptions value);
}
