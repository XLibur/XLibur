using System;
using System.Collections.Generic;

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
/// Excel showed, and the model writes the anchor from its pixel geometry instead.
/// </para>
/// </remarks>
internal sealed class XLKeptAnchor
{
    private readonly XLWorksheet _worksheet;
    private readonly GridLayout _loadedLayout;

    internal XLKeptAnchor(XLWorksheet worksheet, XLPicturePlacement placement, XLMarker from, XLMarker to)
    {
        _worksheet = worksheet;
        Placement = placement;
        From = from;
        To = to;
        _loadedLayout = MeasureLayout();
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
    /// read from where the markers are now.
    /// </summary>
    /// <remarks>
    /// An insert that moves the markers moves the lines with them, so rows inserted above a one-cell
    /// picture leave this true. Rows of the default height inserted inside its span, or above an
    /// absolute picture, leave it true too, because the lines the markers now count across have the
    /// same sizes as before. A row that is resized, hidden or shown, or inserted at another height,
    /// makes it false.
    /// </remarks>
    internal bool IsLayoutUnchanged => _loadedLayout.Matches(MeasureLayout());

    /// <summary>
    /// A one-cell picture's size is the lines from its top-left cell to its bottom-right cell; an
    /// absolute picture's position also takes every line before it, from the sheet's corner.
    /// </summary>
    private GridLayout MeasureLayout()
    {
        var countsFromCorner = Placement == XLPicturePlacement.FreeFloating;
        var firstRow = countsFromCorner ? 1 : From.RowNumber;
        var firstColumn = countsFromCorner ? 1 : From.ColumnNumber;

        return new GridLayout(
            LineLayout.Measure(_worksheet.Internals.RowsCollection, firstRow, To.RowNumber, _worksheet.RowHeight,
                static row => row.IsHidden ? 0 : row.Height),
            LineLayout.Measure(_worksheet.Internals.ColumnsCollection, firstColumn, To.ColumnNumber, _worksheet.ColumnWidth,
                static column => column.IsHidden ? 0 : column.Width));
    }

    private sealed record GridLayout(LineLayout Rows, LineLayout Columns)
    {
        public bool Matches(GridLayout other) => Rows.Matches(other.Rows) && Columns.Matches(other.Columns);
    }

    /// <summary>
    /// The sizes of a run of rows or columns: the sheet's default, and each line that differs from it,
    /// by its position in the run. Lines the sheet has never materialised take the default, so a run
    /// is described by the few lines that differ rather than by every line in it.
    /// </summary>
    private sealed record LineLayout(double DefaultSize, (int Offset, double Size)[] Exceptions)
    {
        public bool Matches(LineLayout other) =>
            DefaultSize.Equals(other.DefaultSize) && Exceptions.AsSpan().SequenceEqual(other.Exceptions);

        public static LineLayout Measure<TLine>(IDictionary<int, TLine> lines, int first, int last,
            double defaultSize, Func<TLine, double> sizeOf)
        {
            var exceptions = new List<(int Offset, double Size)>();
            foreach (var (number, line) in lines)
            {
                if (number < first || number > last)
                    continue;

                var size = sizeOf(line);
                if (!size.Equals(defaultSize))
                    exceptions.Add((number - first, size));
            }

            exceptions.Sort();
            return new LineLayout(defaultSize, [.. exceptions]);
        }
    }
}
