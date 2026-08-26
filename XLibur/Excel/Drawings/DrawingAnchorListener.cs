using System.Linq;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.Drawings;

/// <summary>
/// Moves the anchors of every chart, note and picture on the sheet when rows or columns are inserted
/// or deleted.
/// </summary>
/// <remarks>
/// <para>
/// Before spec 33 nothing notified these. <see cref="XLDrawingPosition"/> holds raw <c>int</c>s and
/// neither the chart collection nor the note slice implemented <see cref="ISheetListener"/>, so a
/// chart anchored at row 10 stayed at row 10 when three rows were inserted above it. Notes were half
/// broken rather than wholly: the note moved with the misc slice while its callout box did not,
/// which <c>XLComment</c>'s own remarks already documented — "shifting rows or columns moves the
/// note's entry within the misc slice <em>without telling the note</em>". This is the telling.
/// </para>
/// <para>
/// The transform is <see cref="GridShift"/>, which is <see cref="XLRangeShiftHelper"/>'s reduced to
/// the integers, so these anchors move exactly as a picture's anchor already moved through the range
/// repository. The anchor cases were established by measuring that picture rather than assumed; they
/// are listed on <see cref="GridShift"/> and recorded in this spec's Results.
/// </para>
/// <para>
/// The picture came here too, in spec 33 task 6. It was the one anchor that already worked, and it
/// worked by allocating a one-cell <c>IXLRange</c> so the range repository would shift it — the
/// workaround this seam exists to replace. Its behaviour is pinned by
/// <c>PictureAnchorShiftTests</c>, which is what makes moving it a proof rather than a rewrite.
/// </para>
/// <para>
/// <b>Two coordinate conventions share one type.</b> A chart's <see cref="XLDrawingPosition"/> is
/// written verbatim into <c>xdr:from</c>/<c>xdr:to</c>, whose <c>col</c> and <c>row</c> are
/// <em>0-based</em> per ECMA-376 §20.5.2.7 and §20.5.2.32. A note's is 1-based: the VML writer emits
/// <c>Position.Row - 1</c> and indexes <c>Worksheet.Row(Position.Row)</c> directly. So each kind
/// declares its own base rather than the adapter assuming one. Recorded as D16 in
/// <c>DEFECTS.md</c>.
/// </para>
/// <para>
/// An absolutely anchored drawing is left alone: its position is written in EMU with no cell
/// reference (<c>xdr:absoluteAnchor</c>, ECMA-376 §20.5.2.1), so the grid cannot move it. A chart's
/// second position is moved only under <see cref="XLDrawingAnchor.MoveAndSizeWithCells"/>, which is
/// the only mode that writes it.
/// </para>
/// </remarks>
internal sealed class DrawingAnchorListener(XLWorksheet worksheet) : ISheetListener
{
    void ISheetListener.OnInsertAreaAndShiftDown(in SheetEdit edit) => MoveAnchors<RowAxis>(in edit);

    void ISheetListener.OnInsertAreaAndShiftRight(in SheetEdit edit) => MoveAnchors<ColumnAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftUp(in SheetEdit edit) => MoveAnchors<RowAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftLeft(in SheetEdit edit) => MoveAnchors<ColumnAxis>(in edit);

    private void MoveAnchors<TAxis>(in SheetEdit edit)
        where TAxis : struct, IGridAxis
    {
        if (edit.Sheet != worksheet)
            return;

        foreach (var chart in worksheet.Charts.OfType<XLChart>())
        {
            if (chart.Anchor == XLDrawingAnchor.Absolute)
                continue;

            // 0-based: the chart's markers are written straight into xdr:from / xdr:to.
            Move<TAxis>(chart.Position, edit, oneBased: false);

            if (chart.Anchor == XLDrawingAnchor.MoveAndSizeWithCells)
                Move<TAxis>(chart.SecondPosition, edit, oneBased: false);
        }

        foreach (var cell in worksheet.Internals.CellsCollection.GetCells(c => c.HasComment))
        {
            var note = cell.SliceComment;

            // XLComment.Anchor, not Style.Properties.Positioning. The two disagree on every note
            // XLibur creates: Initialize sets Anchor to MoveAndSizeWithCells while the inherited
            // DefaultCommentStyle sets Positioning to Absolute, and the VML writer reads the latter.
            // Anchor is the field that names how this note is tied to the grid, and it is the one a
            // caller changing that would reach for. Which of the two should drive what is written
            // into VML is a real question and a separate one — recorded as D17 in DEFECTS.md.
            if (note is null || note.Anchor == XLDrawingAnchor.Absolute)
                continue;

            // 1-based: the VML writer subtracts one and indexes rows and columns with it directly.
            Move<TAxis>(note.Position, edit, oneBased: true);
        }

        foreach (var picture in worksheet.Pictures.OfType<XLPicture>())
        {
            // A free-floating picture is placed in pixels from the sheet's corner and has no cell to
            // follow — the same case as an absolutely anchored chart. Its markers are still built
            // against A1 as a carrier, so moving them would move a picture that should not move.
            if (picture.Placement == XLPicturePlacement.FreeFloating)
                continue;

            Move<TAxis>(picture.Markers[XLMarkerPosition.TopLeft], edit);

            // Only MoveAndSize has a second corner; under Move the picture keeps its pixel size and
            // the bottom-right marker is not maintained.
            if (picture.Placement == XLPicturePlacement.MoveAndSize)
                Move<TAxis>(picture.Markers[XLMarkerPosition.BottomRight], edit);
        }
    }

    /// <summary>
    /// Moves a picture's anchor corner. Same transform as a chart's or a note's; a marker holds a
    /// <see cref="Point"/> outright, so there is no coordinate base to reconcile.
    /// </summary>
    private static void Move<TAxis>(XLMarker? marker, in SheetEdit edit)
        where TAxis : struct, IGridAxis
    {
        if (marker is null)
            return;

        var anchor = marker.Anchor;
        marker.Anchor = GridShift.MoveArea<TAxis>(new Area(anchor, anchor), edit.Range, edit.Shift).FirstPoint;
    }

    /// <summary>
    /// Moves one anchor point, if the edited range covers its cross-axis line. The point is a single
    /// cell, so <see cref="GridShift.MoveArea{TAxis}"/> over a one-cell area gives the cross-axis
    /// coverage test and the index transform together.
    /// </summary>
    private static void Move<TAxis>(IXLDrawingPosition position, in SheetEdit edit, bool oneBased)
        where TAxis : struct, IGridAxis
    {
        var offset = oneBased ? 0 : 1;
        var row = position.Row + offset;
        var column = position.Column + offset;

        // A 0-based anchor on the very first row or column reaches the grid's edge; Point is 1-based
        // and cannot hold anything below it, and such an anchor cannot move up or left in any case.
        if (row < 1 || column < 1)
            return;

        var anchor = new Point(row, column);
        var moved = GridShift.MoveArea<TAxis>(new Area(anchor, anchor), edit.Range, edit.Shift);
        if (moved.FirstPoint == anchor)
            return;

        position.SetRow(moved.FirstPoint.Row - offset);
        position.SetColumn(moved.FirstPoint.Column - offset);
    }
}
