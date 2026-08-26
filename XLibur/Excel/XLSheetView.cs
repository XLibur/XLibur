using System;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

internal sealed class XLSheetView : IXLSheetView, ISheetListener
{
    public XLSheetView(XLWorksheet worksheet)
    {
        Worksheet = worksheet;
        View = XLSheetViewOptions.Normal;

        ZoomScale = 100;
        ZoomScaleNormal = 100;
        ZoomScalePageLayoutView = 100;
        ZoomScaleSheetLayoutView = 100;
    }

    public XLSheetView(XLWorksheet worksheet, XLSheetView sheetView)
        : this(worksheet)
    {
        SplitRow = sheetView.SplitRow;
        SplitColumn = sheetView.SplitColumn;
        FreezePanes = sheetView.FreezePanes;
        TopLeftCellAddress = new XLAddress(Worksheet, sheetView.TopLeftCellAddress.RowNumber,
            sheetView.TopLeftCellAddress.ColumnNumber, sheetView.TopLeftCellAddress.FixedRow,
            sheetView.TopLeftCellAddress.FixedColumn);

        if (sheetView.PaneTopLeftCellAddress is { } pane)
            PaneTopLeftCellAddress = new XLAddress(Worksheet, pane.RowNumber, pane.ColumnNumber,
                pane.FixedRow, pane.FixedColumn);
    }

    public bool FreezePanes { get; set; }

    public int SplitColumn { get; set; }

    public int SplitRow { get; set; }

    IXLAddress IXLSheetView.TopLeftCellAddress
    {
        get => TopLeftCellAddress;
        set => TopLeftCellAddress = (XLAddress)value;
    }

    public XLAddress TopLeftCellAddress
    {
        get;
        set
        {
            if (value.HasWorksheet && !value.Worksheet!.Equals(Worksheet))
                throw new ArgumentException("The value should be on the same worksheet as the sheet view.");

            field = value;
        }
    }

    IXLAddress? IXLSheetView.PaneTopLeftCellAddress
    {
        get => PaneTopLeftCellAddress;
        set => PaneTopLeftCellAddress = (XLAddress?)value;
    }

    public XLAddress? PaneTopLeftCellAddress
    {
        get;
        set
        {
            if (value is { HasWorksheet: true } addr && !addr.Worksheet!.Equals(Worksheet))
                throw new ArgumentException("The value should be on the same worksheet as the sheet view.");

            field = value;
        }
    }

    public XLSheetViewOptions View { get; set; }

    IXLWorksheet IXLSheetView.Worksheet => Worksheet;

    public XLWorksheet Worksheet { get; internal set; }

    public int ZoomScale
    {
        get;
        set
        {
            field = value;
            switch (View)
            {
                case XLSheetViewOptions.Normal:
                    ZoomScaleNormal = value;
                    break;

                case XLSheetViewOptions.PageBreakPreview:
                    ZoomScalePageLayoutView = value;
                    break;

                case XLSheetViewOptions.PageLayout:
                    ZoomScaleSheetLayoutView = value;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), View, "Unsupported sheet view option.");
            }
        }
    }

    public int ZoomScaleNormal { get; set; }

    public int ZoomScalePageLayoutView { get; set; }

    public int ZoomScaleSheetLayoutView { get; set; }

    public void Freeze(int rows, int columns)
    {
        SplitRow = rows;
        SplitColumn = columns;
        FreezePanes = true;
    }

    public void FreezeColumns(int columns)
    {
        SplitColumn = columns;
        FreezePanes = true;
    }

    public void FreezeRows(int rows)
    {
        SplitRow = rows;
        FreezePanes = true;
    }

    public IXLSheetView SetView(XLSheetViewOptions value)
    {
        View = value;
        return this;
    }

    #region ISheetListener

    /// <summary>
    /// Grows or shrinks the frozen region when lines are inserted into or deleted from inside it.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see cref="SplitRow"/> is a <b>count</b> of frozen lines, not an address, which makes the
    /// panes the one feature in spec 33 whose transform is not an area transform — and is why the
    /// port carries <c>Shift</c> rather than only <c>Area</c>. It is also why
    /// <c>Area.TryInsertAreaAndShiftDown</c> is the wrong tool here: it answers "partial cover,
    /// don't move" for an insert that only partly covers a region, which is the wrong answer for a
    /// pane that should grow. <see cref="GridShift.MoveCount"/> does the arithmetic instead.
    /// </para>
    /// <para>
    /// The two bands are independent and each spans the whole of the other axis: freezing five rows
    /// freezes them across every column. So a row edit moves <see cref="SplitRow"/> only when it is
    /// an <em>entire row</em> edit — inserting cells into <c>B1:B5</c> and shifting them down is not
    /// a row insert and does not move the freeze, which is what Excel does.
    /// </para>
    /// <para>
    /// Deleting every line above the split leaves the count at zero, and a zero split is no split:
    /// <c>SheetViewWriter</c> removes the <c>pane</c> element outright when both counts are zero, so
    /// the freeze disappears rather than collapsing onto row 1.
    /// </para>
    /// </remarks>
    void ISheetListener.OnInsertAreaAndShiftDown(in SheetEdit edit) => MoveSplit<RowAxis>(in edit);

    void ISheetListener.OnInsertAreaAndShiftRight(in SheetEdit edit) => MoveSplit<ColumnAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftUp(in SheetEdit edit) => MoveSplit<RowAxis>(in edit);

    void ISheetListener.OnDeleteAreaAndShiftLeft(in SheetEdit edit) => MoveSplit<ColumnAxis>(in edit);

    private void MoveSplit<TAxis>(in SheetEdit edit)
        where TAxis : struct, IGridAxis
    {
        if (edit.Sheet != Worksheet)
            return;

        var axis = default(TAxis);
        if (!axis.IsEntireLine(edit.Range))
            return;

        var editFirstIndex = axis.IndexOf(edit.Range.RangeAddress.FirstAddress);

        var split = axis.ShiftsRows ? SplitRow : SplitColumn;
        if (split > 0)
        {
            var moved = GridShift.MoveCount(split, editFirstIndex, edit.Shift);
            if (axis.ShiftsRows)
                SplitRow = moved;
            else
                SplitColumn = moved;
        }

        MovePaneTopLeft<TAxis>(editFirstIndex, edit.Shift);
    }

    /// <summary>
    /// Moves the scrollable pane's anchor cell with the split it belongs to.
    /// </summary>
    /// <remarks>
    /// <c>SheetViewWriter</c> writes this address verbatim as <c>pane/@topLeftCell</c> whenever it is
    /// set, and derives it from the split only when it is null. <c>ScrollIntoView</c> sets it to
    /// <c>split + 1</c>, so moving the split without moving this leaves the two disagreeing — a
    /// sheet frozen at 5 rows with the anchor on <c>A6</c>, after inserting three rows at row 5,
    /// would write <c>ySplit="8"</c> against <c>topLeftCell="A6"</c>, an anchor inside the frozen
    /// band. It is an address rather than a count, so it takes <see cref="GridShift.MoveIndex"/>
    /// where the split takes <see cref="GridShift.MoveCount"/>.
    /// </remarks>
    private void MovePaneTopLeft<TAxis>(int editFirstIndex, int shift)
        where TAxis : struct, IGridAxis
    {
        if (PaneTopLeftCellAddress is not { } pane)
            return;

        var axis = default(TAxis);
        var moved = GridShift.MoveIndex(axis.IndexOf(pane), editFirstIndex, shift);
        if (moved == axis.IndexOf(pane))
            return;

        PaneTopLeftCellAddress = axis.AddressAt(
            Worksheet, moved, axis.CrossOf(pane), pane.FixedRow, pane.FixedColumn);
    }

    #endregion ISheetListener
}
