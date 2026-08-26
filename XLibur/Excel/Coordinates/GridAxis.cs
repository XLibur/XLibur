
namespace XLibur.Excel.Coordinates;

/// <summary>
/// One of the grid's two axes, as a value the JIT can specialise over. Every method projects a
/// point, an area or a range address onto the axis being operated on ("index") and the axis it
/// spans ("cross"), so an algorithm can be written once and bound to either direction.
/// </summary>
/// <remarks>
/// <para>
/// Implementations are empty <c>readonly struct</c>s and are always passed as a generic type
/// argument constrained <c>where TAxis : struct, IGridAxis</c>. That is what makes the calls
/// devirtualise: the JIT specialises the method body per axis and the receiver's exact type is
/// known at every call site. Never hold an <see cref="IGridAxis"/> in a field, a local or a
/// parameter typed as the interface — that is the one form that reintroduces the dispatch this
/// shape exists to remove, and spec 21 already paid to learn what a mis-shaped struct costs here.
/// </para>
/// <para>
/// For a row insert the index axis is rows and the cross axis is columns; for a column insert it is
/// the other way round. The transposition is not uniform — the formatting pass in
/// <see cref="XLRangeInsertHelper"/> runs on the <em>cross</em> axis while the shift itself runs on
/// the index axis — which is exactly what the two longhand copies kept getting wrong.
/// </para>
/// </remarks>
internal interface IGridAxis
{
    /// <summary>1,048,576 for the row axis, 16,384 for the column axis.</summary>
    int MaxIndex { get; }

    /// <summary>The extent of a single line on this axis: 16,384 for a row, 1,048,576 for a column.</summary>
    int MaxCross { get; }

    /// <summary>"rows" or "columns", for the argument-out-of-range message.</summary>
    string LineNoun { get; }

    /// <summary>True on the row axis. <see cref="XLFormulaShiftPass"/> takes this as a flag.</summary>
    bool ShiftsRows { get; }

    int IndexOf(IXLAddress address);
    int CrossOf(IXLAddress address);

    /// <summary>Builds a point from a position on this axis and one on the cross axis.</summary>
    Point PointAt(int index, int cross);

    /// <summary>
    /// Builds an address from a position on this axis and one on the cross axis. The two fixed flags
    /// are <em>not</em> transposed — they name the row and the column, whichever axis is the index.
    /// </summary>
    XLAddress AddressAt(XLWorksheet worksheet, int index, int cross, bool fixedRow, bool fixedColumn);

    /// <summary><c>IsEntireRow()</c> on the row axis, <c>IsEntireColumn()</c> on the column axis.</summary>
    bool IsEntireLine(XLRangeBase range);

    /// <summary><c>MaxRowUsed</c> on the row axis, <c>MaxColumnUsed</c> on the column axis.</summary>
    int MaxUsedIndex(XLCellsCollection cells);

    /// <summary>Copies a line's height (row axis) or width (column axis) onto another line.</summary>
    void CopyLineSize(XLWorksheet worksheet, int fromIndex, int toIndex);

    void ShiftSparklines(XLSparklineGroups groups, Area area, int shift);
    void InsertAreaAndShift(XLCellsCollection cells, Area area);
    void NotifyRangeShifted(XLWorksheet worksheet, XLRange range, int shift);

    /// <summary>The worksheet range spanning <paramref name="firstIndex"/>..<paramref name="lastIndex"/>
    /// on this axis and <paramref name="firstCross"/>..<paramref name="lastCross"/> on the other.</summary>
    IXLRange RangeFor(XLWorksheet worksheet, int firstIndex, int lastIndex, int firstCross, int lastCross);

    /// <summary>The line immediately before the range on this axis — the column to its left, or the
    /// row above it — used as the model when formatting from the preceding line.</summary>
    IXLRangeBase ModelLineBefore(IXLRange range);

    /// <summary>The style of the model line's cell at a cross-axis offset.</summary>
    IXLStyle ModelCellStyle(IXLRangeBase modelLine, int cross);

    /// <summary>Styles the range's cross-axis line at <paramref name="cross"/>. On a row insert this
    /// styles a <em>column</em> of the inserted block, and vice versa.</summary>
    void SetCrossLineStyle(IXLRange range, int cross, IXLStyle style);

    /// <summary>The last used position on the cross axis, or -1 when the range has no used cells.</summary>
    int LastUsedCross(IXLRange range, XLCellsUsedOptions options);

    /// <summary>The stored style of a cross-axis line, falling back to the worksheet's own style.</summary>
    IXLStyle CrossLineStyle(XLWorksheet worksheet, int cross);
}

/// <summary>The row axis: rows are inserted, deleted and shifted; each row spans 16,384 columns.</summary>
internal readonly struct RowAxis : IGridAxis
{
    public int MaxIndex => XLHelper.MaxRowNumber;
    public int MaxCross => XLHelper.MaxColumnNumber;
    public string LineNoun => "rows";
    public bool ShiftsRows => true;

    public int IndexOf(IXLAddress address) => address.RowNumber;
    public int CrossOf(IXLAddress address) => address.ColumnNumber;

    public Point PointAt(int index, int cross) => new(index, cross);

    public XLAddress AddressAt(XLWorksheet worksheet, int index, int cross, bool fixedRow, bool fixedColumn)
        => new(worksheet, index, cross, fixedRow, fixedColumn);

    public bool IsEntireLine(XLRangeBase range) => range.IsEntireRow();

    public int MaxUsedIndex(XLCellsCollection cells) => cells.MaxRowUsed;

    public void CopyLineSize(XLWorksheet worksheet, int fromIndex, int toIndex)
        => worksheet.Row(toIndex).Height = worksheet.Row(fromIndex).Height;

    public void ShiftSparklines(XLSparklineGroups groups, Area area, int shift)
        => groups.ShiftRows(area, shift);

    public void InsertAreaAndShift(XLCellsCollection cells, Area area)
        => cells.InsertAreaAndShiftDown(area);

    public void NotifyRangeShifted(XLWorksheet worksheet, XLRange range, int shift)
        => worksheet.NotifyRangeShiftedRows(range, shift);

    public IXLRange RangeFor(XLWorksheet worksheet, int firstIndex, int lastIndex, int firstCross, int lastCross)
        => worksheet.Range(firstIndex, firstCross, lastIndex, lastCross);

    public IXLRangeBase ModelLineBefore(IXLRange range) => range.FirstRow()!.RowAbove();

    public IXLStyle ModelCellStyle(IXLRangeBase modelLine, int cross)
        => ((IXLRangeRow)modelLine).Cell(cross).Style;

    public void SetCrossLineStyle(IXLRange range, int cross, IXLStyle style)
        => range.Column(cross).Style = style;

    public int LastUsedCross(IXLRange range, XLCellsUsedOptions options)
        => range.LastColumnUsed(options)?.ColumnNumber() ?? -1;

    public IXLStyle CrossLineStyle(XLWorksheet worksheet, int cross)
        => worksheet.Internals.ColumnsCollection.TryGetValue(cross, out var column)
            ? column.Style
            : worksheet.Style;
}

/// <summary>The column axis: columns are inserted, deleted and shifted; each column spans 1,048,576 rows.</summary>
internal readonly struct ColumnAxis : IGridAxis
{
    public int MaxIndex => XLHelper.MaxColumnNumber;
    public int MaxCross => XLHelper.MaxRowNumber;
    public string LineNoun => "columns";
    public bool ShiftsRows => false;

    public int IndexOf(IXLAddress address) => address.ColumnNumber;
    public int CrossOf(IXLAddress address) => address.RowNumber;

    public Point PointAt(int index, int cross) => new(cross, index);

    public XLAddress AddressAt(XLWorksheet worksheet, int index, int cross, bool fixedRow, bool fixedColumn)
        => new(worksheet, cross, index, fixedRow, fixedColumn);

    public bool IsEntireLine(XLRangeBase range) => range.IsEntireColumn();

    public int MaxUsedIndex(XLCellsCollection cells) => cells.MaxColumnUsed;

    public void CopyLineSize(XLWorksheet worksheet, int fromIndex, int toIndex)
        => worksheet.Column(toIndex).Width = worksheet.Column(fromIndex).Width;

    public void ShiftSparklines(XLSparklineGroups groups, Area area, int shift)
        => groups.ShiftColumns(area, shift);

    public void InsertAreaAndShift(XLCellsCollection cells, Area area)
        => cells.InsertAreaAndShiftRight(area);

    public void NotifyRangeShifted(XLWorksheet worksheet, XLRange range, int shift)
        => worksheet.NotifyRangeShiftedColumns(range, shift);

    public IXLRange RangeFor(XLWorksheet worksheet, int firstIndex, int lastIndex, int firstCross, int lastCross)
        => worksheet.Range(firstCross, firstIndex, lastCross, lastIndex);

    public IXLRangeBase ModelLineBefore(IXLRange range) => range.FirstColumn()!.ColumnLeft();

    public IXLStyle ModelCellStyle(IXLRangeBase modelLine, int cross)
        => ((IXLRangeColumn)modelLine).Cell(cross).Style;

    public void SetCrossLineStyle(IXLRange range, int cross, IXLStyle style)
        => range.Row(cross).Style = style;

    public int LastUsedCross(IXLRange range, XLCellsUsedOptions options)
        => range.LastRowUsed(options)?.RowNumber() ?? -1;

    public IXLStyle CrossLineStyle(XLWorksheet worksheet, int cross)
        => worksheet.Internals.RowsCollection.TryGetValue(cross, out var row)
            ? row.Style
            : worksheet.Style;
}
