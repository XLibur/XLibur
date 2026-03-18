using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Extensions;
using XLibur.Graphics;

namespace XLibur.Excel.Rows;

internal sealed class XLRow : XLRangeBase, IXLRow
{
    #region Private fields

    private readonly XLWorksheet _worksheet;
    private int _rowNumber;

    /// <summary>
    /// Don't use directly, use properties.
    /// </summary>
    private XlRowFlags _flags;

    private double _height;
    private int _outlineLevel;

    #endregion Private fields

    #region Constructor

    /// <summary>
    /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
    /// </summary>
    public XLRow(XLWorksheet worksheet, int row)
        : base(worksheet.StyleValue)
    {
        _worksheet = worksheet;
        _rowNumber = row;

        _height = worksheet.RowHeight;
    }

    #endregion Constructor

    public override XLRangeAddress RangeAddress
    {
        get => XLRangeAddress.EntireRow(_worksheet, _rowNumber);
        protected set => _rowNumber = value.FirstAddress.RowNumber;
    }

    public override XLWorksheet Worksheet => _worksheet;

    public override XLRangeType RangeType => XLRangeType.Row;

    protected override IEnumerable<XLStylizedBase> Children
    {
        get
        {
            var row = RowNumber();

            foreach (var cell in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
                yield return cell;
        }
    }

    public bool Collapsed
    {
        get => _flags.HasFlag(XlRowFlags.Collapsed);
        set
        {
            if (value)
                _flags |= XlRowFlags.Collapsed;
            else
                _flags &= ~XlRowFlags.Collapsed;
        }
    }

    /// <summary>
    /// Distance in pixels from the bottom of the cells in the current row to the typographical
    /// baseline of the cell content if, hypothetically, the zoom level for the sheet containing
    /// this row is 100 percent and the cell has bottom-alignment formatting.
    /// </summary>
    /// <remarks>
    /// If the attribute is set, it sets customHeight to true even if the customHeight is explicitly
    /// set to false. Custom height means no auto-sizing by Excel on load, so if a row has this
    /// attribute, it stops Excel from auto-sizing the height of a row to fit the content on load.
    /// </remarks>
    public double? DyDescent { get; set; }

    /// <summary>
    /// Should cells in the row display phonetic? This doesn't actually affect whether the phonetic are
    /// shown in the row, that depends entirely on the <see cref="IXLCell.ShowPhonetic"/> property
    /// of a cell. This property determines whether a new cell in the row will have its phonetic turned on
    /// (and also the state of the "Show or hide phonetic" in Excel when the whole row is selected).
    /// Default is <c>false</c>.
    /// </summary>
    public bool ShowPhonetic
    {
        get => _flags.HasFlag(XlRowFlags.ShowPhonetic);
        set
        {
            if (value)
                _flags |= XlRowFlags.ShowPhonetic;
            else
                _flags &= ~XlRowFlags.ShowPhonetic;
        }
    }

    /// <summary>
    /// Does row have an individual height or is it derived from the worksheet <see cref="XLWorksheet.RowHeight"/>?
    /// </summary>
    public bool HeightChanged
    {
        get => _flags.HasFlag(XlRowFlags.HeightChanged);
        private set
        {
            if (value)
                _flags |= XlRowFlags.HeightChanged;
            else
                _flags &= ~XlRowFlags.HeightChanged;
        }
    }

    #region IXLRow Members

    public double Height
    {
        get => _height;
        set
        {
            HeightChanged = true;
            _height = value;
        }
    }

    /// <summary>
    /// Set height without marking <see cref="HeightChanged"/>. Used during loading
    /// to assign the worksheet default height without flagging a custom override.
    /// </summary>
    internal void SetHeightNoFlag(double height)
    {
        _height = height;
    }

    IXLCells IXLRow.Cells(string cellsInRow) => Cells(cellsInRow);

    IXLCells IXLRow.Cells(int firstColumn, int lastColumn) => Cells(firstColumn, lastColumn);

    public void ClearHeight()
    {
        Height = Worksheet.RowHeight;
        HeightChanged = false;
    }

    public void Delete()
    {
        var rowNumber = RowNumber();
        AsRange().Delete(XLShiftDeletedCells.ShiftCellsUp);
        Worksheet.DeleteRow(rowNumber);
    }

    public new IXLRows InsertRowsBelow(int numberOfRows)
    {
        var rowNum = RowNumber();
        Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum + 1, numberOfRows);
        var asRange = Worksheet.Row(rowNum).AsRange();
        asRange.InsertRowsBelowVoid(true, numberOfRows);

        var newRows = Worksheet.Rows(rowNum + 1, rowNum + numberOfRows);

        CopyRows(newRows);

        return newRows;
    }

    private void CopyRows(IXLRows newRows)
    {
        foreach (var newRow in newRows)
        {
            var internalRow = Worksheet.Internals.RowsCollection[newRow.RowNumber()];
            internalRow._height = Height;
            internalRow.InnerStyle = InnerStyle;
            internalRow.Collapsed = Collapsed;
            internalRow.IsHidden = IsHidden;
            internalRow._outlineLevel = OutlineLevel;
        }
    }

    public new IXLRows InsertRowsAbove(int numberOfRows)
    {
        var rowNum = RowNumber();
        if (rowNum > 1)
        {
            return Worksheet.Row(rowNum - 1).InsertRowsBelow(numberOfRows);
        }

        Worksheet.Internals.RowsCollection.ShiftRowsDown(rowNum, numberOfRows);
        var asRange = Worksheet.Row(rowNum).AsRange();
        asRange.InsertRowsAboveVoid(true, numberOfRows);

        return Worksheet.Rows(rowNum, rowNum + numberOfRows - 1);
    }

    public new IXLRow Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public IXLCell Cell(int columnNumber)
    {
        return Cell(1, columnNumber);
    }

    public override XLCell Cell(string cellAddressInRange)
    {
        return Cell(1, cellAddressInRange);
    }

    IXLCell IXLRow.Cell(string columnLetter)
    {
        return Cell(columnLetter);
    }

    public override IXLCells Cells()
    {
        return Cells(true, XLCellsUsedOptions.All);
    }

    public override XLCells Cells(bool usedCellsOnly)
    {
        return usedCellsOnly
            ? Cells(true, XLCellsUsedOptions.AllContents)
            : Cells(FirstCellUsed()!.Address.ColumnNumber, LastCellUsed()!.Address.ColumnNumber);
    }

    public override XLCells Cells(string cells)
    {
        var retVal = new XLCells(false, XLCellsUsedOptions.AllContents);
        var rangePairs = cells.Split(',');
        foreach (var pair in rangePairs)
            retVal.Add(Range(pair.Trim()).RangeAddress);
        return retVal;
    }

    public XLCells Cells(int firstColumn, int lastColumn)
    {
        return Cells(firstColumn + ":" + lastColumn);
    }

    public IXLCells Cells(string firstColumn, string lastColumn)
    {
        return Cells(XLHelper.GetColumnNumberFromLetter(firstColumn) + ":"
                                                                     + XLHelper.GetColumnNumberFromLetter(lastColumn));
    }

    public IXLRow AdjustToContents()
    {
        return AdjustToContents(1);
    }

    public IXLRow AdjustToContents(int startColumn)
    {
        return AdjustToContents(startColumn, XLHelper.MaxColumnNumber);
    }

    public IXLRow AdjustToContents(int startColumn, int endColumn)
    {
        return AdjustToContents(startColumn, endColumn, 0, double.MaxValue);
    }

    public IXLRow AdjustToContents(double minHeight, double maxHeight)
    {
        return AdjustToContents(1, XLHelper.MaxColumnNumber, minHeight, maxHeight);
    }

    public IXLRow AdjustToContents(int startColumn, double minHeight, double maxHeight)
    {
        return AdjustToContents(startColumn, XLHelper.MaxColumnNumber, minHeight, maxHeight);
    }

    public IXLRow AdjustToContents(int startColumn, int endColumn, double minHeightPt, double maxHeightPt)
    {
        var engine = Worksheet.Workbook.GraphicEngine;
        var dpi = new Dpi(Worksheet.Workbook.DpiX, Worksheet.Workbook.DpiY);

        var rowHeightPx = CalculateMinRowHeight(startColumn, endColumn, engine, dpi);

        var rowHeightPt = XLHelper.PixelsToPoints(rowHeightPx, dpi.Y);
        if (rowHeightPt <= 0)
            rowHeightPt = Worksheet.RowHeight;

        if (minHeightPt > rowHeightPt)
            rowHeightPt = minHeightPt;

        if (maxHeightPt < rowHeightPt)
            rowHeightPt = maxHeightPt;

        Height = rowHeightPt;

        return this;
    }

    private int CalculateMinRowHeight(int startColumn, int endColumn, IXLGraphicEngine engine, Dpi dpi)
    {
        var glyphs = new List<GlyphBox>();
        XLStyle? cellStyle = null;
        var rowHeightPx = 0;
        foreach (var cell in Row(startColumn, endColumn).CellsUsed().Cast<XLCell>())
        {
            // Clear maintains capacity -> reduce need for GC
            glyphs.Clear();

            if (cell.IsMerged())
                continue;

            // Reuse styles if possible to reduce memory consumption
            if (cellStyle is null || cellStyle.Value != cell.StyleValue)
                cellStyle = (XLStyle)cell.Style;

            cell.GetGlyphBoxes(engine, dpi, glyphs);
            var cellHeightPx = (int)Math.Ceiling(GetContentHeight(cellStyle.Alignment.TextRotation, glyphs));

            rowHeightPx = Math.Max(cellHeightPx, rowHeightPx);
        }

        return rowHeightPx;
    }

    private static double GetContentHeight(int textRotationDeg, List<GlyphBox> glyphs)
    {
        switch (textRotationDeg)
        {
            case 0:
                {
                    var textHeight = 0d;
                    var lineMaxHeight = 0d;
                    foreach (var glyph in glyphs)
                    {
                        if (!glyph.IsLineBreak)
                        {
                            var cellHeightPx = glyph.LineHeight;
                            lineMaxHeight = Math.Max(cellHeightPx, lineMaxHeight);
                        }
                        else
                        {
                            // At the end of each line, add height of the line to total height.
                            // Use glyph.LineHeight as fallback for empty lines (consecutive/leading/trailing newlines).
                            var effectiveLineHeight = lineMaxHeight > 0 ? lineMaxHeight : glyph.LineHeight;
                            textHeight += effectiveLineHeight;
                            lineMaxHeight = 0d;
                        }
                    }

                    // If the last line ends without EOL, it must be also counted
                    textHeight += lineMaxHeight;

                    return textHeight;
                }
            case 255:
                {
                    // Glyphs are vertically aligned.
                    var textHeight = glyphs.Sum(static g => g.LineHeight);
                    return textHeight;
                }
        }

        // Rotated text
        var width = 0d;
        var height = 0d;
        foreach (var glyph in glyphs)
        {
            width += glyph.AdvanceWidth;
            height = Math.Max(glyph.LineHeight, height);
        }

        var projectedWidth = Math.Sin(XLHelper.DegToRad(textRotationDeg)) * width;
        var projectedHeight = Math.Cos(XLHelper.DegToRad(textRotationDeg)) * height;
        return projectedWidth + projectedHeight;
    }

    public IXLRow Hide()
    {
        IsHidden = true;
        return this;
    }

    public IXLRow Unhide()
    {
        IsHidden = false;
        return this;
    }

    public bool IsHidden
    {
        get => _flags.HasFlag(XlRowFlags.IsHidden);
        set
        {
            if (value)
                _flags |= XlRowFlags.IsHidden;
            else
                _flags &= ~XlRowFlags.IsHidden;
        }
    }

    public int OutlineLevel
    {
        get => _outlineLevel;
        set
        {
            if (value is < 0 or > 8)
                throw new ArgumentOutOfRangeException(nameof(value), "Outline level must be between 0 and 8.");

            Worksheet.IncrementColumnOutline(value);
            Worksheet.DecrementColumnOutline(_outlineLevel);
            _outlineLevel = value;
        }
    }

    public IXLRow Group()
    {
        return Group(false);
    }

    public IXLRow Group(int outlineLevel)
    {
        return Group(outlineLevel, false);
    }

    public IXLRow Group(bool collapse)
    {
        if (OutlineLevel < 8)
            OutlineLevel += 1;

        Collapsed = collapse;
        return this;
    }

    public IXLRow Group(int outlineLevel, bool collapse)
    {
        OutlineLevel = outlineLevel;
        Collapsed = collapse;
        return this;
    }

    public IXLRow Ungroup()
    {
        return Ungroup(false);
    }

    public IXLRow Ungroup(bool fromAll)
    {
        if (fromAll)
            OutlineLevel = 0;
        else
        {
            if (OutlineLevel > 0)
                OutlineLevel -= 1;
        }

        return this;
    }

    public IXLRow Collapse()
    {
        Collapsed = true;
        return Hide();
    }

    public IXLRow Expand()
    {
        Collapsed = false;
        return Unhide();
    }

    public int CellCount()
    {
        return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
    }

    public new IXLRow Sort()
    {
        return SortLeftToRight();
    }

    public new IXLRow SortLeftToRight(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false,
        bool ignoreBlanks = true)
    {
        base.SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    IXLRangeRow IXLRow.CopyTo(IXLCell cell)
    {
        var copy = AsRange().CopyTo(cell);
        return copy.Row(1);
    }

    IXLRangeRow IXLRow.CopyTo(IXLRangeBase range)
    {
        var copy = AsRange().CopyTo(range);
        return copy.Row(1);
    }

    public IXLRow CopyTo(IXLRow row)
    {
        row.Clear();
        var newRow = (XLRow)row;
        newRow._height = _height;
        newRow.HeightChanged = HeightChanged;
        newRow.InnerStyle = GetStyle();
        newRow.IsHidden = IsHidden;

        AsRange().CopyTo(row);

        return newRow;
    }

    public IXLRangeRow Row(int start, int end)
    {
        return Range(1, start, 1, end).Row(1);
    }

    public IXLRangeRow Row(IXLCell start, IXLCell end)
    {
        return Row(start.Address.ColumnNumber, end.Address.ColumnNumber);
    }

    public IXLRangeRows Rows(string columns)
    {
        var retVal = new XLRangeRows();
        var rowPairs = columns.Split(',');
        foreach (var pair in rowPairs)
            AsRange().Rows(pair.Trim()).ForEach(retVal.Add);

        return retVal;
    }

    public IXLRow AddHorizontalPageBreak()
    {
        Worksheet.PageSetup.AddHorizontalPageBreak(RowNumber());
        return this;
    }

    public IXLRangeRow RowUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents)
    {
        return Row((this as IXLRangeBase).FirstCellUsed(options)!,
            (this as IXLRangeBase).LastCellUsed(options)!);
    }

    #endregion IXLRow Members

    public override XLRange AsRange()
    {
        return Range(1, 1, 1, XLHelper.MaxColumnNumber);
    }

    internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
    {
        //do nothing
    }

    internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
    {
        // rows are shifted by XLRowCollection
    }

    internal void SetRowNumber(int row)
    {
        var oldAddress = RangeAddress;
        _rowNumber = row;
        OnRangeAddressChanged(oldAddress, RangeAddress);
    }

    public override XLRange Range(string rangeAddressStr)
    {
        string rangeAddressToUse;
        if (rangeAddressStr.Contains(':') || rangeAddressStr.Contains('-'))
        {
            if (rangeAddressStr.Contains('-'))
                rangeAddressStr = rangeAddressStr.Replace('-', ':');

            var arrRange = rangeAddressStr.Split(':');
            var firstPart = arrRange[0];
            var secondPart = arrRange[1];
            rangeAddressToUse = FixRowAddress(firstPart) + ":" + FixRowAddress(secondPart);
        }
        else
            rangeAddressToUse = FixRowAddress(rangeAddressStr);

        var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
        return Range(rangeAddress);
    }

    internal void SetStyleNoColumns(IXLStyle value)
    {
        InnerStyle = value;

        var row = RowNumber();
        foreach (var c in Worksheet.Internals.CellsCollection.GetCellsInRow(row))
            c.InnerStyle = value;
    }

    private XLRow RowShift(int rowsToShift)
    {
        return Worksheet.Row(RowNumber() + rowsToShift);
    }

    #region XLRow Above

    IXLRow IXLRow.RowAbove()
    {
        return RowAbove();
    }

    IXLRow IXLRow.RowAbove(int step)
    {
        return RowAbove(step);
    }

    public XLRow RowAbove()
    {
        return RowAbove(1);
    }

    public XLRow RowAbove(int step)
    {
        return RowShift(step * -1);
    }

    #endregion XLRow Above

    #region XLRow Below

    IXLRow IXLRow.RowBelow()
    {
        return RowBelow();
    }

    IXLRow IXLRow.RowBelow(int step)
    {
        return RowBelow(step);
    }

    public XLRow RowBelow()
    {
        return RowBelow(1);
    }

    public XLRow RowBelow(int step)
    {
        return RowShift(step);
    }

    #endregion XLRow Below

    public override bool IsEmpty()
    {
        return IsEmpty(XLCellsUsedOptions.AllContents);
    }

    public override bool IsEmpty(XLCellsUsedOptions options)
    {
        if (options.HasFlag(XLCellsUsedOptions.NormalFormats) &&
            !StyleValue.Equals(Worksheet.StyleValue))
            return false;

        return base.IsEmpty(options);
    }

    public override bool IsEntireRow()
    {
        return true;
    }

    public override bool IsEntireColumn()
    {
        return false;
    }

    /// <summary>
    /// Flag enum to save space instead of wasting byte for each flag.
    /// </summary>
    [Flags]
    private enum XlRowFlags : byte
    {
        Collapsed = 1,
        IsHidden = 2,
        ShowPhonetic = 4,
        HeightChanged = 8
    }
}
