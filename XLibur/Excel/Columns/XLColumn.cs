using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Extensions;
using XLibur.Graphics;

namespace XLibur.Excel;

internal sealed class XLColumn : XLRangeBase, IXLColumn
{
    #region Private fields

    private readonly XLWorksheet _worksheet;
    private int _columnNumber;
    private int _outlineLevel;

    #endregion Private fields

    #region Constructor

    /// <summary>
    /// The direct constructor should only be used in <see cref="XLWorksheet.RangeFactory"/>.
    /// </summary>
    public XLColumn(XLWorksheet worksheet, int column)
        : base(worksheet.StyleValue)
    {
        _worksheet = worksheet;
        _columnNumber = column;

        Width = worksheet.ColumnWidth;
    }

    #endregion Constructor

    public override XLRangeAddress RangeAddress
    {
        get => XLRangeAddress.EntireColumn(_worksheet, _columnNumber);
        protected set => _columnNumber = value.FirstAddress.ColumnNumber;
    }

    public override XLWorksheet Worksheet => _worksheet;

    public override XLRangeType RangeType
    {
        get { return XLRangeType.Column; }
    }

    protected override IEnumerable<XLStylizedBase> Children
    {
        get
        {
            int column = ColumnNumber();
            foreach (XLCell cell in Worksheet.Internals.CellsCollection.GetCellsInColumn(column))
                yield return cell;
        }
    }

    public bool Collapsed { get; set; }

    #region IXLColumn Members

    public double Width { get; set; }

    IXLCells IXLColumn.Cells(string cellsInColumn) => Cells(cellsInColumn);

    IXLCells IXLColumn.Cells(int firstRow, int lastRow) => Cells(firstRow, lastRow);

    public void Delete()
    {
        int columnNumber = ColumnNumber();
        Delete(XLShiftDeletedCells.ShiftCellsLeft);
        Worksheet.DeleteColumn(columnNumber);
    }

    public new IXLColumn Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        base.Clear(clearOptions);
        return this;
    }

    public IXLCell Cell(int rowNumber)
    {
        return Cell(rowNumber, 1);
    }

    public override XLCells Cells(string cellsInColumn)
    {
        var retVal = new XLCells(false, XLCellsUsedOptions.All);
        var rangePairs = cellsInColumn.Split(',');
        foreach (string pair in rangePairs)
            retVal.Add(Range(pair.Trim()).RangeAddress);
        return retVal;
    }

    public override IXLCells Cells()
    {
        return Cells(true, XLCellsUsedOptions.All);
    }

    public override XLCells Cells(bool usedCellsOnly)
    {
        return usedCellsOnly
            ? Cells(true, XLCellsUsedOptions.AllContents)
            : Cells(FirstCellUsed()!.Address.RowNumber, LastCellUsed()!.Address.RowNumber);
    }

    public XLCells Cells(int firstRow, int lastRow)
    {
        return Cells(firstRow + ":" + lastRow);
    }

    public new IXLColumns InsertColumnsAfter(int numberOfColumns)
    {
        var columnNum = ColumnNumber();
        Worksheet.Internals.ColumnsCollection.ShiftColumnsRight(columnNum + 1, numberOfColumns);
        Worksheet.Column(columnNum).InsertColumnsAfterVoid(true, numberOfColumns);
        var newColumns = Worksheet.Columns(columnNum + 1, columnNum + numberOfColumns);
        CopyColumns(newColumns);
        return newColumns;
    }

    public new IXLColumns InsertColumnsBefore(int numberOfColumns)
    {
        int columnNum = ColumnNumber();
        if (columnNum > 1)
        {
            return Worksheet.Column(columnNum - 1).InsertColumnsAfter(numberOfColumns);
        }

        Worksheet.Internals.ColumnsCollection.ShiftColumnsRight(columnNum, numberOfColumns);
        Worksheet.Column(columnNum).InsertColumnsBeforeVoid(true, numberOfColumns);

        return Worksheet.Columns(columnNum, columnNum + numberOfColumns - 1);
    }

    private void CopyColumns(IXLColumns newColumns)
    {
        foreach (var newColumn in newColumns)
        {
            var internalColumn = Worksheet.Internals.ColumnsCollection[newColumn.ColumnNumber()];
            internalColumn.Width = Width;
            internalColumn.InnerStyle = InnerStyle;
            internalColumn.Collapsed = Collapsed;
            internalColumn.IsHidden = IsHidden;
            internalColumn._outlineLevel = OutlineLevel;
        }
    }

    public IXLColumn AdjustToContents()
    {
        return AdjustToContents(1);
    }

    public IXLColumn AdjustToContents(int startRow)
    {
        return AdjustToContents(startRow, XLHelper.MaxRowNumber);
    }

    public IXLColumn AdjustToContents(int startRow, int endRow)
    {
        return AdjustToContents(startRow, endRow, 0, double.MaxValue);
    }

    public IXLColumn AdjustToContents(double minWidth, double maxWidth)
    {
        return AdjustToContents(1, XLHelper.MaxRowNumber, minWidth, maxWidth);
    }

    public IXLColumn AdjustToContents(int startRow, double minWidth, double maxWidth)
    {
        return AdjustToContents(startRow, XLHelper.MaxRowNumber, minWidth, maxWidth);
    }

    public IXLColumn AdjustToContents(int startRow, int endRow, double minWidth, double maxWidth)
    {
        var engine = Worksheet.Workbook.GraphicEngine;
        var dpi = new Dpi(Worksheet.Workbook.DpiX, Worksheet.Workbook.DpiY);
        var columnWidthPx = CalculateMinColumnWidth(startRow, endRow, engine, dpi);

        // Maximum digit width, rounded to pixels, so Calibri at 11 pts returns 7 pixels MDW (the correct value)
        var mdw = (int)Math.Round(engine.GetMaxDigitWidth(Worksheet.Workbook.Style.Font, dpi.X));

        var minWidthInPx = Math.Ceiling(XLHelper.NoCToPixels(minWidth, mdw));
        if (columnWidthPx < minWidthInPx)
            columnWidthPx = (int)minWidthInPx;

        var maxWidthInPx = Math.Ceiling(XLHelper.NoCToPixels(maxWidth, mdw));
        if (columnWidthPx > maxWidthInPx)
            columnWidthPx = (int)maxWidthInPx;

        var colMaxWidth = XLHelper.PixelToNoC(columnWidthPx, mdw);

        // If there is nothing in the column, use worksheet column width.
        if (colMaxWidth <= 0)
            colMaxWidth = Worksheet.ColumnWidth;

        Width = colMaxWidth;

        return this;
    }

    /// <summary>
    /// Calculate column width in pixels according to the content of cells.
    /// </summary>
    /// <param name="startRow">First row number whose content is used for determination.</param>
    /// <param name="endRow">Last row number whose content is used for determination.</param>
    /// <param name="engine">Engine to determine size of glyphs.</param>
    /// <param name="dpi">DPI of the worksheet.</param>
    private int CalculateMinColumnWidth(int startRow, int endRow, IXLGraphicEngine engine, Dpi dpi)
    {
        var autoFilterRows = new List<int>();
        if (Worksheet.AutoFilter is { Range: not null })
            autoFilterRows.Add(Worksheet.AutoFilter.Range.FirstRow()!.RowNumber());

        autoFilterRows.AddRange(Worksheet.Tables.Where<XLTable>(t =>
                t.AutoFilter is { Range: not null }
                && !autoFilterRows.Contains(t.AutoFilter.Range.FirstRow()!.RowNumber()))
            .Select(t => t.AutoFilter.Range.FirstRow()!.RowNumber()));

        // Reusable buffer
        var glyphs = new List<GlyphBox>();
        XLStyle? cellStyle = null;
        var columnWidthPx = 0;
        foreach (var cell in Column(startRow, endRow).CellsUsed())
        {
            // Clear maintains capacity -> reduce need for GC
            glyphs.Clear();

            if (cell.IsMerged())
                continue;

            // Reuse styles if possible to reduce memory consumption
            if (cellStyle is null || cellStyle.Value != cell.StyleValue)
                cellStyle = (XLStyle)cell.Style;

            cell.GetGlyphBoxes(engine, dpi, glyphs);
            var textWidthPx = (int)Math.Ceiling(GetContentWidth(cellStyle.Alignment.TextRotation, glyphs));

            var scaledMdw = engine.GetMaxDigitWidth(cellStyle.Font, dpi.X);
            scaledMdw = Math.Round(scaledMdw, MidpointRounding.AwayFromZero);

            // Not sure about rounding, but larger is probably better, so use ceiling.
            // Due to mismatched rendering, add 3% instead of 1.75%, to have additional space.
            var oneSidePadding = (int)Math.Ceiling(textWidthPx * 0.03 + scaledMdw / 4);

            // Cell width if calculated as content width + padding on each side of a content.
            // The one side padding is roughly 1.75% of content + MDW/4.
            // The additional pixel is there for lines between cells.
            var cellWidthPx = textWidthPx + 2 * oneSidePadding + 1;

            if (autoFilterRows.Contains(cell.Address.RowNumber))
            {
                // Autofilter arrow is 16px at 96dpi, scaling through DPI, e.g. 20px at 120dpi
                cellWidthPx += (int)Math.Round(16d * dpi.X / 96d, MidpointRounding.AwayFromZero);
            }

            columnWidthPx = Math.Max(cellWidthPx, columnWidthPx);
        }

        return columnWidthPx;
    }

    private static double GetContentWidth(int textRotationDeg, List<GlyphBox> glyphs)
    {
        if (textRotationDeg == 0)
        {
            var maxTextWidth = 0d;
            var lineTextWidth = 0d;
            foreach (var glyph in glyphs)
            {
                if (!glyph.IsLineBreak)
                {
                    lineTextWidth += glyph.AdvanceWidth;
                    maxTextWidth = Math.Max(lineTextWidth, maxTextWidth);
                }
                else
                    lineTextWidth = 0;
            }

            return maxTextWidth;
        }

        if (textRotationDeg == 255)
        {
            // Glyphs are arranged vertically, top to bottom.
            return glyphs.Aggregate(0d, (current, grapheme) => Math.Max(grapheme.AdvanceWidth, current));
        }

        // Glyphs are rotated
        if (textRotationDeg > 90)
            textRotationDeg = 90 - textRotationDeg;

        var totalWidth = 0d;
        var maxHeight = 0d;
        foreach (var glyph in glyphs)
        {
            totalWidth += glyph.AdvanceWidth;
            maxHeight = Math.Max(maxHeight, glyph.LineHeight);
        }

        var projectedHeight = maxHeight * Math.Cos(XLHelper.DegToRad(90 - textRotationDeg));
        var projectedWidth = totalWidth * Math.Cos(XLHelper.DegToRad(textRotationDeg));
        return projectedWidth + projectedHeight;
    }

    public IXLColumn Hide()
    {
        IsHidden = true;
        return this;
    }

    public IXLColumn Unhide()
    {
        IsHidden = false;
        return this;
    }

    public bool IsHidden { get; set; }

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

    public IXLColumn Group()
    {
        return Group(false);
    }

    public IXLColumn Group(bool collapse)
    {
        if (OutlineLevel < 8)
            OutlineLevel += 1;

        Collapsed = collapse;
        return this;
    }

    public IXLColumn Group(int outlineLevel)
    {
        return Group(outlineLevel, false);
    }

    public IXLColumn Group(int outlineLevel, bool collapse)
    {
        OutlineLevel = outlineLevel;
        Collapsed = collapse;
        return this;
    }

    public IXLColumn Ungroup()
    {
        return Ungroup(false);
    }

    public IXLColumn Ungroup(bool fromAll)
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

    public IXLColumn Collapse()
    {
        Collapsed = true;
        return Hide();
    }

    public IXLColumn Expand()
    {
        Collapsed = false;
        return Unhide();
    }

    public int CellCount()
    {
        return RangeAddress.LastAddress.ColumnNumber - RangeAddress.FirstAddress.ColumnNumber + 1;
    }

    public IXLColumn Sort(XLSortOrder sortOrder = XLSortOrder.Ascending, bool matchCase = false,
        bool ignoreBlanks = true)
    {
        Sort(1, sortOrder, matchCase, ignoreBlanks);
        return this;
    }

    IXLRangeColumn IXLColumn.Column(int start, int end) => Column(start, end);

    IXLRangeColumn IXLColumn.CopyTo(IXLCell cell)
    {
        var copy = AsRange().CopyTo(cell);
        return copy.Column(1);
    }

    IXLRangeColumn IXLColumn.CopyTo(IXLRangeBase range)
    {
        var copy = AsRange().CopyTo(range);
        return copy.Column(1);
    }

    public IXLColumn CopyTo(IXLColumn column)
    {
        column.Clear();
        var newColumn = (XLColumn)column;
        newColumn.Width = Width;
        newColumn.InnerStyle = InnerStyle;
        newColumn.IsHidden = IsHidden;

        (this as XLRangeBase).CopyTo(column);

        return newColumn;
    }

    public XLRangeColumn Column(int start, int end)
    {
        return Range(start, 1, end, 1).Column(1);
    }

    public IXLRangeColumn Column(IXLCell start, IXLCell end)
    {
        return Column(start.Address.RowNumber, end.Address.RowNumber);
    }

    public IXLRangeColumns Columns(string columns)
    {
        var retVal = new XLRangeColumns();
        var columnPairs = columns.Split(',');
        foreach (string pair in columnPairs)
            AsRange().Columns(pair.Trim()).ForEach(retVal.Add);
        return retVal;
    }

    /// <summary>
    ///   Adds a vertical page break after this column.
    /// </summary>
    public IXLColumn AddVerticalPageBreak()
    {
        Worksheet.PageSetup.AddVerticalPageBreak(ColumnNumber());
        return this;
    }

    public IXLRangeColumn ColumnUsed(XLCellsUsedOptions options = XLCellsUsedOptions.AllContents)
    {
        return Column((this as IXLRangeBase).FirstCellUsed(options)!,
            (this as IXLRangeBase).LastCellUsed(options)!);
    }

    #endregion IXLColumn Members

    public override XLRange AsRange()
    {
        return Range(1, 1, XLHelper.MaxRowNumber, 1);
    }

    internal override void WorksheetRangeShiftedColumns(XLRange range, int columnsShifted)
    {
    }

    internal override void WorksheetRangeShiftedRows(XLRange range, int rowsShifted)
    {
        //do nothing
    }

    internal void SetColumnNumber(int column)
    {
        var oldAddress = RangeAddress;
        _columnNumber = column;
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
            string firstPart = arrRange[0];
            string secondPart = arrRange[1];
            rangeAddressToUse = FixColumnAddress(firstPart) + ":" + FixColumnAddress(secondPart);
        }
        else
            rangeAddressToUse = FixColumnAddress(rangeAddressStr);

        var rangeAddress = new XLRangeAddress(Worksheet, rangeAddressToUse);
        return Range(rangeAddress);
    }

    public IXLRangeColumn Range(int firstRow, int lastRow)
    {
        return Range(firstRow, 1, lastRow, 1).Column(1);
    }

    private XLColumn ColumnShift(int columnsToShift)
    {
        return Worksheet.Column(ColumnNumber() + columnsToShift);
    }

    #region XLColumn Left

    IXLColumn IXLColumn.ColumnLeft()
    {
        return ColumnLeft();
    }

    IXLColumn IXLColumn.ColumnLeft(int step)
    {
        return ColumnLeft(step);
    }

    public XLColumn ColumnLeft()
    {
        return ColumnLeft(1);
    }

    public XLColumn ColumnLeft(int step)
    {
        return ColumnShift(step * -1);
    }

    #endregion XLColumn Left

    #region XLColumn Right

    IXLColumn IXLColumn.ColumnRight()
    {
        return ColumnRight();
    }

    IXLColumn IXLColumn.ColumnRight(int step)
    {
        return ColumnRight(step);
    }

    public XLColumn ColumnRight()
    {
        return ColumnRight(1);
    }

    public XLColumn ColumnRight(int step)
    {
        return ColumnShift(step);
    }

    #endregion XLColumn Right

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
        return false;
    }

    public override bool IsEntireColumn()
    {
        return true;
    }
}
