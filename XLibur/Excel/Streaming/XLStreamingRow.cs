namespace XLibur.Excel.Streaming;

/// <summary>
/// The row currently being written by an <see cref="XLStreamingWorksheet"/>. Cells are written
/// left to right as they are added; the row cannot be revisited once it is closed.
/// </summary>
/// <remarks>
/// A <c>ref struct</c> so that a row cannot outlive the sheet write it belongs to, and so
/// building a row allocates nothing. Closing the row with <c>using</c> is optional - the row is
/// also closed when the next row starts or the worksheet completes - but it makes the row's
/// extent obvious at the call site.
/// </remarks>
/// <example>
/// <code>
/// using (var row = sheet.AddRow())
/// {
///     row.Cell("Widget");
///     row.Cell(12, highlight);
///     row.Skip(1);
///     row.Formula("B2*2", cachedValue: 24);
/// }
/// </code>
/// </example>
public readonly ref struct XLStreamingRow
{
    private readonly XLStreamingWorksheet _worksheet;

    internal XLStreamingRow(XLStreamingWorksheet worksheet, int rowNumber)
    {
        _worksheet = worksheet;
        RowNumber = rowNumber;
    }

    /// <summary>The 1-based number of this row.</summary>
    public int RowNumber { get; }

    /// <summary>
    /// Write a value into the next free column.
    /// </summary>
    public XLStreamingRow Cell(XLCellValue value)
    {
        _worksheet.WriteValueCell(RowNumber, value, null);
        return this;
    }

    /// <summary>
    /// Write a value into the next free column with its own style, overriding the row style.
    /// </summary>
    public XLStreamingRow Cell(XLCellValue value, IXLStyle? style)
    {
        _worksheet.WriteValueCell(RowNumber, value, style);
        return this;
    }

    /// <summary>
    /// Write a formula into the next free column. The formula string is stored verbatim - it is
    /// never parsed or evaluated - and is accepted with or without a leading <c>=</c>.
    /// </summary>
    /// <remarks>
    /// Without a <c>cachedValue</c> the cell has no result stored, so it shows as empty until
    /// Excel recalculates the sheet. Supply the value you already know to have it display
    /// immediately.
    /// </remarks>
    public XLStreamingRow Formula(string formula, XLCellValue cachedValue = default, IXLStyle? style = null)
    {
        _worksheet.WriteFormulaCell(RowNumber, formula, cachedValue, style);
        return this;
    }

    /// <summary>
    /// Leave the next <paramref name="columnCount"/> columns empty.
    /// </summary>
    public XLStreamingRow Skip(int columnCount)
    {
        _worksheet.SkipCells(RowNumber, columnCount);
        return this;
    }

    /// <summary>
    /// Continue writing at a specific column, 1-based. The column must be at or after the next
    /// free one, since cells are written left to right.
    /// </summary>
    public XLStreamingRow At(int columnNumber)
    {
        _worksheet.MoveToColumn(RowNumber, columnNumber);
        return this;
    }

    /// <summary>
    /// Close the row. Optional: the row also closes when the next row starts or the worksheet
    /// completes.
    /// </summary>
    public void Dispose() => _worksheet.EndRow(RowNumber);
}
