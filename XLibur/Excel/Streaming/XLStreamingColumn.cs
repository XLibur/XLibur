namespace XLibur.Excel.Streaming;

/// <summary>
/// Presentation settings for a column (or a range of columns) of a streamed worksheet, as
/// returned by <see cref="XLStreamingWorksheet.Column"/> and
/// <see cref="XLStreamingWorksheet.Columns"/>.
/// </summary>
/// <remarks>
/// Columns are written before <c>sheetData</c>, so these must be set before the first row is
/// appended to the sheet.
/// </remarks>
public sealed class XLStreamingColumn
{
    internal XLStreamingColumn(int firstColumn, int lastColumn)
    {
        FirstColumn = firstColumn;
        LastColumn = lastColumn;
    }

    /// <summary>First column this applies to, 1-based.</summary>
    public int FirstColumn { get; }

    /// <summary>Last column this applies to, 1-based and inclusive.</summary>
    public int LastColumn { get; }

    /// <summary>
    /// Column width in Excel's character units, or <c>null</c> to leave the default width.
    /// </summary>
    public double? Width { get; set; }

    /// <summary>Whether the column is hidden.</summary>
    public bool Hidden { get; set; }

    /// <summary>Outline (grouping) level, 0 for none.</summary>
    public int OutlineLevel { get; set; }

    /// <summary>Whether the outline group the column belongs to is collapsed.</summary>
    public bool Collapsed { get; set; }

    /// <summary>
    /// Style applied to cells in the column that carry no style of their own.
    /// </summary>
    public IXLStyle? Style { get; set; }
}
