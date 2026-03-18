namespace XLibur.Excel;

/// <summary>
/// Represents an in-cell image ("Place in Cell" feature in Excel 365+).
/// Stored on a cell's MiscSlice, references a workbook-level image blob.
/// </summary>
internal sealed class XLCellImage(int workbookImageIndex, string altText)
{
    /// <summary>
    /// 0-based index into <see cref="XLInCellImageStore"/>.
    /// </summary>
    internal int WorkbookImageIndex { get; } = workbookImageIndex;

    /// <summary>
    /// Alt text for accessibility. May be empty.
    /// </summary>
    internal string AltText { get; } = altText;
}
