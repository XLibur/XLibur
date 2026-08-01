namespace XLibur.Excel;

internal struct XLMiscSliceContent
{
    internal XLComment? Comment { get; set; }

    /// <summary>
    /// The root of the cell's comment thread, if any. Mutually exclusive with <see cref="Comment"/>:
    /// a cell shows either a legacy note or a thread, never both. Lives here rather than in a side
    /// table so that row/column shifting, swapping and clearing — all of which run over the slices —
    /// treat a thread exactly like a note.
    /// </summary>
    internal XLThreadedComment? ThreadedComment { get; set; }

    internal uint? CellMetaIndex { get; set; }

    internal uint? ValueMetaIndex { get; set; }

    internal bool HasPhonetic { get; set; }

    internal XLCellImage? CellImage { get; set; }
}
