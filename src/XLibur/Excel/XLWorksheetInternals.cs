using System;
using XLibur.Excel.Rows;

namespace XLibur.Excel;

internal sealed class XLWorksheetInternals : IDisposable
{
    public XLWorksheetInternals(
        XLCellsCollection cellsCollection,
        XLColumnsCollection columnsCollection,
        XLRowsCollection rowsCollection,
        XLRanges mergedRanges
    )
    {
        CellsCollection = cellsCollection;
        ColumnsCollection = columnsCollection;
        RowsCollection = rowsCollection;
        MergedRanges = mergedRanges;
    }

    public XLCellsCollection CellsCollection { get; }

    public XLColumnsCollection ColumnsCollection { get; }

    public XLRowsCollection RowsCollection { get; }

    public XLRanges MergedRanges { get; internal set; }

    public void Dispose()
    {
        CellsCollection.ValueSlice.DereferenceSlice();
        CellsCollection.Clear();
        ReleaseRest();
    }

    /// <summary>
    /// Releases the sheet's storage when its workbook is disposed. Unlike <see cref="Dispose"/>,
    /// which deleting one sheet uses, it does not give back each cell's shared string one by one:
    /// the workbook drops the whole shared-string table next, so that work would be thrown away.
    /// </summary>
    public void DisposeWithWorkbook()
    {
        CellsCollection.Reset();
        ReleaseRest();
    }

    private void ReleaseRest()
    {
        ColumnsCollection.Clear();
        RowsCollection.Clear();
        MergedRanges.RemoveAll();
    }
}
