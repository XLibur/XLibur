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
        ColumnsCollection.Clear();
        RowsCollection.Clear();
        MergedRanges.RemoveAll();
    }
}
