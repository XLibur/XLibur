using XLibur.Excel.InsertData;
using System;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel;

/// <summary>
/// Handles data and table insertion operations for a worksheet.
/// </summary>
internal sealed class XLWorksheetDataInserter(XLWorksheet worksheet)
{
    public IXLTable InsertTable(XLSheetPoint origin, IInsertDataReader reader, string? tableName, bool createTable, bool addHeadings, bool transpose)
    {
        if (createTable && worksheet.Tables.Any<XLTable>(t => t.Area.Contains(origin)))
            throw new InvalidOperationException($"This cell '{origin}' is already part of a table.");

        var range = InsertData(origin, reader, addHeadings, transpose);

        if (createTable)
            // Create a table and save it in the file
            return tableName == null ? range.CreateTable() : range.CreateTable(tableName);
        // Create a table but keep it in memory. Saved file will contain only "raw" data and column headers
        return tableName == null ? range.AsTable() : range.AsTable(tableName);
    }

    public XLRange InsertData(XLSheetPoint origin, IInsertDataReader reader, bool addHeadings, bool transpose)
    {
        // Prepare data. Heading is basically just another row of data, so unify it.
        var rows = reader.GetRecords();
        var propCount = reader.GetPropertiesCount();
        if (addHeadings)
        {
            var headings = new XLCellValue[propCount];
            for (var i = 0; i < propCount; i++)
                headings[i] = reader.GetPropertyName(i);

            rows = new[] { headings }.Concat(rows);
        }

        if (transpose)
        {
            rows = TransposeJaggedArray(rows);
        }

        var valueSlice = worksheet.Internals.CellsCollection.ValueSlice;
        var styleSlice = worksheet.Internals.CellsCollection.StyleSlice;

        // A buffer to avoid multiple enumerations of the source.
        var rowBuffer = new List<XLCellValue>();
        var maximumColumn = origin.Column;
        var rowNumber = origin.Row;
        foreach (var row in rows)
        {
            rowBuffer.AddRange(row);

            // InsertData should also clear data and if row doesn't have enough data,
            // fill in the rest. Only fill up to the props to be consistent. We can't
            // know how long any next row will be, so props are used as a source of truth
            // for which columns should be cleared.
            for (var i = rowBuffer.Count; i < propCount; ++i)
                rowBuffer.Add(Blank.Value);

            // Each row can have different number of values, so we have to check every row.
            maximumColumn = Math.Max(origin.Column + rowBuffer.Count - 1, maximumColumn);
            if (maximumColumn > XLHelper.MaxColumnNumber || rowNumber > XLHelper.MaxRowNumber)
                throw new ArgumentException("Data would write out of the sheet.");

            var column = origin.Column;
            foreach (var t in rowBuffer)
            {
                var value = t;
                var point = new XLSheetPoint(rowNumber, column);
                var modifiedStyle = worksheet.GetStyleForValue(value, point);
                if (modifiedStyle is not null)
                {
                    if (value.IsText && value.GetText()[0] == '\'')
                        value = value.GetText().Substring(1);

                    styleSlice.Set(point, modifiedStyle);
                }

                valueSlice.SetCellValue(point, value);
                column++;
            }

            rowBuffer.Clear();
            rowNumber++;
        }

        // If there is no row, rowNumber is kept at origin instead of last row + 1 .
        var lastRow = Math.Max(rowNumber - 1, origin.Row);
        var insertedArea = new XLSheetRange(origin, new XLSheetPoint(lastRow, maximumColumn));

        // If inserted area affected a table, we must fix headings and totals, because these values
        // are duplicated. Basically the table values are the truth and cells are a reflection of the
        // truth, but here we inserted shadow first.
        foreach (var table in worksheet.Tables)
            table.RefreshFieldsFromCells(insertedArea);

        // Invalidate only once, not for every cell.
        worksheet.Workbook.CalcEngine.MarkDirty(worksheet, insertedArea);

        // Return area that contains all inserted cells, no matter how jagged were data.
        return worksheet.Range(
            insertedArea.FirstPoint.Row,
            insertedArea.FirstPoint.Column,
            insertedArea.LastPoint.Row,
            insertedArea.LastPoint.Column);
    }

    // Rather memory inefficient, but the original code also materialized
    // data through Linq/required multiple enumerations.
    private static List<List<XLCellValue>> TransposeJaggedArray(IEnumerable<IEnumerable<XLCellValue>> enumerable)
    {
        var destination = new List<List<XLCellValue>>();

        var sourceRow = 1;
        foreach (var row in enumerable)
        {
            var sourceColumn = 1;
            foreach (var sourceValue in row)
            {
                // The original has `sourceValue` at [sourceRow, sourceColumn]
                var destinationRowCount = destination.Count;
                if (sourceColumn > destinationRowCount)
                    destination.Add([]);

                // There can be jagged arrays and the destination can have spaces between columns.
                var destinationRow = destination[sourceColumn - 1];
                while (destinationRow.Count < sourceRow - 1)
                    destinationRow.Add(Blank.Value);

                destinationRow.Add(sourceValue);
                sourceColumn++;
            }

            sourceRow++;
        }

        return destination;
    }
}
