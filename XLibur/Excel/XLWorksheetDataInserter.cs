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
        var rows = PrepareRows(reader, addHeadings, transpose);
        var propCount = reader.GetPropertiesCount();

        var valueSlice = worksheet.Internals.CellsCollection.ValueSlice;
        var styleSlice = worksheet.Internals.CellsCollection.StyleSlice;

        var rowBuffer = new List<XLCellValue>();
        var maximumColumn = origin.Column;
        var rowNumber = origin.Row;
        foreach (var row in rows)
        {
            rowBuffer.AddRange(row);

            for (var i = rowBuffer.Count; i < propCount; ++i)
                rowBuffer.Add(Blank.Value);

            maximumColumn = Math.Max(origin.Column + rowBuffer.Count - 1, maximumColumn);
            if (maximumColumn > XLHelper.MaxColumnNumber || rowNumber > XLHelper.MaxRowNumber)
                throw new ArgumentException("Data would write out of the sheet.");

            WriteRow(rowBuffer, rowNumber, origin.Column, valueSlice, styleSlice);

            rowBuffer.Clear();
            rowNumber++;
        }

        var lastRow = Math.Max(rowNumber - 1, origin.Row);
        var insertedArea = new XLSheetRange(origin, new XLSheetPoint(lastRow, maximumColumn));

        foreach (var table in worksheet.Tables)
            table.RefreshFieldsFromCells(insertedArea);

        worksheet.Workbook.CalcEngine.MarkDirty(worksheet, insertedArea);

        return worksheet.Range(
            insertedArea.FirstPoint.Row,
            insertedArea.FirstPoint.Column,
            insertedArea.LastPoint.Row,
            insertedArea.LastPoint.Column);
    }

    private static IEnumerable<IEnumerable<XLCellValue>> PrepareRows(IInsertDataReader reader, bool addHeadings, bool transpose)
    {
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
            rows = TransposeJaggedArray(rows);

        return rows;
    }

    private void WriteRow(List<XLCellValue> rowBuffer, int rowNumber, int startColumn,
        ValueSlice valueSlice, Slice<XLStyleValue?> styleSlice)
    {
        var column = startColumn;
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
