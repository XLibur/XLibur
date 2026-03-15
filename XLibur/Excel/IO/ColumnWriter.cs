using XLibur.Excel.ContentManagers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Extensions;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

internal sealed class ColumnWriter
{
    internal static void WriteColumns(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        double worksheetColumnWidth,
        SaveContext context)
    {
        var worksheetStyleId = context.SharedStyles[xlWorksheet.StyleValue].StyleId;
        if (xlWorksheet.Internals.CellsCollection.IsEmpty &&
            xlWorksheet.Internals.ColumnsCollection.Count == 0
            && worksheetStyleId == 0)
        {
            worksheet.RemoveAllChildren<Columns>();
            return;
        }

        if (!worksheet.Elements<Columns>().Any())
        {
            var previousElement = cm.GetPreviousElementFor(XLWorksheetContents.Columns);
            worksheet.InsertAfter(new Columns(), previousElement);
        }

        var columns = worksheet.Elements<Columns>().First();
        cm.SetElement(XLWorksheetContents.Columns, columns);

        var sheetColumnsByMin = columns.Elements<Column>().ToDictionary(c => c.Min!.Value, c => c);

        var (minInColumnsCollection, maxInColumnsCollection) = GetColumnsRange(xlWorksheet);

        WritePreColumns(columns, sheetColumnsByMin, minInColumnsCollection, worksheetStyleId, worksheetColumnWidth);
        var maxCol = WriteMainColumns(columns, sheetColumnsByMin, xlWorksheet, minInColumnsCollection,
            maxInColumnsCollection, worksheetStyleId, worksheetColumnWidth, context);
        WritePostColumns(columns, maxCol, worksheetStyleId, worksheetColumnWidth);

        CollapseColumns(columns, sheetColumnsByMin);

        if (!columns.Any())
        {
            worksheet.RemoveAllChildren<Columns>();
            cm.SetElement(XLWorksheetContents.Columns, null);
        }
    }

    private static (int min, int max) GetColumnsRange(XLWorksheet xlWorksheet)
    {
        if (xlWorksheet.Internals.ColumnsCollection.Count > 0)
        {
            return (xlWorksheet.Internals.ColumnsCollection.Keys.Min(),
                    xlWorksheet.Internals.ColumnsCollection.Keys.Max());
        }

        return (1, 0);
    }

    private static void WritePreColumns(Columns columns, Dictionary<uint, Column> sheetColumnsByMin,
        int minInColumnsCollection, uint worksheetStyleId, double worksheetColumnWidth)
    {
        if (minInColumnsCollection <= 1)
            return;

        UInt32Value min = 1;
        UInt32Value max = (uint)(minInColumnsCollection - 1);

        for (var co = min; co <= max; co++)
        {
            var column = new Column
            {
                Min = co,
                Max = co,
                Style = worksheetStyleId,
                Width = worksheetColumnWidth,
                CustomWidth = true
            };

            UpdateColumn(column, columns, sheetColumnsByMin);
        }
    }

    private static int WriteMainColumns(Columns columns, Dictionary<uint, Column> sheetColumnsByMin,
        XLWorksheet xlWorksheet, int minInColumnsCollection, int maxInColumnsCollection,
        uint worksheetStyleId, double worksheetColumnWidth, SaveContext context)
    {
        for (var co = minInColumnsCollection; co <= maxInColumnsCollection; co++)
        {
            uint styleId;
            double columnWidth;
            var isHidden = false;
            var collapsed = false;
            var outlineLevel = 0;
            if (xlWorksheet.Internals.ColumnsCollection.TryGetValue(co, out var col))
            {
                styleId = context.SharedStyles[col.StyleValue].StyleId;
                columnWidth = GetColumnWidth(col.Width).SaveRound();
                isHidden = col.IsHidden;
                collapsed = col.Collapsed;
                outlineLevel = col.OutlineLevel;
            }
            else
            {
                styleId = context.SharedStyles[xlWorksheet.StyleValue].StyleId;
                columnWidth = worksheetColumnWidth;
            }

            var column = new Column
            {
                Min = (uint)co,
                Max = (uint)co,
                Style = styleId,
                Width = columnWidth,
                CustomWidth = true
            };

            if (isHidden)
                column.Hidden = true;
            if (collapsed)
                column.Collapsed = true;
            if (outlineLevel > 0)
                column.OutlineLevel = (byte)outlineLevel;

            UpdateColumn(column, columns, sheetColumnsByMin);
        }

        foreach (
            var col in
            columns.Elements<Column>().Where(c => c.Min! > (uint)(maxInColumnsCollection)).OrderBy(c => c.Min!.Value))
        {
            col.Style = worksheetStyleId;
            col.Width = worksheetColumnWidth;
            col.CustomWidth = true;

            if ((int)col.Max!.Value > maxInColumnsCollection)
                maxInColumnsCollection = (int)col.Max.Value;
        }

        return maxInColumnsCollection;
    }

    private static void WritePostColumns(Columns columns, int maxInColumnsCollection,
        uint worksheetStyleId, double worksheetColumnWidth)
    {
        if (maxInColumnsCollection >= XLHelper.MaxColumnNumber || worksheetStyleId == 0)
            return;

        var column = new Column
        {
            Min = (uint)(maxInColumnsCollection + 1),
            Max = (uint)(XLHelper.MaxColumnNumber),
            Style = worksheetStyleId,
            Width = worksheetColumnWidth,
            CustomWidth = true
        };
        columns.AppendChild(column);
    }

    internal static double GetColumnWidth(double columnWidth)
    {
        return Math.Min(255.0, Math.Max(0.0, columnWidth + XLConstants.ColumnWidthOffset));
    }

    private static void CollapseColumns(Columns columns, Dictionary<uint, Column> sheetColumns)
    {
        uint lastMin = 1;
        var count = sheetColumns.Count;
        var arr = sheetColumns.OrderBy(kp => kp.Key).ToArray();
        for (var i = 0; i < count; i++)
        {
            var kp = arr[i];
            if (i + 1 != count && ColumnsAreEqual(kp.Value, arr[i + 1].Value)) continue;

            var newColumn = (Column)kp.Value.CloneNode(true);
            newColumn.Min = lastMin;
            var newColumnMax = newColumn.Max!.Value;
            var columnsToRemove =
                columns.Elements<Column>().Where(co => co.Min! >= lastMin && co.Max! <= newColumnMax).Select(co => co)
                    .ToList();
            columnsToRemove.ForEach(c => columns.RemoveChild(c));

            columns.AppendChild(newColumn);
            lastMin = kp.Key + 1;
        }
    }

    private static void UpdateColumn(Column column, Columns columns, Dictionary<uint, Column> sheetColumnsByMin)
    {
        if (!sheetColumnsByMin.TryGetValue(column.Min!.Value, out var newColumn))
        {
            newColumn = (Column)column.CloneNode(true);
            columns.AppendChild(newColumn);
            sheetColumnsByMin.Add(column.Min.Value, newColumn);
        }
        else
        {
            UpdateExistingColumn(column, columns, sheetColumnsByMin);
        }
    }

    private static void UpdateExistingColumn(Column column, Columns columns, Dictionary<uint, Column> sheetColumnsByMin)
    {
        var existingColumn = sheetColumnsByMin[column.Min!.Value];
        var newColumn = (Column)existingColumn.CloneNode(true);
        newColumn.Min = column.Min;
        newColumn.Max = column.Max;
        newColumn.Style = column.Style;
        newColumn.Width = column.Width!.SaveRound();
        newColumn.CustomWidth = column.CustomWidth;

        newColumn.Hidden = column.Hidden != null ? true : null;
        newColumn.Collapsed = column.Collapsed != null ? true : null;
        newColumn.OutlineLevel = column.OutlineLevel != null && column.OutlineLevel > 0
            ? (byte)column.OutlineLevel
            : null;

        sheetColumnsByMin.Remove(column.Min!.Value);
        if (existingColumn.Min! + 1 > existingColumn.Max!)
        {
            columns.RemoveChild(existingColumn);
            columns.AppendChild(newColumn);
            sheetColumnsByMin.Add(newColumn.Min.Value, newColumn);
        }
        else
        {
            columns.AppendChild(newColumn);
            sheetColumnsByMin.Add(newColumn.Min!.Value, newColumn);
            existingColumn.Min = existingColumn.Min! + 1;
            sheetColumnsByMin.Add(existingColumn.Min!.Value, existingColumn);
        }
    }

    private static bool ColumnsAreEqual(Column left, Column right)
    {
        return NullableValuesEqual(left.Style, right.Style)
            && NullableDoublesEqual(left.Width, right.Width)
            && NullableValuesEqual(left.Hidden, right.Hidden)
            && NullableValuesEqual(left.Collapsed, right.Collapsed)
            && NullableValuesEqual(left.OutlineLevel, right.OutlineLevel);
    }

    private static bool NullableValuesEqual<T>(OpenXmlSimpleValue<T>? left, OpenXmlSimpleValue<T>? right)
        where T : struct
    {
        if (left == null && right == null) return true;
        if (left == null || right == null) return false;
        return left.Value.Equals(right.Value);
    }

    private static bool NullableDoublesEqual(DoubleValue? left, DoubleValue? right)
    {
        if (left == null && right == null) return true;
        if (left == null || right == null) return false;
        return Math.Abs(left.Value - right.Value) < XLHelper.Epsilon;
    }
}
