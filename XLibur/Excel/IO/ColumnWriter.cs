using XLibur.Excel.ContentManagers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Extensions;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

internal static class ColumnWriter
{
    private readonly record struct ColumnWriteContext(
        Columns Columns,
        Dictionary<uint, Column> SheetColumnsByMin,
        uint WorksheetStyleId,
        double WorksheetColumnWidth);

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
        var ctx = new ColumnWriteContext(columns, sheetColumnsByMin, worksheetStyleId, worksheetColumnWidth);

        var (minInColumnsCollection, maxInColumnsCollection) = GetColumnsRange(xlWorksheet);

        WritePreColumns(ctx, minInColumnsCollection);
        var maxCol = WriteMainColumns(ctx, xlWorksheet, minInColumnsCollection, maxInColumnsCollection, context);
        WritePostColumns(ctx, maxCol);

        CollapseColumns(columns, sheetColumnsByMin);

        if (!columns.Any())
        {
            worksheet.RemoveAllChildren<Columns>();
            cm.SetElement(XLWorksheetContents.Columns, null);
        }
    }

    private static (int min, int max) GetColumnsRange(XLWorksheet xlWorksheet)
    {
        var keys = xlWorksheet.Internals.ColumnsCollection.Keys;
        if (keys.Count == 0)
            return (1, 0);

        var min = int.MaxValue;
        var max = int.MinValue;
        foreach (var key in keys)
        {
            if (key < min) min = key;
            if (key > max) max = key;
        }

        return (min, max);
    }

    private static void WritePreColumns(ColumnWriteContext ctx, int minInColumnsCollection)
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
                Style = ctx.WorksheetStyleId,
                Width = ctx.WorksheetColumnWidth,
                CustomWidth = true
            };

            UpdateColumn(column, ctx.Columns, ctx.SheetColumnsByMin);
        }
    }

    private static int WriteMainColumns(ColumnWriteContext ctx, XLWorksheet xlWorksheet,
        int minInColumnsCollection, int maxInColumnsCollection, SaveContext context)
    {
        for (var co = minInColumnsCollection; co <= maxInColumnsCollection; co++)
        {
            var column = BuildColumnElement(ctx, xlWorksheet, co, context);
            UpdateColumn(column, ctx.Columns, ctx.SheetColumnsByMin);
        }

        foreach (
            var col in
            ctx.Columns.Elements<Column>().Where(c => c.Min! > (uint)(maxInColumnsCollection)).OrderBy(c => c.Min!.Value))
        {
            col.Style = ctx.WorksheetStyleId;
            col.Width = ctx.WorksheetColumnWidth;
            col.CustomWidth = true;

            if ((int)col.Max!.Value > maxInColumnsCollection)
                maxInColumnsCollection = (int)col.Max.Value;
        }

        return maxInColumnsCollection;
    }

    private static Column BuildColumnElement(ColumnWriteContext ctx, XLWorksheet xlWorksheet,
        int columnNumber, SaveContext context)
    {
        uint styleId;
        double columnWidth;
        var isHidden = false;
        var collapsed = false;
        var outlineLevel = 0;

        if (xlWorksheet.Internals.ColumnsCollection.TryGetValue(columnNumber, out var col))
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
            columnWidth = ctx.WorksheetColumnWidth;
        }

        var column = new Column
        {
            Min = (uint)columnNumber,
            Max = (uint)columnNumber,
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

        return column;
    }

    private static void WritePostColumns(ColumnWriteContext ctx, int maxInColumnsCollection)
    {
        if (maxInColumnsCollection >= XLHelper.MaxColumnNumber || ctx.WorksheetStyleId == 0)
            return;

        var column = new Column
        {
            Min = (uint)(maxInColumnsCollection + 1),
            Max = (uint)(XLHelper.MaxColumnNumber),
            Style = ctx.WorksheetStyleId,
            Width = ctx.WorksheetColumnWidth,
            CustomWidth = true
        };
        ctx.Columns.AppendChild(column);
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
