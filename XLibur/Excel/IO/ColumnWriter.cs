using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using XLibur.Extensions;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

internal static class ColumnWriter
{
    /// <param name="Columns">The <c>&lt;cols&gt;</c> element being built.</param>
    /// <param name="SheetColumnsByMin">The <c>&lt;col&gt;</c> elements built so far, keyed by <c>min</c>.</param>
    /// <param name="WorksheetStyleId">The worksheet's own style id.</param>
    /// <param name="WorksheetColumnWidth">
    /// The worksheet default width, already through <see cref="GetColumnWidth"/> and
    /// <c>SaveRound</c> by <see cref="SheetViewWriter.WriteSheetFormatProperties"/>.
    /// </param>
    /// <param name="RawWorksheetColumnWidth">
    /// The same default as the model holds it. <see cref="XLColumnSettings.Resolve"/> applies the
    /// width rule itself, so it must be handed the raw value; passing the resolved one rounds twice.
    /// </param>
    private readonly record struct ColumnWriteContext(
        Columns Columns,
        Dictionary<uint, Column> SheetColumnsByMin,
        uint WorksheetStyleId,
        double WorksheetColumnWidth,
        double RawWorksheetColumnWidth);

    /// <remarks>
    /// This took the whole ten-member save bag until spec 29. The shared style map is the only
    /// part of it this writer ever touched, and it only reads it.
    /// </remarks>
    internal static void WriteColumns(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        double worksheetColumnWidth,
        IReadOnlyDictionary<XLStyleValue, StyleInfo> sharedStyles)
    {
        var worksheetStyleId = sharedStyles[xlWorksheet.StyleValue].StyleId;
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
        var ctx = new ColumnWriteContext(columns, sheetColumnsByMin, worksheetStyleId, worksheetColumnWidth,
            xlWorksheet.ColumnWidth);

        var (minInColumnsCollection, maxInColumnsCollection) = GetColumnsRange(xlWorksheet);

        WritePreColumns(ctx, minInColumnsCollection);
        var maxCol = WriteMainColumns(ctx, xlWorksheet, minInColumnsCollection, maxInColumnsCollection, sharedStyles);
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
            UpdateColumn(WorksheetDefaultColumn(ctx, co, co), ctx.Columns, ctx.SheetColumnsByMin);
    }

    private static int WriteMainColumns(ColumnWriteContext ctx, XLWorksheet xlWorksheet,
        int minInColumnsCollection, int maxInColumnsCollection,
        IReadOnlyDictionary<XLStyleValue, StyleInfo> sharedStyles)
    {
        for (var co = minInColumnsCollection; co <= maxInColumnsCollection; co++)
        {
            var column = BuildColumnElement(ctx, xlWorksheet, co, sharedStyles);
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
        int columnNumber, IReadOnlyDictionary<XLStyleValue, StyleInfo> sharedStyles)
    {
        if (!xlWorksheet.Internals.ColumnsCollection.TryGetValue(columnNumber, out var col))
            return WorksheetDefaultColumn(ctx, (uint)columnNumber, (uint)columnNumber);

        // The raw width, not GetColumnWidth(col.Width).SaveRound() - Resolve applies that itself.
        var settings = XLColumnSettings.Resolve(
            (uint)columnNumber, (uint)columnNumber,
            sharedStyles[col.StyleValue].StyleId, col.Width,
            col.IsHidden, col.Collapsed, col.OutlineLevel);

        return ToColumnElement(settings);
    }

    /// <summary>
    /// A <c>&lt;col&gt;</c> carrying the worksheet's own style and default width, used to back-fill
    /// the columns either side of the ones the sheet actually configured.
    /// </summary>
    private static Column WorksheetDefaultColumn(ColumnWriteContext ctx, uint min, uint max)
        => ToColumnElement(XLColumnSettings.Resolve(
            min, max, ctx.WorksheetStyleId, ctx.RawWorksheetColumnWidth,
            hidden: false, collapsed: false, outlineLevel: 0));

    private static Column ToColumnElement(XLColumnSettings settings)
    {
        var column = new Column
        {
            Min = settings.Min,
            Max = settings.Max,
            Style = settings.StyleId,
            Width = settings.Width,
            CustomWidth = settings.CustomWidth ? true : null,
        };

        if (settings.Hidden)
            column.Hidden = true;
        if (settings.Collapsed)
            column.Collapsed = true;
        if (settings.OutlineLevel > 0)
            column.OutlineLevel = settings.OutlineLevel;

        return column;
    }

    private static void WritePostColumns(ColumnWriteContext ctx, int maxInColumnsCollection)
    {
        if (maxInColumnsCollection >= XLHelper.MaxColumnNumber || ctx.WorksheetStyleId == 0)
            return;

        ctx.Columns.AppendChild(
            WorksheetDefaultColumn(ctx, (uint)(maxInColumnsCollection + 1), (uint)XLHelper.MaxColumnNumber));
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

        sheetColumnsByMin.Remove(column.Min.Value);
        if (existingColumn.Min! + 1 > existingColumn.Max!)
        {
            columns.RemoveChild(existingColumn);
            columns.AppendChild(newColumn);
            sheetColumnsByMin.Add(newColumn.Min.Value, newColumn);
        }
        else
        {
            columns.AppendChild(newColumn);
            sheetColumnsByMin.Add(newColumn.Min.Value, newColumn);
            existingColumn.Min = existingColumn.Min! + 1;
            sheetColumnsByMin.Add(existingColumn.Min.Value, existingColumn);
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
