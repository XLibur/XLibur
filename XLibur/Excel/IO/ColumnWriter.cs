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
    /// <param name="DefaultColumn">
    /// The worksheet default, resolved once: its own style and width, no flags. Both the columns
    /// back-filled either side of the configured range and the trailing fix-up in
    /// <see cref="WriteMainColumns"/> take their width from here, so the default is decided in one
    /// place rather than stored twice and kept in agreement by hand. Its <c>Min</c> and <c>Max</c>
    /// are placeholders; each use sets its own.
    /// </param>
    /// <param name="BackFillsColumns">
    /// Are the columns the model does not hold written with the sheet's width and style? A loaded
    /// sheet whose own style is the default does not need them: a column with no <c>&lt;col&gt;</c>
    /// has the sheet's width and style 0 already, and Excel writes none (#722).
    /// </param>
    private readonly record struct ColumnWriteContext(
        Columns Columns,
        Dictionary<uint, Column> SheetColumnsByMin,
        uint WorksheetStyleId,
        XLColumnSettings DefaultColumn,
        bool BackFillsColumns);

    /// <remarks>
    /// This took the whole ten-member save bag until spec 29. The shared style map is the only
    /// part of it this writer ever touched, and it only reads it.
    /// </remarks>
    /// <param name="worksheet">The worksheet DOM being written.</param>
    /// <param name="cm">Where each worksheet element is.</param>
    /// <param name="xlWorksheet">The sheet.</param>
    /// <param name="sharedStyles">The style ids of the save.</param>
    /// <param name="partIsEmpty">Is the part new, rather than one the load read?</param>
    internal static void WriteColumns(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        IReadOnlyDictionary<XLStyleValue, StyleInfo> sharedStyles,
        bool partIsEmpty)
    {
        var worksheetStyleId = sharedStyles[xlWorksheet.StyleValue].StyleId;
        // A loaded <cols> that holds only ranges to the last column is written back even here: the
        // load took such a range as the sheet's width rather than as columns, and that width is not
        // written as a defaultColWidth (#709).
        if (xlWorksheet.Internals.CellsCollection.IsEmpty &&
            xlWorksheet.Internals.ColumnsCollection.Count == 0
            && worksheetStyleId == 0
            && !HoldsOnlyTheSheetWidth(worksheet.Elements<Columns>().FirstOrDefault()))
        {
            worksheet.RemoveAllChildren<Columns>();
            cm.SetElement(XLWorksheetContents.Columns, null);
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
        // The raw model width, not a pre-resolved one: Resolve applies GetColumnWidth and SaveRound
        // itself, and handing it an already-resolved width rounds twice.
        var ctx = new ColumnWriteContext(columns, sheetColumnsByMin, worksheetStyleId,
            XLColumnSettings.Resolve(1, 1, worksheetStyleId, xlWorksheet.ColumnWidth,
                hidden: false, collapsed: false, outlineLevel: 0),
            BackFillsColumns: partIsEmpty || worksheetStyleId != 0);

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

    /// <summary>
    /// Does the loaded <c>&lt;cols&gt;</c> hold only ranges that run to the last column with nothing
    /// but a width and a style, which the load takes as the sheet's own?
    /// </summary>
    private static bool HoldsOnlyTheSheetWidth(Columns? columns) =>
        columns is not null && columns.Elements<Column>().Any() && columns.Elements<Column>().All(c =>
            c.Max?.Value == XLHelper.MaxColumnNumber
            && c.Hidden is null or { Value: false }
            && c.Collapsed is null or { Value: false }
            && c.OutlineLevel is null);

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
            if (BacksFill(ctx, co))
                UpdateColumn(WorksheetDefaultColumn(ctx, co, co), ctx.Columns, ctx.SheetColumnsByMin);
        }
    }

    /// <summary>
    /// Is a column the model does not hold written with the sheet's width and style? Always where
    /// the context back-fills, and where the file had a <c>&lt;col&gt;</c>.
    /// </summary>
    private static bool BacksFill(ColumnWriteContext ctx, uint column) =>
        ctx.BackFillsColumns || ctx.SheetColumnsByMin.ContainsKey(column);

    private static int WriteMainColumns(ColumnWriteContext ctx, XLWorksheet xlWorksheet,
        int minInColumnsCollection, int maxInColumnsCollection,
        IReadOnlyDictionary<XLStyleValue, StyleInfo> sharedStyles)
    {
        for (var co = minInColumnsCollection; co <= maxInColumnsCollection; co++)
        {
            var column = BuildColumnElement(ctx, xlWorksheet, co, sharedStyles);
            if (column is not null)
                UpdateColumn(column, ctx.Columns, ctx.SheetColumnsByMin);
        }

        foreach (
            var col in
            ctx.Columns.Elements<Column>().Where(c => c.Min! > (uint)(maxInColumnsCollection)).OrderBy(c => c.Min!.Value))
        {
            // The sheet's own style and width, so customWidth stays as the file had it: Excel
            // writes a column that only carries the sheet's width without one.
            if (col.Style is not { HasValue: true } loadedStyle
                || !xlWorksheet.Internals.ColumnsCollection.IsLoadedStyle(loadedStyle.Value, xlWorksheet.StyleValue))
                col.Style = SchemaDefault.UInt(col.Style, ctx.WorksheetStyleId, 0);

            col.Width = ctx.DefaultColumn.Width is { } sheetWidth ? SchemaDefault.Measure(col.Width, sheetWidth) : null;

            if ((int)col.Max!.Value > maxInColumnsCollection)
                maxInColumnsCollection = (int)col.Max.Value;
        }

        return maxInColumnsCollection;
    }

    /// <summary>
    /// The style index of the file's <c>&lt;col&gt;</c> at <paramref name="column"/>, while the column
    /// has the style the load read from it; otherwise null. The save finds a style at the first index
    /// that holds it, and a file can hold the same style at more than one (#722).
    /// </summary>
    private static uint? LoadedStyleId(ColumnWriteContext ctx, XLWorksheet xlWorksheet, uint column, XLStyleValue style) =>
        ctx.SheetColumnsByMin.TryGetValue(column, out var loaded)
        && loaded.Style is { HasValue: true } index
        && xlWorksheet.Internals.ColumnsCollection.IsLoadedStyle(index.Value, style)
            ? index.Value
            : null;

    /// <returns>The column's <c>&lt;col&gt;</c>, or null when it is not written.</returns>
    private static Column? BuildColumnElement(ColumnWriteContext ctx, XLWorksheet xlWorksheet,
        int columnNumber, IReadOnlyDictionary<XLStyleValue, StyleInfo> sharedStyles)
    {
        if (!xlWorksheet.Internals.ColumnsCollection.TryGetValue(columnNumber, out var col))
        {
            return BacksFill(ctx, (uint)columnNumber)
                ? WorksheetDefaultColumn(ctx, (uint)columnNumber, (uint)columnNumber)
                : null;
        }

        // The raw width, not GetColumnWidth(col.Width).SaveRound() - Resolve applies that itself.
        var settings = XLColumnSettings.Resolve(
            (uint)columnNumber, (uint)columnNumber,
            LoadedStyleId(ctx, xlWorksheet, (uint)columnNumber, col.StyleValue) ?? sharedStyles[col.StyleValue].StyleId,
            col.Width, col.IsHidden, col.Collapsed, col.OutlineLevel);

        // A column with the sheet's width has no customWidth, as Excel writes it. The model does not
        // know whether a width was set, only what it is.
        if (settings.Width is { } width && ctx.DefaultColumn.Width is { } sheetWidth
                                        && Math.Abs(width - sheetWidth) < XLHelper.Epsilon)
            settings = settings with { CustomWidth = false };

        return ToColumnElement(ctx, settings);
    }

    /// <summary>
    /// A <c>&lt;col&gt;</c> carrying the worksheet's own style and default width, used to back-fill
    /// the columns either side of the ones the sheet actually configured. Its width is the sheet's,
    /// so it has no <c>customWidth</c>, as Excel writes such a column.
    /// </summary>
    private static Column WorksheetDefaultColumn(ColumnWriteContext ctx, uint min, uint max)
        => ToColumnElement(ctx, ctx.DefaultColumn with { Min = min, Max = max, CustomWidth = false });

    private static Column ToColumnElement(ColumnWriteContext ctx, XLColumnSettings settings)
    {
        var column = new Column
        {
            Min = settings.Min,
            Max = settings.Max,
            // Style 0 is what a missing style means, so it is left out, as Excel leaves it out. Not
            // on a sheet with a style of its own: a column without a style loads with the sheet's.
            Style = settings.StyleId is 0 && ctx.WorksheetStyleId == 0 ? null : settings.StyleId,
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

    /// <summary>
    /// Joins each run of adjacent columns that are written alike into one <c>&lt;col&gt;</c>.
    /// </summary>
    /// <remarks>
    /// Only columns that touch are joined, and a run starts at its own first column. A column with
    /// no <c>&lt;col&gt;</c> between two alike is not given their width (#722).
    /// </remarks>
    private static void CollapseColumns(Columns columns, Dictionary<uint, Column> sheetColumns)
    {
        uint? runMin = null;
        var count = sheetColumns.Count;
        var arr = sheetColumns.OrderBy(kp => kp.Key).ToArray();
        for (var i = 0; i < count; i++)
        {
            var kp = arr[i];
            runMin ??= kp.Key;
            if (i + 1 != count
                && arr[i + 1].Key == kp.Value.Max!.Value + 1
                && ColumnsAreEqual(kp.Value, arr[i + 1].Value))
                continue;

            var newColumn = (Column)kp.Value.CloneNode(true);
            var min = runMin.Value;
            newColumn.Min = min;
            var newColumnMax = newColumn.Max!.Value;
            var columnsToRemove =
                columns.Elements<Column>().Where(co => co.Min! >= min && co.Max! <= newColumnMax).Select(co => co)
                    .ToList();
            columnsToRemove.ForEach(c => columns.RemoveChild(c));

            columns.AppendChild(newColumn);
            runMin = null;
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

        // A style="0" the file had stays; the column writes none of its own.
        newColumn.Style = column.Style ?? SchemaDefault.UInt(existingColumn.Style, 0, 0);
        // The width and customWidth stay as the file had them while the width is the same: the clone
        // already carries both. The file can spell a width with more digits than a save keeps
        // (#722). The model does not know whether a width was set, only what it is.
        var width = column.Width!.SaveRound();
        if (!SchemaDefault.SameMeasure(existingColumn.Width, width))
        {
            newColumn.Width = width;
            newColumn.CustomWidth = column.CustomWidth;
        }

        // A flag the file spelled out stays, set or not (#722).
        newColumn.Hidden = SchemaDefault.Bool(newColumn.Hidden, column.Hidden != null, false);
        newColumn.Collapsed = SchemaDefault.Bool(newColumn.Collapsed, column.Collapsed != null, false);
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
               && NullableValuesEqual(left.CustomWidth, right.CustomWidth)
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

    /// <summary>
    /// Are the widths the same at the six decimals a save writes? A loaded width keeps the file's
    /// digits, such as 14.85546875, and is the same width as 14.855469 (#722).
    /// </summary>
    private static bool NullableDoublesEqual(DoubleValue? left, DoubleValue? right)
    {
        if (left == null && right == null) return true;
        if (left == null || right == null) return false;
        return Math.Abs(Math.Round(left.Value, 6) - Math.Round(right.Value, 6)) < XLHelper.Epsilon;
    }
}
