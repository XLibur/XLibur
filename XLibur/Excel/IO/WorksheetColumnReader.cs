using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Extensions;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads the <c>&lt;cols&gt;</c> element of a worksheet part.
/// </summary>
/// <remarks>
/// Separate from <see cref="WorksheetSheetDataReader"/> because <c>&lt;cols&gt;</c> is a worksheet
/// element in its own right rather than part of <c>&lt;sheetData&gt;</c>, which is how the loader
/// already dispatches it.
/// </remarks>
internal static class WorksheetColumnReader
{
    internal static void LoadColumns(StylesheetData styles, XLWorksheet ws, Columns columns)
    {
        var wsDefaultColumn =
            columns.Elements<Column>().FirstOrDefault(c => c.Max?.Value == XLHelper.MaxColumnNumber);

        if (wsDefaultColumn != null && wsDefaultColumn.Width != null)
            ws.ColumnWidth = wsDefaultColumn.Width - XLConstants.ColumnWidthOffset;

        var styleIndexDefault = wsDefaultColumn != null && wsDefaultColumn.Style != null
            ? int.Parse(wsDefaultColumn.Style.InnerText!)
            : -1;
        if (styleIndexDefault >= 0)
            StyleDecoder.ApplyStyle(ws, styleIndexDefault, styles);

        foreach (var col in columns.Elements<Column>())
        {
            if (col.Max?.Value == XLHelper.MaxColumnNumber) continue;

            LoadColumn(col, ws, styles);
        }
    }

    private static void LoadColumn(Column col, XLWorksheet ws, StylesheetData styles)
    {
        var xlColumns = (XLColumns)ws.Columns((int)col.Min!.Value, (int)col.Max!.Value);
        if (col.Width != null)
        {
            var width = col.Width - XLConstants.ColumnWidthOffset;
            xlColumns.Width = width;
        }
        else
            xlColumns.Width = ws.ColumnWidth;

        if (col.Hidden != null && col.Hidden)
            xlColumns.Hide();

        if (col.Collapsed != null && col.Collapsed)
            xlColumns.CollapseOnly();

        if (col.OutlineLevel != null)
        {
            var outlineLevel = col.OutlineLevel;
            xlColumns.ForEach(c => c.OutlineLevel = outlineLevel);
        }

        var styleIndex = col.Style != null ? int.Parse(col.Style.InnerText!) : -1;
        if (styleIndex >= 0)
        {
            StyleDecoder.ApplyStyle(xlColumns, styleIndex, styles);
        }
        else
        {
            xlColumns.Style = ws.Style;
        }
    }
}
