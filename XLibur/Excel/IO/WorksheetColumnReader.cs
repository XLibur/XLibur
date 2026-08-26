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
    /// <summary>
    /// Reads a <c>&lt;cols&gt;</c> element: the sheet's default column width and style come from
    /// the range that ends at the last column, and every other range is loaded individually.
    /// </summary>
    /// <remarks>
    /// <b>Known defect (D14), pre-existing and moved here unchanged by spec 28.</b> Every range
    /// whose <c>max</c> is the last column is treated as the sheet default and skipped, so a file
    /// stating <c>&lt;col min="2" max="16384" hidden="1" outlineLevel="1"/&gt;</c> loses its
    /// <c>hidden</c>, <c>collapsed</c> and <c>outlineLevel</c> — only the width and style survive.
    /// XLibur's own writer never emits those three on its trailing default range
    /// (<c>ColumnWriter.WritePostColumns</c> writes style, width and <c>customWidth</c> only), so
    /// its own files round-trip; a foreign one may not. Fixing it needs a discriminator that tells
    /// the writer's default tail from a genuine explicit range, which is not this module's move.
    /// </remarks>
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

    /// <summary>
    /// Reads one <c>&lt;col&gt;</c> onto the columns it spans: width, visibility, collapsed state,
    /// outline level and style. A range stating no style inherits the worksheet's.
    /// </summary>
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
