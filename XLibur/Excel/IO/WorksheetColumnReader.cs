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
    /// A range that ends at the last column supplies the sheet's default width and style, and is
    /// then normally skipped — expanding it would materialise an <c>XLColumn</c> for all
    /// <see cref="XLHelper.MaxColumnNumber"/> columns on every load, which is why the skip exists
    /// and why it stays for the shape that motivated it. <c>ColumnWriter.WritePostColumns</c>
    /// writes exactly that shape: style, width and <c>customWidth</c>, never the three per-column
    /// flags.
    /// <para>
    /// It is <see cref="StatesPerColumnFlags"/> that decides, not the <c>max</c> attribute alone.
    /// A range ending at the last column may also carry <c>hidden</c>, <c>collapsed</c> or
    /// <c>outlineLevel</c> — Excel writes exactly that when a user hides or groups from some
    /// column rightwards — and those cannot be expressed as a sheet default, so such a range is
    /// loaded like any other. Before this was fixed (D14) the skip was unconditional and all three
    /// were silently dropped: <c>&lt;col min="2" max="16384" hidden="1" outlineLevel="1"/&gt;</c>
    /// loaded with the width applied and the column neither hidden nor grouped.
    /// </para>
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
            if (col.Max?.Value == XLHelper.MaxColumnNumber && !StatesPerColumnFlags(col))
                continue;

            LoadColumn(col, ws, styles);
        }
    }

    /// <summary>
    /// Whether a <c>&lt;col&gt;</c> carries per-column state that a sheet default cannot express.
    /// Width and style can be defaulted for the whole sheet; being hidden, collapsed or in an
    /// outline group cannot, so a range stating any of the three has to be expanded onto the
    /// columns it covers even when it runs to the last one.
    /// </summary>
    private static bool StatesPerColumnFlags(Column col)
    {
        return (col.Hidden is not null && col.Hidden)
               || (col.Collapsed is not null && col.Collapsed)
               || col.OutlineLevel is not null;
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
