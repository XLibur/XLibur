using System;
using System.Globalization;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Rows;
using XLibur.Extensions;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// The box Excel draws a two-cell anchor in: from its <c>from</c> marker to its <c>to</c> marker, in
/// pixels of the workbook's DPI, measured from the top-left corner of the sheet (issue #782).
/// </summary>
/// <remarks>
/// <para>
/// Excel places and sizes every <c>xdr:twoCellAnchor</c> from its markers, whatever <c>editAs</c>
/// says, and ignores the picture's <c>a:xfrm</c>. It lays the grid out in whole pixels: a column is
/// its width in pixels for the workbook's maximum digit width, a row is its height in points
/// converted to pixels and truncated (25 pt is 41 px at 120 DPI, not 42), and a hidden row or column
/// is 0 px. A marker's offset ends at the edge of its cell, so an offset wider than the cell, or any
/// offset in a hidden one, adds no more than the cell. All of this was measured in Excel over COM.
/// </para>
/// <para>
/// A row or column the sheet does not hold is at the sheet's default height or width. When a file
/// states neither (no <c>customHeight</c>, no <c>defaultColWidth</c> or <c>baseColWidth</c>), Excel
/// derives the default from the font at its own DPI and XLibur uses its stored default instead. The
/// two agree at 96 DPI.
/// </para>
/// </remarks>
internal readonly record struct AnchorBox(int Left, int Top, int Width, int Height)
{
    private const double EmuPerInch = 914400;

    /// <summary>Measures the box between <paramref name="from"/> and <paramref name="to"/>.</summary>
    /// <remarks>
    /// A size can come out 0 or less: when every row or column of the box is hidden, or when
    /// <paramref name="to"/> lies before <paramref name="from"/>. The caller decides what to do then.
    /// </remarks>
    internal static AnchorBox Measure(XLWorksheet ws, Xdr.MarkerType from, Xdr.MarkerType to)
    {
        var grid = new Grid(ws);
        var (left, top) = grid.PointOf(from);
        var (right, bottom) = grid.PointOf(to);

        return new AnchorBox(
            ToPixels(left, grid.DpiX),
            ToPixels(top, grid.DpiY),
            ToPixels(right - left, grid.DpiX),
            ToPixels(bottom - top, grid.DpiY));
    }

    private static int ToPixels(double emu, double dpi) =>
        DrawingPartReader.ConvertFromEnglishMetricUnits((long)Math.Round(emu), dpi);

    /// <summary>The pixel width of each column and height of each row of one sheet.</summary>
    private sealed class Grid
    {
        private readonly XLWorksheet _ws;
        private readonly int _mdw;
        private readonly int _defaultColumnPx;
        private readonly int _defaultRowPx;

        internal Grid(XLWorksheet ws)
        {
            _ws = ws;
            DpiX = ws.Workbook.DpiX;
            DpiY = ws.Workbook.DpiY;

            // Column widths are in characters of the workbook font, the same as XLColumn.AdjustToContents.
            _mdw = XLHelper.GetMdw(ws.Workbook.Style.Font, ws.Workbook);
            _defaultColumnPx = XLHelper.NoCToPixels(ws.ColumnWidth, _mdw).RoundToInt();
            _defaultRowPx = PointsToPixels(ws.RowHeight);
        }

        internal double DpiX { get; }

        internal double DpiY { get; }

        /// <summary>Where <paramref name="marker"/> is, in EMU from the sheet's top-left corner.</summary>
        internal (double X, double Y) PointOf(Xdr.MarkerType marker)
        {
            var column = Math.Clamp(ParseInt(marker.ColumnId?.Text) + 1, 1, XLHelper.MaxColumnNumber);
            var row = Math.Clamp(ParseInt(marker.RowId?.Text) + 1, 1, XLHelper.MaxRowNumber);

            var emuPerPxX = EmuPerInch / DpiX;
            var emuPerPxY = EmuPerInch / DpiY;
            var x = SumPx(_ws.Internals.ColumnsCollection, column, _defaultColumnPx, ColumnPx) * emuPerPxX +
                    Math.Clamp(ParseLong(marker.ColumnOffset?.Text), 0, ColumnPx(column) * emuPerPxX);
            var y = SumPx(_ws.Internals.RowsCollection, row, _defaultRowPx, RowPx) * emuPerPxY +
                    Math.Clamp(ParseLong(marker.RowOffset?.Text), 0, RowPx(row) * emuPerPxY);
            return (x, y);
        }

        private int ColumnPx(int number) =>
            _ws.Internals.ColumnsCollection.TryGetValue(number, out var column) ? ColumnPx(column) : _defaultColumnPx;

        private int RowPx(int number) =>
            _ws.Internals.RowsCollection.TryGetValue(number, out var row) ? RowPx(row) : _defaultRowPx;

        /// <summary>
        /// A column the file gave its own width holds that width less
        /// <see cref="XLConstants.ColumnWidthOffset"/> (<see cref="WorksheetColumnReader"/>), so the
        /// file's width comes back exactly and converts as Excel converts it. A column at the sheet's
        /// default holds the default, which the load converted to characters through the digit width
        /// (<see cref="WorksheetElementReader"/>).
        /// </summary>
        private int ColumnPx(XLColumn column)
        {
            if (column.IsHidden)
                return 0;

            if (Math.Abs(column.Width - _ws.ColumnWidth) < XLHelper.Epsilon)
                return _defaultColumnPx;

            return XLHelper.WidthToPixels(column.Width + XLConstants.ColumnWidthOffset, _mdw);
        }

        private int RowPx(XLRow row) => row.IsHidden ? 0 : PointsToPixels(row.Height);

        /// <summary>Excel draws a row in whole pixels and drops the fraction.</summary>
        private int PointsToPixels(double points) =>
            (int)Math.Floor(XLHelper.PointsToPixels(points, DpiY) + 1e-6);

        /// <summary>
        /// The sum over lines 1 up to, not including, <paramref name="end"/>: the default for each,
        /// corrected for each line the sheet holds. Walks whichever is shorter, the range or the held
        /// lines, so a picture far down a long sheet does not cost a walk of every row.
        /// </summary>
        private static long SumPx<TLine, TAxis>(XLLineCollection<TLine, TAxis> lines, int end, int defaultPx,
            Func<TLine, int> px)
            where TLine : XLRangeBase
            where TAxis : struct, IGridAxis
        {
            var count = end - 1;
            var total = (long)count * defaultPx;
            if (count <= lines.Count)
            {
                for (var number = 1; number < end; number++)
                {
                    if (lines.TryGetValue(number, out var line))
                        total += px(line) - defaultPx;
                }
            }
            else
            {
                foreach (var (number, line) in lines)
                {
                    if (number < end)
                        total += px(line) - defaultPx;
                }
            }

            return total;
        }

        private static int ParseInt(string? text) =>
            int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : 0;

        private static long ParseLong(string? text) =>
            long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : 0;
    }
}
