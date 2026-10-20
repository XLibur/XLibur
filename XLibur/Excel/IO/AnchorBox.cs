using System;
using System.Collections.Generic;
using System.Globalization;
using XLibur.Excel.Rows;
using XLibur.Extensions;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// The box Excel draws a two-cell anchor in: from its <c>from</c> marker to its <c>to</c> marker, in
/// pixels of the workbook's DPI, measured from the top-left corner of the sheet (issue #782).
/// </summary>
/// <remarks>
/// Excel places and sizes every <c>xdr:twoCellAnchor</c> from its markers, whatever <c>editAs</c>
/// says, and ignores the picture's <c>a:xfrm</c>. <see cref="AnchorGrid"/> lays the grid out as
/// Excel does. A marker's offset ends at the edge of its cell, so an offset wider than the cell, or
/// any offset in a hidden one, adds no more than the cell. All of this was measured in Excel over COM.
/// </remarks>
internal readonly record struct AnchorBox(int Left, int Top, int Width, int Height)
{
    private const double EmuPerInch = 914400;

    /// <summary>Measures the box between <paramref name="from"/> and <paramref name="to"/>.</summary>
    /// <remarks>
    /// A size can come out 0 or less: when every row or column of the box is hidden, or when
    /// <paramref name="to"/> lies before <paramref name="from"/>. The caller decides what to do then.
    /// </remarks>
    internal static AnchorBox Measure(AnchorGrid grid, Xdr.MarkerType from, Xdr.MarkerType to)
    {
        var (left, top) = PointOf(grid, from);
        var (right, bottom) = PointOf(grid, to);

        return new AnchorBox(
            ToPixels(left, grid.DpiX),
            ToPixels(top, grid.DpiY),
            ToPixels(right - left, grid.DpiX),
            ToPixels(bottom - top, grid.DpiY));
    }

    /// <summary>Where <paramref name="marker"/> is, in EMU from the sheet's top-left corner.</summary>
    private static (double X, double Y) PointOf(AnchorGrid grid, Xdr.MarkerType marker)
    {
        var column = Math.Clamp(ParseInt(marker.ColumnId?.Text) + 1, 1, XLHelper.MaxColumnNumber);
        var row = Math.Clamp(ParseInt(marker.RowId?.Text) + 1, 1, XLHelper.MaxRowNumber);

        var emuPerPxX = EmuPerInch / grid.DpiX;
        var emuPerPxY = EmuPerInch / grid.DpiY;
        var x = grid.Columns.Start(column) * emuPerPxX +
                Math.Clamp(ParseLong(marker.ColumnOffset?.Text), 0, grid.Columns.Size(column) * emuPerPxX);
        var y = grid.Rows.Start(row) * emuPerPxY +
                Math.Clamp(ParseLong(marker.RowOffset?.Text), 0, grid.Rows.Size(row) * emuPerPxY);
        return (x, y);
    }

    private static int ToPixels(double emu, double dpi) =>
        DrawingPartReader.ConvertFromEnglishMetricUnits((long)Math.Round(emu), dpi);

    private static int ParseInt(string? text) =>
        int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : 0;

    private static long ParseLong(string? text) =>
        long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : 0;
}

/// <summary>
/// The pixel width of each column and height of each row of one sheet, as Excel lays them out.
/// Built once for a sheet and shared by every anchor on it.
/// </summary>
/// <remarks>
/// <para>
/// Excel lays the grid out in whole pixels. A column is its width in pixels for the workbook's
/// maximum digit width, which Excel takes from the first font of the stylesheet, whichever font the
/// default cell format and the Normal style use. A row is its height in points converted to pixels
/// and truncated: 25 pt is 41 px at 120 DPI, not 42. A hidden row or column, and a column of width
/// 0, is 0 px.
/// </para>
/// <para>
/// The columns come from the widths the file states (<see cref="StatedColumnWidths"/>), not the
/// model, which holds them in two units it cannot tell apart. A <c>&lt;col&gt;</c> range that runs
/// to the last column is a run of columns of its own, and the columns before it are at the sheet's
/// default. That default is <c>defaultColWidth</c> converted like a column width; failing that,
/// <c>baseColWidth</c> digits and their padding, rounded up to a multiple of 8 px. All of this was
/// measured in Excel over COM.
/// </para>
/// <para>
/// The rows come from the model, which holds every row height in points. A row the sheet does not
/// hold is at the sheet's default height. When the file states no default (no
/// <c>customHeight</c>), Excel derives it from the font at its own DPI, and XLibur uses its stored
/// default instead. The two agree at 96 DPI.
/// </para>
/// </remarks>
internal sealed class AnchorGrid
{
    internal AnchorGrid(XLWorksheet ws, StatedColumnWidths stated, int maxDigitWidth)
    {
        DpiX = ws.Workbook.DpiX;
        DpiY = ws.Workbook.DpiY;
        Columns = ColumnAxis(stated, maxDigitWidth);
        Rows = RowAxis(ws, DpiY);
    }

    internal double DpiX { get; }

    internal double DpiY { get; }

    internal PixelAxis Columns { get; }

    internal PixelAxis Rows { get; }

    private static PixelAxis ColumnAxis(StatedColumnWidths stated, int mdw)
    {
        var defaultPx = stated.DefaultColumnWidth is { } defaultWidth
            ? WidthToPixels(defaultWidth, mdw)
            : RoundUpToMultipleOf8(XLHelper.NoCToPixels(stated.BaseColumnWidth, mdw).RoundToInt());

        var runs = new List<(int First, int Last, int Px)>(stated.Columns.Count);
        foreach (var (min, max, width, hidden) in stated.Columns)
        {
            var px = hidden ? 0 : width is { } w ? WidthToPixels(w, mdw) : defaultPx;
            runs.Add((min, max, px));
        }

        return new PixelAxis(defaultPx, XLHelper.MaxColumnNumber, runs);
    }

    private static PixelAxis RowAxis(XLWorksheet ws, double dpiY)
    {
        var defaultPx = PointsToPixels(ws.RowHeight, dpiY);
        var runs = new List<(int First, int Last, int Px)>();
        foreach (var (number, row) in ws.Internals.RowsCollection)
        {
            var px = RowPx(row, dpiY);
            if (px != defaultPx)
                runs.Add((number, number, px));
        }

        return new PixelAxis(defaultPx, XLHelper.MaxRowNumber, runs);
    }

    private static int RowPx(XLRow row, double dpiY) => row.IsHidden ? 0 : PointsToPixels(row.Height, dpiY);

    /// <summary>Excel draws a row in whole pixels and drops the fraction.</summary>
    private static int PointsToPixels(double points, double dpi) =>
        Math.Max(0, (int)Math.Floor(XLHelper.PointsToPixels(points, dpi) + 1e-6));

    private static int WidthToPixels(double width, int mdw) => Math.Max(0, XLHelper.WidthToPixels(width, mdw));

    private static int RoundUpToMultipleOf8(int px) => (px + 7) / 8 * 8;
}

/// <summary>
/// The pixel size of every line of one axis of a sheet: a default, and the runs of lines that
/// differ from it. Finds where a line starts by a binary search over the runs, so a sheet with many
/// sized rows costs one pass to build, not one for each marker.
/// </summary>
internal sealed class PixelAxis
{
    private readonly int _defaultPx;
    private readonly int[] _first;
    private readonly int[] _last;
    private readonly int[] _px;

    /// <summary>The pixels the runs before run <c>i</c> add to their lines' default sizes.</summary>
    private readonly long[] _extraBefore;

    /// <param name="defaultPx">The size of a line no run covers.</param>
    /// <param name="maxLine">The last line of the axis; runs are clipped to it.</param>
    /// <param name="runs">
    /// Inclusive runs of lines of one size, in any order. Where two overlap, the one that starts
    /// first keeps the lines they share.
    /// </param>
    internal PixelAxis(int defaultPx, int maxLine, List<(int First, int Last, int Px)> runs)
    {
        _defaultPx = defaultPx;
        runs.Sort(static (a, b) => a.First.CompareTo(b.First));

        var first = new List<int>(runs.Count);
        var last = new List<int>(runs.Count);
        var px = new List<int>(runs.Count);
        var previousLast = 0;
        foreach (var run in runs)
        {
            var start = Math.Max(Math.Max(run.First, 1), previousLast + 1);
            var end = Math.Min(run.Last, maxLine);
            if (start > end)
                continue;

            first.Add(start);
            last.Add(end);
            px.Add(run.Px);
            previousLast = end;
        }

        _first = first.ToArray();
        _last = last.ToArray();
        _px = px.ToArray();
        _extraBefore = new long[_first.Length + 1];
        for (var i = 0; i < _first.Length; i++)
            _extraBefore[i + 1] = _extraBefore[i] + (long)(_last[i] - _first[i] + 1) * (_px[i] - _defaultPx);
    }

    /// <summary>The pixels from the start of line 1 to the start of <paramref name="line"/>.</summary>
    internal long Start(int line)
    {
        var total = (long)(line - 1) * _defaultPx;

        // The last run that starts before the line: every run before it lies wholly before the line.
        var i = RunStartingAtOrBefore(line - 1);
        if (i < 0)
            return total;

        var covered = Math.Min(_last[i], line - 1) - _first[i] + 1;
        return total + _extraBefore[i] + (long)covered * (_px[i] - _defaultPx);
    }

    /// <summary>The size of <paramref name="line"/> in pixels.</summary>
    internal int Size(int line)
    {
        var i = RunStartingAtOrBefore(line);
        return i >= 0 && line <= _last[i] ? _px[i] : _defaultPx;
    }

    /// <summary>The index of the last run whose first line is at or before <paramref name="line"/>, or -1.</summary>
    private int RunStartingAtOrBefore(int line)
    {
        var index = Array.BinarySearch(_first, line);
        return index >= 0 ? index : ~index - 1;
    }
}
