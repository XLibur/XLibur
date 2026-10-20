using System;
using System.Collections.Generic;
using System.Threading;

namespace XLibur.Excel.Drawings;

/// <summary>
/// The size of every row and column of one sheet at one moment, for kept picture anchors to compare
/// against (<see cref="XLKeptAnchor"/>). Taken once for a sheet and shared by every anchor on it: at
/// load, and again at save (issue #785).
/// </summary>
internal sealed class GridLayout
{
    private GridLayout(LineLayout rows, LineLayout columns)
    {
        Rows = rows;
        Columns = columns;
    }

    internal LineLayout Rows { get; }

    internal LineLayout Columns { get; }

    /// <summary>The layout <paramref name="worksheet"/> has now. A hidden line has size 0.</summary>
    internal static GridLayout Of(XLWorksheet worksheet) => new(
        LineLayout.Of(worksheet.Internals.RowsCollection, worksheet.RowHeight,
            static row => row.IsHidden ? 0 : row.Height),
        LineLayout.Of(worksheet.Internals.ColumnsCollection, worksheet.ColumnWidth,
            static column => column.IsHidden ? 0 : column.Width));

    /// <summary>
    /// The layout <paramref name="worksheet"/> has when it is first asked for: one for every kept
    /// anchor on the sheet, taken only if the sheet has one (#785).
    /// </summary>
    internal static Lazy<GridLayout> OnFirstUse(XLWorksheet worksheet) =>
        new(() => Of(worksheet), LazyThreadSafetyMode.None);
}

/// <summary>The rows or the columns <see cref="First"/> to <see cref="Last"/>, both included.</summary>
internal readonly record struct LineRun(int First, int Last);

/// <summary>
/// The sizes of the rows or the columns of a sheet: the sheet's default, and each line that differs
/// from it, in line order. Lines the sheet has never materialised take the default, so the layout is
/// described by the few lines that differ rather than by every line.
/// </summary>
/// <remarks>
/// A run of lines is found by a binary search, so comparing one anchor's lines costs the lines in
/// its run that differ from the default, not every line on the sheet.
/// </remarks>
internal sealed class LineLayout
{
    private readonly double _defaultSize;

    /// <summary>The number of each line whose size differs from the default, ascending.</summary>
    private readonly int[] _lines;

    /// <summary>The size of the line at the same index in <see cref="_lines"/>.</summary>
    private readonly double[] _sizes;

    /// <summary>The layout <see cref="_sharedPrefix"/> was counted against.</summary>
    private LineLayout? _sharedPrefixWith;

    private int _sharedPrefix;

    private LineLayout(double defaultSize, int[] lines, double[] sizes)
    {
        _defaultSize = defaultSize;
        _lines = lines;
        _sizes = sizes;
    }

    internal static LineLayout Of<TLine>(IDictionary<int, TLine> lines, double defaultSize, Func<TLine, double> sizeOf)
    {
        var numbers = new List<int>();
        var sizes = new List<double>();
        foreach (var (number, line) in lines)
        {
            var size = sizeOf(line);
            if (SameSize(size, defaultSize))
                continue;

            numbers.Add(number);
            sizes.Add(size);
        }

        var numberArray = numbers.ToArray();
        var sizeArray = sizes.ToArray();
        Array.Sort(numberArray, sizeArray);
        return new LineLayout(defaultSize, numberArray, sizeArray);
    }

    /// <summary>
    /// Whether the lines of <paramref name="run"/> in this layout have the sizes the lines of
    /// <paramref name="beforeRun"/> had in <paramref name="before"/>, line for line from the first.
    /// Lines past the end of the shorter run must be of the default size.
    /// </summary>
    internal bool HasSameSizes(LineRun run, LineLayout before, LineRun beforeRun)
    {
        if (!SameSize(_defaultSize, before._defaultSize))
            return false;

        var first = run.First;
        var beforeFirst = beforeRun.First;
        var index = FirstAtOrAfter(first);
        var end = FirstAtOrAfter(run.Last + 1);
        var beforeIndex = before.FirstAtOrAfter(beforeFirst);
        var beforeEnd = before.FirstAtOrAfter(beforeRun.Last + 1);
        if (end - index != beforeEnd - beforeIndex)
            return false;

        // Where both runs start at the same line, the lines the two layouts share from the top of
        // the sheet need no comparing one by one. An absolute picture's run starts at the top, so
        // without this every such picture would compare every sized line above it.
        if (first == beforeFirst && index == beforeIndex)
        {
            index = Math.Max(index, Math.Min(SharedPrefixWith(before), end));
            beforeIndex = index;
        }

        for (; index < end; index++, beforeIndex++)
        {
            if (_lines[index] - first != before._lines[beforeIndex] - beforeFirst ||
                !SameSize(_sizes[index], before._sizes[beforeIndex]))
                return false;
        }

        return true;
    }

    /// <summary>The index of the first sized line at or after <paramref name="line"/>.</summary>
    private int FirstAtOrAfter(int line)
    {
        var index = Array.BinarySearch(_lines, line);
        return index >= 0 ? index : ~index;
    }

    /// <summary>
    /// How many sized lines, from the first, this layout and <paramref name="other"/> have alike, by
    /// number and size. Counted once for the pair, and kept for the next anchor on the sheet.
    /// </summary>
    private int SharedPrefixWith(LineLayout other)
    {
        if (ReferenceEquals(_sharedPrefixWith, other))
            return _sharedPrefix;

        var count = Math.Min(_lines.Length, other._lines.Length);
        var shared = 0;
        while (shared < count && _lines[shared] == other._lines[shared] &&
               SameSize(_sizes[shared], other._sizes[shared]))
            shared++;

        _sharedPrefixWith = other;
        _sharedPrefix = shared;
        return shared;
    }

    private static bool SameSize(double left, double right) => Math.Abs(left - right) < XLHelper.Epsilon;
}
