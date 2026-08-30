using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.CalcEngine.Functions;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// Modern dynamic-array worksheet functions (SEQUENCE, UNIQUE, SORT, SORTBY, FILTER, XLOOKUP,
/// XMATCH). They are registered as <see cref="FunctionFlags.ReturnsArray"/> so the array-formula
/// engine (see <c>FunctionDefinition.CallAsArray</c>) uses their whole array output. In a plain
/// (non-array) formula the result collapses to its top-left element, matching Excel's implicit-
/// intersection behaviour for pre-spill hosts; entered as an array formula the full result is
/// written across the range. Actual grid <em>spilling</em> is a separate, later piece of work.
/// </summary>
internal static class DynamicArray
{
    /// <summary>
    /// What Excel reports as <c>#CALC!</c> when a dynamic-array function would produce nothing at
    /// all — every row dropped, every value ignored. XLibur's value model has no <c>#CALC!</c>, so
    /// these cases report <c>#VALUE!</c> instead; the argument was of the right shape, it just left
    /// no result behind.
    /// </summary>
    private const XLError EmptyResult = XLError.IncompatibleValue;

    private const FunctionFlags Spilling = FunctionFlags.Range | FunctionFlags.ReturnsArray;

    public static void Register(FunctionRegistry ce)
    {
        ce.RegisterFunction("CHOOSECOLS", 2, 255, ChooseCols, Spilling, AllowRange.All); // Returns the specified columns from an array
        ce.RegisterFunction("CHOOSEROWS", 2, 255, ChooseRows, Spilling, AllowRange.All); // Returns the specified rows from an array
        ce.RegisterFunction("DROP", 2, 3, Drop, Spilling, AllowRange.All); // Drops rows or columns from the start or end of an array
        ce.RegisterFunction("EXPAND", 2, 4, Expand, Spilling, AllowRange.All); // Expands an array to the given dimensions
        ce.RegisterFunction("FILTER", 2, 3, Filter, Spilling, AllowRange.All); // Filters a range or array based on criteria
        ce.RegisterFunction("HSTACK", 1, 255, HStack, Spilling, AllowRange.All); // Appends arrays side by side
        ce.RegisterFunction("SEQUENCE", 1, 4, Sequence, Spilling, AllowRange.All); // Generates a list of sequential numbers
        ce.RegisterFunction("SORT", 1, 4, Sort, Spilling, AllowRange.All); // Sorts the contents of a range or array
        ce.RegisterFunction("SORTBY", 2, 255, SortBy, Spilling, AllowRange.All); // Sorts a range or array based on the values in a corresponding range or array
        ce.RegisterFunction("TAKE", 2, 3, Take, Spilling, AllowRange.All); // Takes rows or columns from the start or end of an array
        ce.RegisterFunction("TOCOL", 1, 3, ToCol, Spilling, AllowRange.All); // Returns the array as one column
        ce.RegisterFunction("TOROW", 1, 3, ToRow, Spilling, AllowRange.All); // Returns the array as one row
        ce.RegisterFunction("UNIQUE", 1, 3, Unique, Spilling, AllowRange.All); // Returns the distinct values from a range or array
        ce.RegisterFunction("VSTACK", 1, 255, VStack, Spilling, AllowRange.All); // Appends arrays one below another
        ce.RegisterFunction("WRAPCOLS", 2, 3, WrapCols, Spilling, AllowRange.All); // Wraps a vector into columns of a given length
        ce.RegisterFunction("WRAPROWS", 2, 3, WrapRows, Spilling, AllowRange.All); // Wraps a vector into rows of a given length
        ce.RegisterFunction("XLOOKUP", 3, 6, XLookup, Spilling, AllowRange.All); // Searches a range or array and returns the matching item(s)
        ce.RegisterFunction("XMATCH", 2, 4, XMatch, FunctionFlags.Range, AllowRange.All); // Returns the relative position of an item in a range or array
    }

    #region Stacking

    private static AnyValue VStack(CalcContext ctx, Span<AnyValue> args)
        => Stack(ctx, args, vertically: true);

    private static AnyValue HStack(CalcContext ctx, Span<AnyValue> args)
        => Stack(ctx, args, vertically: false);

    /// <summary>
    /// Append the arguments along one axis. The other axis grows to the widest (or tallest)
    /// argument, and the arguments that fall short are padded with <c>#N/A</c>.
    /// </summary>
#pragma warning disable S3776 // Measure the arguments, then fill; both walks are flat and sequential
    private static AnyValue Stack(CalcContext ctx, Span<AnyValue> args, bool vertically)
    {
        var parts = new List<Array>(args.Length);
        foreach (var arg in args)
        {
            if (!arg.TryPickCollectionArray(out var array, ctx))
                return XLError.IncompatibleValue;

            parts.Add(vertically ? array! : new TransposedArray(array!));
        }

        var width = 0;
        var height = 0;
        foreach (var part in parts)
        {
            width = Math.Max(width, part.Width);
            height += part.Height;
        }

        var data = new ScalarValue[height, width];
        var row = 0;
        foreach (var part in parts)
        {
            for (var y = 0; y < part.Height; y++, row++)
            {
                for (var x = 0; x < width; x++)
                    data[row, x] = x < part.Width ? part[y, x] : XLError.NoValueAvailable;
            }
        }

        return Orient(new ConstArray(data), !vertically);
    }
#pragma warning restore S3776

    #endregion

    #region Flattening and wrapping

    private static AnyValue ToRow(CalcContext ctx, Span<AnyValue> args)
        => Flatten(ctx, args, intoRow: true);

    private static AnyValue ToCol(CalcContext ctx, Span<AnyValue> args)
        => Flatten(ctx, args, intoRow: false);

    /// <summary>
    /// TOROW/TOCOL(array, [ignore], [scan_by_column]) — read every value in scan order, optionally
    /// skipping blanks (1), errors (2) or both (3), and lay the result out along a single axis.
    /// </summary>
#pragma warning disable S3776 // Optional-argument guards ahead of one filtered walk; already reduced from 21
    private static AnyValue Flatten(CalcContext ctx, Span<AnyValue> args, bool intoRow)
    {
        if (!args[0].TryPickCollectionArray(out var array, ctx))
            return XLError.IncompatibleValue;

        var ignore = 0;
        if (args.Length > 1 && !TryIntArg(ctx, args[1], out ignore, out var ignoreError))
            return ignoreError;
        if (ignore is < 0 or > 3)
            return XLError.IncompatibleValue;

        var byColumn = false;
        if (args.Length > 2 && !TryBoolArg(ctx, args[2], out byColumn, out var byColumnError))
            return byColumnError;

        var skipBlanks = (ignore & 1) != 0;
        var skipErrors = (ignore & 2) != 0;

        // Scanning by column is the same walk over the transpose.
        var source = byColumn ? new TransposedArray(array!) : array!;
        var kept = new List<ScalarValue>(source.Height * source.Width);
        foreach (var value in source)
        {
            if (skipBlanks && value.IsBlank)
                continue;
            if (skipErrors && value.IsError)
                continue;

            kept.Add(value);
        }

        if (kept.Count == 0)
            return EmptyResult;

        // Build the column and let TOROW read it sideways, rather than branching per value.
        var data = new ScalarValue[kept.Count, 1];
        for (var i = 0; i < kept.Count; i++)
            data[i, 0] = kept[i];

        return Orient(new ConstArray(data), intoRow);
    }
#pragma warning restore S3776

    private static AnyValue WrapRows(CalcContext ctx, Span<AnyValue> args)
        => Wrap(ctx, args, intoRows: true);

    private static AnyValue WrapCols(CalcContext ctx, Span<AnyValue> args)
        => Wrap(ctx, args, intoRows: false);

    /// <summary>
    /// WRAPROWS/WRAPCOLS(vector, wrap_count, [pad_with]) — cut a one-dimensional vector into pieces
    /// of <c>wrap_count</c> values. Only the last piece can be short, and it is padded.
    /// </summary>
    private static AnyValue Wrap(CalcContext ctx, Span<AnyValue> args, bool intoRows)
    {
        if (!args[0].TryPickCollectionArray(out var array, ctx))
            return XLError.IncompatibleValue;

        // The input has to be a vector; a rectangle has no unambiguous reading order to wrap.
        if (array!.Width != 1 && array.Height != 1)
            return XLError.IncompatibleValue;

        if (!TryIntArg(ctx, args[1], out var wrapCount, out var wrapCountError))
            return wrapCountError;
        if (wrapCount < 1)
            return XLError.NumberInvalid;

        var padding = args.Length > 2 ? ScalarOf(ctx, args[2]) : XLError.NoValueAvailable;

        var values = new List<ScalarValue>(array.Height * array.Width);
        foreach (var value in array)
            values.Add(value);

        // Lay the pieces out as rows and let WRAPCOLS read them as columns, rather than branching
        // on the orientation for every value.
        var pieces = (values.Count + wrapCount - 1) / wrapCount;
        var data = new ScalarValue[pieces, wrapCount];
        for (var piece = 0; piece < pieces; piece++)
        {
            for (var offset = 0; offset < wrapCount; offset++)
            {
                var index = piece * wrapCount + offset;
                data[piece, offset] = index < values.Count ? values[index] : padding;
            }
        }

        return Orient(new ConstArray(data), !intoRows);
    }

    #endregion

    #region Selecting rows and columns

    private static AnyValue ChooseRows(CalcContext ctx, Span<AnyValue> args)
        => Choose(ctx, args, byRow: true);

    private static AnyValue ChooseCols(CalcContext ctx, Span<AnyValue> args)
        => Choose(ctx, args, byRow: false);

    /// <summary>
    /// CHOOSEROWS/CHOOSECOLS(array, num1, …) — pick lines out of the array in the order asked for,
    /// repeats included. A negative index counts back from the end.
    /// </summary>
#pragma warning disable S3776 // One index-validation rule applied over two argument shapes
    private static AnyValue Choose(CalcContext ctx, Span<AnyValue> args, bool byRow)
    {
        if (!args[0].TryPickCollectionArray(out var array, ctx))
            return XLError.IncompatibleValue;

        var source = byRow ? array! : new TransposedArray(array!);
        var count = source.Height;

        var selected = new List<int>();
        for (var i = 1; i < args.Length; i++)
        {
            // An argument is usually a single index, but Excel also accepts a whole array of them,
            // as in CHOOSEROWS(A1:C5, {1,3}).
            IEnumerable<ScalarValue> indices;
            if (args[i].TryPickScalar(out var scalar, out _))
                indices = [scalar];
            else if (args[i].TryPickCollectionArray(out var indexArray, ctx))
                indices = indexArray!;
            else
                return XLError.IncompatibleValue;

            foreach (var index in indices)
            {
                if (!index.ToNumber(ctx.Culture).TryPickT0(out var number, out var indexError))
                    return indexError;

                var line = (int)Math.Truncate(number);

                // A negative index counts back from the end: -1 is the last line.
                if (line < 0)
                    line = count + line + 1;

                if (line < 1 || line > count)
                    return XLError.IncompatibleValue;

                selected.Add(line - 1);
            }
        }

        if (selected.Count == 0)
            return EmptyResult;

        return Orient(BuildRows(source, selected), !byRow);
    }
#pragma warning restore S3776

    #endregion

    #region Slicing and padding

    private static AnyValue Take(CalcContext ctx, Span<AnyValue> args)
        => Slice(ctx, args, dropping: false);

    private static AnyValue Drop(CalcContext ctx, Span<AnyValue> args)
        => Slice(ctx, args, dropping: true);

    /// <summary>
    /// TAKE/DROP(array, rows, [columns]) — keep (or discard) that many lines from the start of each
    /// axis, or from the end when the count is negative. An omitted or blank count leaves the axis
    /// alone.
    /// </summary>
    private static AnyValue Slice(CalcContext ctx, Span<AnyValue> args, bool dropping)
    {
        if (!args[0].TryPickCollectionArray(out var array, ctx))
            return XLError.IncompatibleValue;

        if (!TryOptionalIntArg(ctx, args, 1, out var rows, out var rowsError))
            return rowsError;
        if (!TryOptionalIntArg(ctx, args, 2, out var columns, out var columnsError))
            return columnsError;

        if (!TrySliceAxis(array!.Height, rows, dropping, out var rowOffset, out var rowCount) ||
            !TrySliceAxis(array.Width, columns, dropping, out var columnOffset, out var columnCount))
        {
            return EmptyResult;
        }

        if (rowOffset == 0 && rowCount == array.Height && columnOffset == 0 && columnCount == array.Width)
            return array;

        return new SlicedArray(array, rowOffset, rowCount, columnOffset, columnCount);
    }

    /// <summary>Resolve one axis of TAKE/DROP into an offset and a length, or false if nothing is left.</summary>
    private static bool TrySliceAxis(int length, int? count, bool dropping, out int offset, out int result)
    {
        offset = 0;
        result = length;
        if (count is null)
            return true;

        var requested = count.Value;
        var magnitude = Math.Min(Math.Abs(requested), length);

        // DROP(n) keeps what TAKE(-(length - n)) would; both directions collapse to "how many
        // lines survive, counted from which end".
        var kept = dropping ? length - magnitude : magnitude;
        if (kept <= 0)
            return false;

        var fromEnd = requested < 0 ? !dropping : dropping;
        offset = fromEnd ? length - kept : 0;
        result = kept;
        return true;
    }

    /// <summary>
    /// EXPAND(array, rows, [columns], [pad_with]) — grow the array to the given size, filling the
    /// new cells. Shrinking is not expansion, so a smaller size is rejected.
    /// </summary>
    private static AnyValue Expand(CalcContext ctx, Span<AnyValue> args)
    {
        if (!args[0].TryPickCollectionArray(out var array, ctx))
            return XLError.IncompatibleValue;

        if (!TryOptionalIntArg(ctx, args, 1, out var rows, out var rowsError))
            return rowsError;
        if (!TryOptionalIntArg(ctx, args, 2, out var columns, out var columnsError))
            return columnsError;

        var height = rows ?? array!.Height;
        var width = columns ?? array!.Width;
        if (height < array!.Height || width < array.Width)
            return XLError.IncompatibleValue;
        if (height > XLHelper.MaxRowNumber || width > XLHelper.MaxColumnNumber)
            return XLError.NumberInvalid;

        var padding = args.Length > 3 ? ScalarOf(ctx, args[3]) : XLError.NoValueAvailable;

        var data = new ScalarValue[height, width];
        for (var y = 0; y < height; y++)
        {
            for (var x = 0; x < width; x++)
                data[y, x] = y < array.Height && x < array.Width ? array[y, x] : padding;
        }

        return new ConstArray(data);
    }

    #endregion

    /// <summary>
    /// Read an argument that may be left out entirely or written as an empty placeholder, as in
    /// <c>TAKE(A1:C3,,2)</c>. Both mean "leave this axis alone".
    /// </summary>
    private static bool TryOptionalIntArg(CalcContext ctx, Span<AnyValue> args, int index, out int? value, out XLError error)
    {
        value = null;
        error = default;
        if (args.Length <= index)
            return true;

        if (args[index].TryPickScalar(out var scalar, out _) && scalar.IsBlank)
            return true;

        if (!TryIntArg(ctx, args[index], out var number, out error))
            return false;

        value = number;
        return true;
    }

    private static ScalarValue ScalarOf(CalcContext ctx, in AnyValue value)
        => TryScalarArg(ctx, value, out var scalar) ? scalar : XLError.NoValueAvailable;

    private static AnyValue Sequence(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryIntArg(ctx, args[0], out var rows, out var rowsError))
            return rowsError;

        var columns = 1;
        if (args.Length > 1 && !TryIntArg(ctx, args[1], out columns, out var columnsError))
            return columnsError;

        double start = 1;
        if (args.Length > 2 && !TryNumberArg(ctx, args[2], out start, out var startError))
            return startError;

        double step = 1;
        if (args.Length > 3 && !TryNumberArg(ctx, args[3], out step, out var stepError))
            return stepError;

        if (rows < 1 || columns < 1 || rows > XLHelper.MaxRowNumber || columns > XLHelper.MaxColumnNumber)
            return XLError.NumberInvalid;

        var data = new ScalarValue[rows, columns];
        var current = start;
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < columns; c++)
            {
                data[r, c] = current;
                current += step;
            }
        }

        return new ConstArray(data);
    }

#pragma warning disable S3776 // Argument guards, then a linear group-by over rows; each stage is flat
    private static AnyValue Unique(CalcContext ctx, Span<AnyValue> args)
    {
        if (!args[0].TryPickCollectionArray(out var array, ctx))
            return XLError.IncompatibleValue;

        var byColumn = false;
        if (args.Length > 1 && !TryBoolArg(ctx, args[1], out byColumn, out var byColumnError))
            return byColumnError;

        var exactlyOnce = false;
        if (args.Length > 2 && !TryBoolArg(ctx, args[2], out exactlyOnce, out var exactlyOnceError))
            return exactlyOnceError;

        // Work row-wise; when comparing columns, operate on the transpose and transpose back.
        var source = byColumn ? new TransposedArray(array!) : array!;
        var width = source.Width;

        var representatives = new List<int>();
        var counts = new List<int>();
        for (var r = 0; r < source.Height; r++)
        {
            var matched = -1;
            for (var k = 0; k < representatives.Count; k++)
            {
                if (RowsEqual(source, r, representatives[k], width))
                {
                    matched = k;
                    break;
                }
            }

            if (matched == -1)
            {
                representatives.Add(r);
                counts.Add(1);
            }
            else
            {
                counts[matched]++;
            }
        }

        var kept = new List<int>();
        for (var k = 0; k < representatives.Count; k++)
        {
            if (!exactlyOnce || counts[k] == 1)
                kept.Add(representatives[k]);
        }

        if (kept.Count == 0)
            return XLError.NoValueAvailable;

        var result = new ScalarValue[kept.Count, width];
        for (var i = 0; i < kept.Count; i++)
        {
            for (var c = 0; c < width; c++)
                result[i, c] = source[kept[i], c];
        }

        return Orient(new ConstArray(result), byColumn);
    }
#pragma warning restore S3776

    private static AnyValue Sort(CalcContext ctx, Span<AnyValue> args)
    {
        if (!args[0].TryPickCollectionArray(out var array, ctx))
            return XLError.IncompatibleValue;

        var sortIndex = 1;
        if (args.Length > 1 && !TryIntArg(ctx, args[1], out sortIndex, out var sortIndexError))
            return sortIndexError;

        var sortOrder = 1;
        if (args.Length > 2 && !TryIntArg(ctx, args[2], out sortOrder, out var sortOrderError))
            return sortOrderError;

        var byColumn = false;
        if (args.Length > 3 && !TryBoolArg(ctx, args[3], out byColumn, out var byColumnError))
            return byColumnError;

        if (sortOrder != 1 && sortOrder != -1)
            return XLError.IncompatibleValue;

        var source = byColumn ? new TransposedArray(array!) : array!;
        var width = source.Width;
        if (sortIndex < 1 || sortIndex > width)
            return XLError.IncompatibleValue;

        var key = sortIndex - 1;
        var comparer = ScalarValueComparer.SortIgnoreCase;
        // OrderBy is a stable sort, so equal keys keep their original order (Excel behaviour).
        var order = Enumerable.Range(0, source.Height)
            .OrderBy(r => r, Comparer<int>.Create((a, b) => sortOrder * comparer.Compare(source[a, key], source[b, key])))
            .ToList();

        return Orient(BuildRows(source, order), byColumn);
    }

#pragma warning disable S3776 // Parsing the (by_array, [order]) pairs is one loop with one comparator
    private static AnyValue SortBy(CalcContext ctx, Span<AnyValue> args)
    {
        if (!args[0].TryPickCollectionArray(out var array, ctx))
            return XLError.IncompatibleValue;

        var height = array!.Height;

        // Parse (by_array, [order]) groups. A range/array argument starts a new key; a following
        // scalar argument is that key's sort order (1 ascending, -1 descending).
        var keys = new List<(Array By, int Order)>();
        var i = 1;
        while (i < args.Length)
        {
            if (!args[i].TryPickCollectionArray(out var by, ctx))
                return XLError.IncompatibleValue;
            if (by!.Width != 1 || by.Height != height)
                return XLError.IncompatibleValue;
            i++;

            var order = 1;
            // A following argument is the order for this key when it does NOT itself validate as a
            // by_array (the (Height x 1) shape by_array needs) — rather than when it fails
            // IsScalarType, which is what let a single-cell reference be mistaken for the start of a
            // new by_array instead of being read as the order. Trying the by_array interpretation
            // first, rather than checking for a 1x1 shape, is what keeps a genuine one-row sort with
            // two single-cell by_arrays (SORTBY(A1,B1,C1)) working: there height is 1, so a 1x1
            // reference is ambiguous by shape alone, and Excel prefers the by_array reading.
            if (i < args.Length && !IsValidByArray(args[i], ctx, height))
            {
                if (!TryIntArg(ctx, args[i], out order, out var orderError))
                    return orderError;
                if (order != 1 && order != -1)
                    return XLError.IncompatibleValue;
                i++;
            }

            keys.Add((by, order));
        }

        var comparer = ScalarValueComparer.SortIgnoreCase;
        var indices = Enumerable.Range(0, height)
            .OrderBy(r => r, Comparer<int>.Create((a, b) =>
            {
                foreach (var (by, order) in keys)
                {
                    var cmp = order * comparer.Compare(by[a, 0], by[b, 0]);
                    if (cmp != 0)
                        return cmp;
                }

                return 0;
            }))
            .ToList();

        return BuildRows(array, indices);
    }
#pragma warning restore S3776

    /// <summary>Does <paramref name="value"/> have the (Height x 1) shape a SORTBY by_array needs?</summary>
    private static bool IsValidByArray(in AnyValue value, CalcContext ctx, int height)
        => value.TryPickCollectionArray(out var array, ctx) && array!.Width == 1 && array.Height == height;

#pragma warning disable S3776 // Mask-shape detection then one mask walk; already reduced from 27
    private static AnyValue Filter(CalcContext ctx, Span<AnyValue> args)
    {
        if (!args[0].TryPickCollectionArray(out var array, ctx))
            return XLError.IncompatibleValue;
        if (!args[1].TryPickCollectionArray(out var include, ctx))
            return XLError.IncompatibleValue;

        var height = array!.Height;
        var width = array.Width;

        // The mask selects rows when it's a column vector matching the height, or columns when it's
        // a row vector matching the width.
        bool filterRows;
        if (include!.Width == 1 && include.Height == height)
            filterRows = true;
        else if (include.Height == 1 && include.Width == width)
            filterRows = false;
        else
            return XLError.IncompatibleValue;

        // Keeping columns is keeping the rows of the transpose, so only the row case is written out.
        var source = filterRows ? array : new TransposedArray(array);

        var kept = new List<int>();
        for (var i = 0; i < source.Height; i++)
        {
            var mask = filterRows ? include[i, 0] : include[0, i];
            if (!mask.TryCoerceLogicalOrBlankOrNumberOrText(out var flag, out var maskError))
                return maskError;
            if (flag)
                kept.Add(i);
        }

        if (kept.Count == 0)
            return args.Length > 2 ? args[2] : XLError.CellReference;

        return Orient(BuildRows(source, kept), !filterRows);
    }
#pragma warning restore S3776

#pragma warning disable S3776 // Six optional arguments to validate before one lookup; already reduced from 23
    private static AnyValue XLookup(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryScalarArg(ctx, args[0], out var lookupValue))
            return XLError.IncompatibleValue;
        if (lookupValue.TryPickError(out var lookupError))
            return lookupError;

        if (!args[1].TryPickCollectionArray(out var lookupArray, ctx))
            return XLError.IncompatibleValue;
        if (!args[2].TryPickCollectionArray(out var returnArray, ctx))
            return XLError.IncompatibleValue;

        var matchMode = 0;
        if (args.Length > 4 && !TryIntArg(ctx, args[4], out matchMode, out var matchModeError))
            return matchModeError;

        var searchMode = 1;
        if (args.Length > 5 && !TryIntArg(ctx, args[5], out searchMode, out var searchModeError))
            return searchModeError;

        var vertical = !(lookupArray!.Height == 1 && lookupArray.Width > 1);
        var length = vertical ? lookupArray.Height : lookupArray.Width;

        var index = FindMatch(lookupArray, vertical, lookupValue, matchMode, searchMode);
        if (index < 0)
            return args.Length > 3 ? args[3] : XLError.NoValueAvailable;

        // Return the matching row (vertical lookup) or column (horizontal lookup) of return_array.
        // A horizontal lookup returns the matching row of the transpose, so only one case is written
        // out and the result is turned back the right way round at the end.
        var source = vertical ? returnArray! : new TransposedArray(returnArray!);
        if (source.Height != length)
            return XLError.IncompatibleValue;
        if (source.Width == 1)
            return source[index, 0].ToAnyValue();

        var line = new ScalarValue[1, source.Width];
        for (var c = 0; c < source.Width; c++)
            line[0, c] = source[index, c];

        return Orient(new ConstArray(line), !vertical);
    }
#pragma warning restore S3776

    private static AnyValue XMatch(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryScalarArg(ctx, args[0], out var lookupValue))
            return XLError.IncompatibleValue;
        if (lookupValue.TryPickError(out var lookupError))
            return lookupError;

        if (!args[1].TryPickCollectionArray(out var lookupArray, ctx))
            return XLError.IncompatibleValue;

        var matchMode = 0;
        if (args.Length > 2 && !TryIntArg(ctx, args[2], out matchMode, out var matchModeError))
            return matchModeError;

        var searchMode = 1;
        if (args.Length > 3 && !TryIntArg(ctx, args[3], out searchMode, out var searchModeError))
            return searchModeError;

        var vertical = !(lookupArray!.Height == 1 && lookupArray.Width > 1);
        var index = FindMatch(lookupArray, vertical, lookupValue, matchMode, searchMode);
        return index < 0 ? XLError.NoValueAvailable : index + 1;
    }

    /// <summary>
    /// Find the index of <paramref name="target"/> within a one-dimensional lookup array, honouring
    /// XLOOKUP/XMATCH match modes (0 exact, -1 exact-or-next-smaller, 1 exact-or-next-larger,
    /// 2 wildcard) and search modes (1 first-to-last, -1 last-to-first; binary modes fall back to a
    /// linear scan, which is correct if slower).
    /// </summary>
#pragma warning disable S3776 // Four match modes over one scan; the modes are the specification
    private static int FindMatch(Array array, bool vertical, ScalarValue target, int matchMode, int searchMode)
    {
        var length = vertical ? array.Height : array.Width;
        var comparer = ScalarValueComparer.SortIgnoreCase;

        if (matchMode == 2 && target.TryPickText(out var pattern, out _))
        {
            var wildcard = new Wildcard(pattern!);
            foreach (var i in SearchOrder(length, searchMode))
            {
                if (Element(array, vertical, i).TryPickText(out var text, out _) && wildcard.Matches(text!.AsSpan()))
                    return i;
            }

            return -1;
        }

        var best = -1;
        var bestValue = ScalarValue.Blank;
        foreach (var i in SearchOrder(length, searchMode))
        {
            var value = Element(array, vertical, i);
            if (!target.HaveSameType(value))
                continue;

            var compare = comparer.Compare(value, target);
            if (compare == 0)
                return i;

#pragma warning disable S1871 // The two arms are the two match modes; one condition would hide that
            if (matchMode == -1 && compare < 0 && (best == -1 || comparer.Compare(value, bestValue) > 0))
            {
                best = i;
                bestValue = value;
            }
            else if (matchMode == 1 && compare > 0 && (best == -1 || comparer.Compare(value, bestValue) < 0))
            {
                best = i;
                bestValue = value;
            }
#pragma warning restore S1871
        }

        return best;
    }
#pragma warning restore S3776

    private static ScalarValue Element(Array array, bool vertical, int index)
        => vertical ? array[index, 0] : array[0, index];

    private static IEnumerable<int> SearchOrder(int length, int searchMode)
    {
        if (searchMode == -1)
        {
            for (var i = length - 1; i >= 0; i--)
                yield return i;
        }
        else
        {
            for (var i = 0; i < length; i++)
                yield return i;
        }
    }

    private static bool RowsEqual(Array array, int rowA, int rowB, int width)
    {
        for (var c = 0; c < width; c++)
        {
            if (ScalarValueComparer.SortIgnoreCase.Compare(array[rowA, c], array[rowB, c]) != 0)
                return false;
        }

        return true;
    }

    private static ConstArray BuildRows(Array source, List<int> rowOrder)
    {
        var result = new ScalarValue[rowOrder.Count, source.Width];
        for (var i = 0; i < rowOrder.Count; i++)
        {
            for (var c = 0; c < source.Width; c++)
                result[i, c] = source[rowOrder[i], c];
        }

        return new ConstArray(result);
    }

    private static AnyValue Orient(Array array, bool transposed)
        => transposed ? new TransposedArray(array) : array;

    private static bool TryScalarArg(CalcContext ctx, in AnyValue arg, out ScalarValue scalar)
        => arg.TryReduceToScalar(ctx, out scalar, out _);

    private static bool TryNumberArg(CalcContext ctx, in AnyValue arg, out double number, out XLError error)
    {
        error = default;
        if (!TryScalarArg(ctx, arg, out var scalar))
        {
            number = 0;
            error = XLError.IncompatibleValue;
            return false;
        }

        return scalar.ToNumber(ctx.Culture).TryPickT0(out number, out error);
    }

    private static bool TryIntArg(CalcContext ctx, in AnyValue arg, out int value, out XLError error)
    {
        if (!TryNumberArg(ctx, arg, out var number, out error))
        {
            value = 0;
            return false;
        }

        value = (int)Math.Truncate(number);
        return true;
    }

    private static bool TryBoolArg(CalcContext ctx, in AnyValue arg, out bool value, out XLError error)
    {
        error = default;
        if (!TryScalarArg(ctx, arg, out var scalar))
        {
            value = false;
            error = XLError.IncompatibleValue;
            return false;
        }

        return scalar.TryCoerceLogicalOrBlankOrNumberOrText(out value, out error);
    }
}
