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
    public static void Register(FunctionRegistry ce)
    {
        ce.RegisterFunction("SEQUENCE", 1, 4, Sequence, FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Generates a list of sequential numbers
        ce.RegisterFunction("UNIQUE", 1, 3, Unique, FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Returns the distinct values from a range or array
        ce.RegisterFunction("SORT", 1, 4, Sort, FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Sorts the contents of a range or array
        ce.RegisterFunction("SORTBY", 2, 255, SortBy, FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Sorts a range or array based on the values in a corresponding range or array
        ce.RegisterFunction("FILTER", 2, 3, Filter, FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Filters a range or array based on criteria
        ce.RegisterFunction("XLOOKUP", 3, 6, XLookup, FunctionFlags.Range | FunctionFlags.ReturnsArray, AllowRange.All); // Searches a range or array and returns the matching item(s)
        ce.RegisterFunction("XMATCH", 2, 4, XMatch, FunctionFlags.Range, AllowRange.All); // Returns the relative position of an item in a range or array
    }

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
            if (i < args.Length && args[i].TryPickScalar(out _, out _))
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

        var kept = new List<int>();
        var count = filterRows ? height : width;
        for (var i = 0; i < count; i++)
        {
            var mask = filterRows ? include[i, 0] : include[0, i];
            if (!mask.TryCoerceLogicalOrBlankOrNumberOrText(out var flag, out var maskError))
                return maskError;
            if (flag)
                kept.Add(i);
        }

        if (kept.Count == 0)
            return args.Length > 2 ? args[2] : XLError.CellReference;

        if (filterRows)
        {
            var rows = new ScalarValue[kept.Count, width];
            for (var i = 0; i < kept.Count; i++)
            {
                for (var c = 0; c < width; c++)
                    rows[i, c] = array[kept[i], c];
            }

            return new ConstArray(rows);
        }

        var columns = new ScalarValue[height, kept.Count];
        for (var i = 0; i < kept.Count; i++)
        {
            for (var r = 0; r < height; r++)
                columns[r, i] = array[r, kept[i]];
        }

        return new ConstArray(columns);
    }

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
        if (vertical)
        {
            if (returnArray!.Height != length)
                return XLError.IncompatibleValue;
            if (returnArray.Width == 1)
                return returnArray[index, 0].ToAnyValue();

            var row = new ScalarValue[1, returnArray.Width];
            for (var c = 0; c < returnArray.Width; c++)
                row[0, c] = returnArray[index, c];
            return new ConstArray(row);
        }

        if (returnArray!.Width != length)
            return XLError.IncompatibleValue;
        if (returnArray.Height == 1)
            return returnArray[0, index].ToAnyValue();

        var column = new ScalarValue[returnArray.Height, 1];
        for (var r = 0; r < returnArray.Height; r++)
            column[r, 0] = returnArray[r, index];
        return new ConstArray(column);
    }

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
        }

        return best;
    }

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

    private static ConstArray BuildRows(Array source, IReadOnlyList<int> rowOrder)
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
    {
        if (arg.TryPickScalar(out scalar, out _))
            return true;

        return arg.ImplicitIntersection(ctx).TryPickScalar(out scalar, out _);
    }

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
