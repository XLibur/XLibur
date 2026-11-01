using System;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine.Functions;

/// <summary>
/// The functions that take a <c>LAMBDA</c> and call it for each element, row or column of an array:
/// <c>MAP</c>, <c>REDUCE</c>, <c>SCAN</c>, <c>BYROW</c>, <c>BYCOL</c> and <c>MAKEARRAY</c> (#847, spec 08
/// phase 3). The function can also be one of Excel's passed by its name, <c>BYROW(D1:E2,SUM)</c>.
/// </summary>
/// <remarks>
/// <para>
/// Every rule here is what Excel 365 calculated for the formulas in <c>LambdaHelperTests</c>, typed over
/// COM. An element of a range is passed as a reference to its cell, so <c>MAP(A1:A3,LAMBDA(v,ROW(v)))</c>
/// is 1, 2, 3, and a row of a range as a reference to the row. A function whose parameters do not take
/// the arguments the helper passes is <c>#VALUE!</c> for the whole result. A result for an element that
/// is itself many values, such as an array, makes the whole result <c>#CALC!</c>.
/// </para>
/// <para>
/// They take every argument as it is, a range staying a reference, and return their array whole, so
/// they are registered as functions that return an array and accept a range for every parameter.
/// </para>
/// </remarks>
internal static class LambdaHelpers
{
    private const FunctionFlags Spilling = FunctionFlags.Range | FunctionFlags.ReturnsArray;

    public static void Register(FunctionRegistry ce)
    {
        ce.RegisterFunction("BYCOL", 2, 2, ByCol, Spilling, AllowRange.All); // Applies a LAMBDA to each column and returns an array of the results
        ce.RegisterFunction("BYROW", 2, 2, ByRow, Spilling, AllowRange.All); // Applies a LAMBDA to each row and returns an array of the results
        ce.RegisterFunction("MAKEARRAY", 3, 3, MakeArray, Spilling, AllowRange.All); // Returns a calculated array of a specified row and column size, by applying a LAMBDA
        ce.RegisterFunction("MAP", 2, 255, Map, Spilling, AllowRange.All); // Returns an array formed by mapping each value in the array(s) to a new value by applying a LAMBDA
        ce.RegisterFunction("REDUCE", 2, 3, Reduce, Spilling, AllowRange.All); // Reduces an array to an accumulated value by applying a LAMBDA to each value
        ce.RegisterFunction("SCAN", 2, 3, Scan, Spilling, AllowRange.All); // Scans an array by applying a LAMBDA to each value and returns an array that has each intermediate value
    }

    /// <summary>
    /// <c>MAP(array1, [array2, …], lambda)</c>: the function of the elements at each position of the
    /// arrays. An array smaller than the others has <c>#N/A</c> where it has no element.
    /// </summary>
    private static AnyValue Map(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetFunction(args[^1], args.Length - 1, out var lambda, out var error))
            return error;

        var arrays = new Elements[args.Length - 1];
        var rows = 0;
        var columns = 0;
        for (var i = 0; i < arrays.Length; ++i)
        {
            if (!Elements.TryCreate(args[i], ctx, out arrays[i], out error))
                return error;

            rows = Math.Max(rows, arrays[i].Rows);
            columns = Math.Max(columns, arrays[i].Columns);
        }

        var result = new ScalarValue[rows, columns];
        var arguments = new AnyValue?[arrays.Length];
        for (var row = 0; row < rows; ++row)
        {
            for (var column = 0; column < columns; ++column)
            {
                if (!TryGetElements(arrays, row, column, arguments))
                {
                    result[row, column] = XLError.NoValueAvailable;
                    continue;
                }

                if (!TryGetSingle(ctx.CalcEngine.Invoke(ctx, lambda, arguments), ctx, out result[row, column]))
                    return XLError.Calc;
            }
        }

        return new ConstArray(result);
    }

    /// <summary>
    /// <c>REDUCE([initial_value], array, lambda(accumulator, value))</c>: the accumulator after the
    /// function has taken each element in turn, by rows. An initial value left out is blank. The result
    /// can be many values, which spill.
    /// </summary>
    private static AnyValue Reduce(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetAccumulation(ctx, args, out var accumulator, out var array, out var lambda, out var error))
            return error;

        var arguments = new AnyValue?[2];
        for (var row = 0; row < array.Rows; ++row)
        {
            for (var column = 0; column < array.Columns; ++column)
            {
                arguments[0] = accumulator;
                arguments[1] = array[row, column];
                accumulator = ctx.CalcEngine.Invoke(ctx, lambda, arguments);
            }
        }

        return accumulator;
    }

    /// <summary>
    /// <c>SCAN([initial_value], array, lambda(accumulator, value))</c>: as <c>REDUCE</c>, with the
    /// accumulator after each element, in an array the shape of <c>array</c>.
    /// </summary>
    private static AnyValue Scan(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetAccumulation(ctx, args, out var accumulator, out var array, out var lambda, out var error))
            return error;

        var result = new ScalarValue[array.Rows, array.Columns];
        var arguments = new AnyValue?[2];
        for (var row = 0; row < array.Rows; ++row)
        {
            for (var column = 0; column < array.Columns; ++column)
            {
                arguments[0] = accumulator;
                arguments[1] = array[row, column];
                accumulator = ctx.CalcEngine.Invoke(ctx, lambda, arguments);
                if (!TryGetSingle(accumulator, ctx, out result[row, column]))
                    return XLError.Calc;
            }
        }

        return new ConstArray(result);
    }

    /// <summary><c>BYROW(array, lambda(row))</c>: a column of the function of each row.</summary>
    private static AnyValue ByRow(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetFunction(args[1], 1, out var lambda, out var error))
            return error;

        if (!Elements.TryCreate(args[0], ctx, out var array, out error))
            return error;

        var result = new ScalarValue[array.Rows, 1];
        var arguments = new AnyValue?[1];
        for (var row = 0; row < array.Rows; ++row)
        {
            arguments[0] = array.Row(row);
            if (!TryGetSingle(ctx.CalcEngine.Invoke(ctx, lambda, arguments), ctx, out result[row, 0]))
                return XLError.Calc;
        }

        return new ConstArray(result);
    }

    /// <summary><c>BYCOL(array, lambda(column))</c>: a row of the function of each column.</summary>
    private static AnyValue ByCol(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetFunction(args[1], 1, out var lambda, out var error))
            return error;

        if (!Elements.TryCreate(args[0], ctx, out var array, out error))
            return error;

        var result = new ScalarValue[1, array.Columns];
        var arguments = new AnyValue?[1];
        for (var column = 0; column < array.Columns; ++column)
        {
            arguments[0] = array.Column(column);
            if (!TryGetSingle(ctx.CalcEngine.Invoke(ctx, lambda, arguments), ctx, out result[0, column]))
                return XLError.Calc;
        }

        return new ConstArray(result);
    }

    /// <summary>
    /// <c>MAKEARRAY(rows, columns, lambda(row, column))</c>: the function of each row and column number,
    /// from 1. The size must be whole numbers from 1: Excel gives <c>#VALUE!</c> for 0, and for 1.9.
    /// </summary>
    private static AnyValue MakeArray(CalcContext ctx, Span<AnyValue> args)
    {
        if (!TryGetSize(args[0], ctx, XLHelper.MaxRowNumber, out var rows, out var error) ||
            !TryGetSize(args[1], ctx, XLHelper.MaxColumnNumber, out var columns, out error))
            return error;

        if (!TryGetFunction(args[2], 2, out var lambda, out error))
            return error;

        var result = new ScalarValue[rows, columns];
        var arguments = new AnyValue?[2];
        for (var row = 0; row < rows; ++row)
        {
            for (var column = 0; column < columns; ++column)
            {
                arguments[0] = row + 1;
                arguments[1] = column + 1;
                if (!TryGetSingle(ctx.CalcEngine.Invoke(ctx, lambda, arguments), ctx, out result[row, column]))
                    return XLError.Calc;
            }
        }

        return new ConstArray(result);
    }

    /// <summary>
    /// The arguments of <c>REDUCE</c> and <c>SCAN</c>: the initial value, which is blank when it is left
    /// out, the array, and the function of the accumulator and a value.
    /// </summary>
    private static bool TryGetAccumulation(CalcContext ctx, Span<AnyValue> args, out AnyValue initial,
        out Elements array, out LambdaFunction lambda, out XLError error)
    {
        initial = args.Length == 3 ? args[0] : AnyValue.Blank;
        array = default;
        return TryGetFunction(args[^1], 2, out lambda, out error) &&
               Elements.TryCreate(args[^2], ctx, out array, out error);
    }

    /// <summary>
    /// The function an argument holds, which must take <paramref name="argumentCount"/> arguments. One of
    /// Excel's functions is checked only when it is called, as Excel gives <c>#VALUE!</c> for each
    /// element of <c>MAP(A1:A3,IF)</c> and for the whole of <c>MAP(A1:A3,LAMBDA(a,b,a))</c>.
    /// </summary>
    private static bool TryGetFunction(in AnyValue value, int argumentCount, out LambdaFunction lambda, out XLError error)
    {
        if (!value.TryPickLambda(out var function))
        {
            lambda = null!;
            error = value.TryPickError(out var valueError) ? valueError : XLError.IncompatibleValue;
            return false;
        }

        lambda = function;
        error = XLError.IncompatibleValue;
        return function.Function is not null ||
               (argumentCount <= function.Parameters.Count && argumentCount >= function.RequiredParameterCount);
    }

    /// <summary>The elements of the arrays at a position, or <c>false</c> when one of them has none there.</summary>
    private static bool TryGetElements(Elements[] arrays, int row, int column, AnyValue?[] elements)
    {
        for (var i = 0; i < arrays.Length; ++i)
        {
            if (row >= arrays[i].Rows || column >= arrays[i].Columns)
                return false;

            elements[i] = arrays[i][row, column];
        }

        return true;
    }

    /// <summary>
    /// The one value a call gave for an element. <c>false</c> when it gave many, such as an array, or a
    /// function, which make the whole result <c>#CALC!</c>.
    /// </summary>
    private static bool TryGetSingle(in AnyValue result, CalcContext ctx, out ScalarValue value)
    {
        if (result.IsLambda)
        {
            value = default;
            return false;
        }

        if (result.TryPickSingleOrMultiValue(out value, out var array, ctx))
            return true;

        if (array!.Width == 1 && array.Height == 1)
        {
            value = array[0, 0];
            return true;
        }

        return false;
    }

    private static bool TryGetSize(in AnyValue value, CalcContext ctx, int max, out int size, out XLError error)
    {
        size = 0;
        if (!value.TryReduceToNumber(ctx, out var number, out error))
            return false;

        error = XLError.IncompatibleValue;
        if (number < 1 || number > max || number != Math.Floor(number))
            return false;

        size = (int)number;
        return true;
    }

    /// <summary>
    /// An array argument, element by element: a range gives each cell, row or column as a reference, an
    /// array its values, and a single value is a 1×1 array.
    /// </summary>
    private readonly struct Elements
    {
        private readonly XLRangeAddress _area;
        private readonly Array? _array;
        private readonly AnyValue _single;
        private readonly bool _isArea;

        private Elements(XLRangeAddress area)
        {
            _area = area;
            _isArea = true;
            Rows = area.RowSpan;
            Columns = area.ColumnSpan;
        }

        private Elements(Array array)
        {
            _array = array;
            Rows = array.Height;
            Columns = array.Width;
        }

        private Elements(AnyValue single)
        {
            _single = single;
            Rows = 1;
            Columns = 1;
        }

        internal int Rows { get; }

        internal int Columns { get; }

        internal AnyValue this[int row, int column] => _isArea
            ? Area(row, column, row, column)
            : _array?[row, column].ToAnyValue() ?? _single;

        internal AnyValue Row(int row) => _isArea
            ? Area(row, 0, row, Columns - 1)
            : _array is null ? _single : Slice(row, 0, 1, Columns);

        internal AnyValue Column(int column) => _isArea
            ? Area(0, column, Rows - 1, column)
            : _array is null ? _single : Slice(0, column, Rows, 1);

        internal static bool TryCreate(in AnyValue value, CalcContext ctx, out Elements elements, out XLError error)
        {
            error = default;
            if (value.IsReference)
            {
                elements = default;
                if (!value.TryPickArea(out var area, out error))
                    return false;

                elements = new Elements(area);
                return true;
            }

            if (value.TryPickArray(out var array))
            {
                elements = new Elements(array!);
                return true;
            }

            if (value.IsLambda)
            {
                elements = default;
                error = XLError.IncompatibleValue;
                return false;
            }

            elements = new Elements(value);
            return true;
        }

        private AnyValue Area(int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            var top = _area.FirstAddress.RowNumber;
            var left = _area.FirstAddress.ColumnNumber;
            var area = new Area(top + firstRow, left + firstColumn, top + lastRow, left + lastColumn);
            return new Reference(XLRangeAddress.FromSheetRange(_area.Sheet as XLWorksheet, area));
        }

        private ConstArray Slice(int firstRow, int firstColumn, int rows, int columns)
        {
            var slice = new ScalarValue[rows, columns];
            for (var row = 0; row < rows; ++row)
            {
                for (var column = 0; column < columns; ++column)
                    slice[row, column] = _array![firstRow + row, firstColumn + column];
            }

            return new ConstArray(slice);
        }
    }
}
