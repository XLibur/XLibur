using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// A base class for an 2D array. Every array is at least 1x1.
/// </summary>
internal abstract class Array : IEnumerable<ScalarValue>
{
    /// <summary>
    /// Width of the array, at least 1.
    /// </summary>
    public abstract int Width { get; }

    /// <summary>
    /// Height of the array, at least 1.
    /// </summary>
    public abstract int Height { get; }

    /// <summary>
    /// Get a value at a specified coordinate.
    /// </summary>
    /// <param name="y">Uses 0-based notation.</param>
    /// <param name="x">Uses 0-based notation.</param>
    public abstract ScalarValue this[int y, int x] { get; }

    /// <summary>
    /// An iterator over all elements of an array, from top to bottom, from left to right.
    /// </summary>
    public virtual IEnumerator<ScalarValue> GetEnumerator() => FlattenArray().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    protected IEnumerable<ScalarValue> FlattenArray()
    {
        for (var row = 0; row < Height; row++)
        {
            for (var col = 0; col < Width; col++)
            {
                yield return this[row, col];
            }
        }
    }

    /// <summary>
    /// Return a new array that was created by applying a function to each element of the array.
    /// </summary>
    /// <remarks>
    /// The result is a lazy view — see <see cref="MappedArray"/>.
    /// </remarks>
    public Array Apply(Func<ScalarValue, ScalarValue> op) => new MappedArray(this, op);

    /// <summary>
    /// Return a new array that was created by applying a function to each element of the left and right array.
    /// Arrays can have different size and missing values are replaced by <c>#N/A</c>.
    /// </summary>
    /// <remarks>
    /// The result is a lazy view — see <see cref="BinaryArray"/>.
    /// </remarks>
    public Array Apply(Array rightArray, BinaryFunc func, CalcContext ctx) => new BinaryArray(this, rightArray, func, ctx);

    /// <summary>
    /// Broadcast array for calculation of array formulas.
    /// </summary>
    public Array Broadcast(int rows, int columns)
    {
        if (Width == columns && Height == rows)
            return this;

        if (Width == 1 && Height == 1)
            return new ScalarArray(this[0, 0], columns, rows);

        if (Width == 1)
            return new RepeatedColumnArray(this, rows, columns);

        if (Height == 1)
            return new RepeatedRowArray(this, rows, columns);

        return new ResizedArray(this, rows, columns);
    }
}

/// <summary>
/// The element-wise result of a binary operator over two arrays, computed on access rather than
/// stored. Where the two differ in size, the missing side is <c>#N/A</c>.
/// </summary>
/// <remarks>
/// <para>
/// This is the same shape as <see cref="ReferenceArray"/> and the broadcast views: cheap element
/// access, no backing storage. It exists because the eager version allocated
/// <c>ScalarValue[height, width]</c> for the whole rectangle whatever the consumer wanted from it —
/// and a whole-column operand is 1,048,576 elements, about 24 MB. A formula that then keeps one
/// element, which is what a legacy formula does, paid all of it. The fuzz corpus holds a 16-byte
/// input spanning 566 columns that allocated roughly 13.6 GB and held a core for six minutes (D38).
/// </para>
/// <para>
/// The trade the lazy form makes: a consumer that reads the same element twice computes it twice.
/// Element access is cheap and allocation-free, which is what makes that acceptable, but it does
/// mean an operator chain costs its depth per access rather than being flattened once.
/// </para>
/// <para>
/// <b>The view must not outlive the evaluation that built it.</b> The <see cref="CalcContext"/> is
/// captured, and reading a cell through it can raise <c>GettingDataException</c> to demand that a
/// dirty precedent be calculated first — which only the enclosing evaluation knows how to answer.
/// <see cref="ReferenceArray"/> has always carried the same constraint.
/// </para>
/// </remarks>
internal sealed class BinaryArray : Array
{
    private readonly Array _left;
    private readonly Array _right;
    private readonly BinaryFunc _func;
    private readonly CalcContext _ctx;

    public BinaryArray(Array left, Array right, BinaryFunc func, CalcContext ctx)
    {
        _left = left;
        _right = right;
        _func = func;
        _ctx = ctx;
        Width = Math.Max(left.Width, right.Width);
        Height = Math.Max(left.Height, right.Height);
    }

    public override int Width { get; }

    public override int Height { get; }

    public override ScalarValue this[int y, int x]
    {
        get
        {
            if (y < 0 || y >= Height || x < 0 || x >= Width)
                throw new ArgumentOutOfRangeException(nameof(y), "Index was out of range.");

            var leftItem = x < _left.Width && y < _left.Height ? _left[y, x] : XLError.NoValueAvailable;
            var rightItem = x < _right.Width && y < _right.Height ? _right[y, x] : XLError.NoValueAvailable;
            return _func(in leftItem, in rightItem, _ctx);
        }
    }
}

/// <summary>
/// The result of a unary function over each element of an array, computed on access rather than
/// stored. The lazy counterpart of <see cref="BinaryArray"/>, and it carries the same caveats.
/// </summary>
internal sealed class MappedArray : Array
{
    private readonly Array _original;
    private readonly Func<ScalarValue, ScalarValue> _op;

    public MappedArray(Array original, Func<ScalarValue, ScalarValue> op)
    {
        _original = original;
        _op = op;
    }

    public override int Width => _original.Width;

    public override int Height => _original.Height;

    public override ScalarValue this[int y, int x] => _op(_original[y, x]);
}

/// <summary>
/// An array of scalar values.
/// </summary>
internal sealed class ConstArray : Array
{
    private readonly ScalarValue[,] _data;

    public ConstArray(ScalarValue[,] data)
    {
        if (data.GetLength(0) < 1 || data.GetLength(1) < 1)
            throw new ArgumentException("Array must be at least 1x1.", nameof(data));
        _data = data;
    }

    public override ScalarValue this[int y, int x] => _data[y, x];

    public override int Width => _data.GetLength(1);

    public override int Height => _data.GetLength(0);
}

/// <summary>
/// Array for array literal from a parser. It uses a 1D array of values as a storage.
/// </summary>
internal sealed class LiteralArray : Array
{
    private readonly IReadOnlyList<ScalarValue> _elements;

    /// <summary>
    /// Create a new instance of a <see cref="LiteralArray"/>.
    /// </summary>
    /// <param name="rows">Number of rows of an array/</param>
    /// <param name="columns">Number of columns of an array.</param>
    /// <param name="elements">Row by row data of the array. Has the expected size of an array.</param>
    public LiteralArray(int rows, int columns, IReadOnlyList<ScalarValue> elements)
    {
        if (rows * columns != elements.Count)
            throw new ArgumentException("Number of elements in not the same as size of an array.", nameof(elements));

        Height = rows;
        Width = columns;
        _elements = elements;
    }

    public override ScalarValue this[int y, int x]
    {
        get
        {
            if (x < 0 || x >= Width)
                throw new ArgumentOutOfRangeException(nameof(x));

            return _elements[y * Width + x];
        }
    }

    public override int Width { get; }

    public override int Height { get; }
}

/// <summary>
/// A special case of an array that is actually only numbers.
/// </summary>
internal sealed class NumberArray : Array
{
    private readonly double[,] _data;

    public NumberArray(double[,] data)
    {
        _data = data;
    }

    public override ScalarValue this[int y, int x] => _data[y, x];

    public override int Width => _data.GetLength(1);

    public override int Height => _data.GetLength(0);
}

/// <summary>
/// An array that retrieves its value directly from the worksheet without allocating extra memory.
/// </summary>
internal sealed class ReferenceArray : Array
{
    private readonly XLRangeAddress _area;
    private readonly CalcContext _context;
    private readonly int _offsetColumn;
    private readonly int _offsetRow;

    public ReferenceArray(XLRangeAddress area, CalcContext context)
    {
        _area = area;
        _context = context;
        _offsetColumn = _area.FirstAddress.ColumnNumber;
        _offsetRow = area.FirstAddress.RowNumber;
    }

    public override ScalarValue this[int y, int x] => _context.GetCellValue(_area.Worksheet, y + _offsetRow, x + _offsetColumn);

    public override int Width => _area.ColumnSpan;

    public override int Height => _area.RowSpan;
}

internal sealed class RepeatedColumnArray : Array
{
    private readonly Array _columnArray;

    public RepeatedColumnArray(Array oneColumnArray, int rows, int columns)
    {
        Debug.Assert(oneColumnArray.Width == 1);
        _columnArray = oneColumnArray;
        Width = columns;
        Height = rows;
    }

    public override int Width { get; }

    public override int Height { get; }

    public override ScalarValue this[int row, int column]
    {
        get
        {
            if (row >= Height || column >= Width)
                throw new ArgumentOutOfRangeException(nameof(row), "Index was out of range.");

            if (row >= _columnArray.Height)
                return XLError.NoValueAvailable;

            return _columnArray[row, 0];
        }
    }
}

internal sealed class RepeatedRowArray : Array
{
    private readonly Array _rowArray;

    internal RepeatedRowArray(Array oneRowArray, int rows, int columns)
    {
        Debug.Assert(oneRowArray.Height == 1);
        _rowArray = oneRowArray;
        Width = columns;
        Height = rows;
    }

    public override int Width { get; }

    public override int Height { get; }

    public override ScalarValue this[int row, int column]
    {
        get
        {
            if (row >= Height || column >= Width)
                throw new ArgumentOutOfRangeException(nameof(row), "Index was out of range.");

            if (column >= _rowArray.Width)
                return XLError.NoValueAvailable;

            return _rowArray[0, column];
        }
    }
}

/// <summary>
/// A resize array from another array. Extra items without value have <c>#N/A</c>.
/// </summary>
internal sealed class ResizedArray : Array
{
    private readonly Array _original;

    public ResizedArray(Array original, int rows, int columns)
    {
        _original = original;
        Height = rows;
        Width = columns;
    }

    public override int Width { get; }

    public override int Height { get; }

    public override ScalarValue this[int y, int x]
    {
        get
        {
            if (y >= Height || x >= Width)
                throw new ArgumentOutOfRangeException(nameof(y), "Index was out of range.");

            return y < _original.Height && x < _original.Width
                ? _original[y, x]
                : XLError.NoValueAvailable;
        }
    }
}

/// <summary>
/// An array where all elements have the same value.
/// </summary>
internal sealed class ScalarArray : Array
{
    private readonly ScalarValue _value;

    public ScalarArray(ScalarValue value, int columns, int rows)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(columns, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(rows, 1);
        _value = value;
        Width = columns;
        Height = rows;
    }

    public override int Width { get; }

    public override int Height { get; }

    public override ScalarValue this[int y, int x]
    {
        get
        {
            if (x < 0 || x >= Width || y < 0 || y >= Height)
                throw new ArgumentOutOfRangeException(nameof(y), "Index was out of range.");

            return _value;
        }
    }

    public override IEnumerator<ScalarValue> GetEnumerator()
    {
        return Enumerable.Range(0, Width * Height).Select(_ => _value).GetEnumerator();
    }
}

internal sealed class TransposedArray : Array
{
    private readonly Array _original;

    public TransposedArray(Array original)
    {
        _original = original;
    }

    public override ScalarValue this[int y, int x] => _original[x, y];

    public override int Width => _original.Height;

    public override int Height => _original.Width;
}

/// <summary>
/// An array that is a rectangular slice of the original array.
/// </summary>
internal sealed class SlicedArray : Array
{
    private readonly Array _original;
    private readonly int _rowOfs;
    private readonly int _colOfs;

    /// <summary>
    /// Create a sliced array from the original array.
    /// </summary>
    /// <param name="original">Original array.</param>
    /// <param name="rowOfs">The row offset indicating the starting row of the slice in the original array.</param>
    /// <param name="rows">The number of rows in the sliced array.</param>
    /// <param name="colOfs">The column offset indicating the starting column of the slice in the original array.</param>
    /// <param name="cols">The number of columns in the sliced array.</param>
    public SlicedArray(Array original, int rowOfs, int rows, int colOfs, int cols)
    {
        ArgumentNullException.ThrowIfNull(original);

        if (rowOfs < 0 || rows < 1 || colOfs < 0 || cols < 1 ||
            rowOfs + rows > original.Height ||
            colOfs + cols > original.Width)
            throw new ArgumentOutOfRangeException(nameof(original), "Slice dimensions exceed the bounds of the original array.");

        _original = original;
        _rowOfs = rowOfs;
        Height = rows;
        _colOfs = colOfs;
        Width = cols;
    }

    public override ScalarValue this[int y, int x] => _original[y + _rowOfs, x + _colOfs];

    public override int Width { get; }

    public override int Height { get; }
}
