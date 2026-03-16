using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace XLibur.Excel;

/// <summary>
/// An point (address) in a worksheet, an equivalent of <c>ST_CellRef</c>.
/// Row and column are packed into a single <c>ulong</c> as
/// <c>[row0:20 bits][column0:14 bits]</c> (0-based internally) for fast
/// equality, hashing, and row-major comparison.
/// </summary>
/// <remarks>Unlike the XLAddress, sheet can never be invalid.</remarks>
[DebuggerDisplay("{XLHelper.GetColumnLetterFromNumber(Column)+Row}")]
internal readonly struct XLSheetPoint : IEquatable<XLSheetPoint>, IComparable<XLSheetPoint>
{
    internal const int ColumnBits = 14;
    private const ulong ColumnMask = (1UL << ColumnBits) - 1; // 0x3FFF

    /// <summary>
    /// Packed representation: bits 0-13 = column 0-based, bits 14-33 = row 0-based.
    /// Values are stored 0-based for bit-packing efficiency; the <see cref="Row"/> and
    /// <see cref="Column"/> properties return the 1-based values used everywhere else.
    /// </summary>
    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    internal readonly ulong PackedValue;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public XLSheetPoint(int row, int column)
    {
        PackedValue = ((ulong)(uint)(row - 1) << ColumnBits) | (uint)(column - 1);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private XLSheetPoint(ulong packed)
    {
        PackedValue = packed;
    }

    /// <summary>
    /// 1-based row number in a sheet.
    /// </summary>
    public int Row
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => (int)(PackedValue >> ColumnBits) + 1;
    }

    /// <summary>
    /// 1-based column number in a sheet.
    /// </summary>
    public int Column
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => (int)(PackedValue & ColumnMask) + 1;
    }

    public static implicit operator XLSheetRange(XLSheetPoint point)
    {
        return new XLSheetRange(point);
    }

    public override bool Equals(object? obj)
    {
        return obj is XLSheetPoint point && Equals(point);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool Equals(XLSheetPoint other)
    {
        return PackedValue == other.PackedValue;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override int GetHashCode()
    {
        return PackedValue.GetHashCode();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool operator ==(XLSheetPoint a, XLSheetPoint b)
    {
        return a.PackedValue == b.PackedValue;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool operator !=(XLSheetPoint a, XLSheetPoint b)
    {
        return a.PackedValue != b.PackedValue;
    }

    /// <summary>
    /// Get offset that must be added to <paramref name="origin"/> so we can get <paramref name="target"/>.
    /// </summary>
    public static XLSheetOffset operator -(XLSheetPoint target, XLSheetPoint origin)
    {
        return new XLSheetOffset(target.Row - origin.Row, target.Column - origin.Column);
    }

    /// <inheritdoc cref="Parse(ReadOnlySpan{char})"/>
    public static XLSheetPoint Parse(string text) => Parse(text.AsSpan());

    /// <summary>
    /// Parse point per type <c>ST_CellRef</c> from
    /// <a href="https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/db11a912-b1cb-4dff-b46d-9bedfd10cef0">2.1.1108 Part 4 Section 3.18.8, ST_CellRef (Cell Reference)</a>
    /// </summary>
    /// <param name="input">Input text</param>
    /// <exception cref="FormatException">If the input doesn't match expected grammar.</exception>
    public static XLSheetPoint Parse(ReadOnlySpan<char> input)
    {
        if (!TryParse(input, out var point))
            throw new FormatException($"Sheet point doesn't have correct format: '{input.ToString()}'.");

        return point;
    }

    /// <summary>
    /// Try to parse sheet point. Doesn't accept any extra whitespace anywhere in the input.
    /// Letters must be upper case.
    /// </summary>
    public static bool TryParse(ReadOnlySpan<char> input, out XLSheetPoint point)
    {
        point = default;

        // Don't reuse inefficient logic from XLAddress
        if (input.Length < 2)
            return false;

        var i = 0;
        var c = input[i++];
        if (!IsLetter(c))
            return false;

        var columnIndex = c - 'A' + 1;
        while (i < input.Length && IsLetter(input[i]))
        {
            c = input[i];
            columnIndex = columnIndex * 26 + c - 'A' + 1;
            i++;
        }

        if (i > 3)
            return false;

        if (i == input.Length)
            return false;

        // Everything else must be digits
        c = input[i++];

        // First letter can't be 0
        if (c is < '1' or > '9')
            return false;

        var rowIndex = c - '0';
        while (i < input.Length && IsDigit(input[i]))
        {
            c = input[i];
            rowIndex = rowIndex * 10 + c - '0';
            i++;
        }

        if (i != input.Length)
            return false;

        if (rowIndex > XLHelper.MaxRowNumber || columnIndex > XLHelper.MaxColumnNumber)
            return false;

        point = new XLSheetPoint(rowIndex, columnIndex);
        return true;

        static bool IsLetter(char c) => c is >= 'A' and <= 'Z';
        static bool IsDigit(char c) => c is >= '0' and <= '9';
    }

    /// <summary>
    /// Write the sheet point as a reference to the span (e.g. <c>A1</c>).
    /// </summary>
    /// <param name="output">Must be at least 10 chars long</param>
    /// <returns>Number of chars </returns>
    public int Format(Span<char> output)
    {
        var columnLetters = XLHelper.GetColumnLetterFromNumber(Column);
        for (var i = 0; i < columnLetters.Length; ++i)
            output[i] = columnLetters[i];

        var row = Row;
        var digitCount = GetDigitCount(row);
        var rowRemainder = row;
        var formattedLength = digitCount + columnLetters.Length;
        for (var i = formattedLength - 1; i >= columnLetters.Length; --i)
        {
            var digit = rowRemainder % 10;
            rowRemainder /= 10;
            output[i] = (char)(digit + '0');
        }

        return formattedLength;
    }

    public override string ToString()
    {
        Span<char> text = stackalloc char[10];
        var len = Format(text);
        return text.Slice(0, len).ToString();
    }

    private static int GetDigitCount(int n)
    {
        if (n < 10L) return 1;
        if (n < 100L) return 2;
        if (n < 1000L) return 3;
        if (n < 10000L) return 4;
        if (n < 100000L) return 5;
        if (n < 1000000L) return 6;
        return 7; // Row can't have more digits
    }

    /// <summary>
    /// Create a sheet point from the address. Workbook is ignored.
    /// </summary>
    public static XLSheetPoint FromAddress(IXLAddress address)
        => new(address.RowNumber, address.ColumnNumber);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public int CompareTo(XLSheetPoint other)
    {
        return PackedValue.CompareTo(other.PackedValue);
    }

    /// <summary>
    /// Is the point within the range or below the range?
    /// </summary>
    internal bool InRangeOrBelow(in XLSheetRange range)
    {
        return Row >= range.FirstPoint.Row &&
               Column >= range.FirstPoint.Column &&
               Column <= range.LastPoint.Column;
    }

    /// <summary>
    /// Is the point within the range or to the left of the range?
    /// </summary>
    internal bool InRangeOrToLeft(in XLSheetRange range)
    {
        return Column >= range.FirstPoint.Column &&
               Row >= range.FirstPoint.Row &&
               Row <= range.LastPoint.Row;
    }

    /// <summary>
    /// Return a new point that has its row coordinate shifted by <paramref name="rowShift"/>.
    /// </summary>
    /// <param name="rowShift">How many rows will new point be shifted. Positive - new point
    ///     is downwards, negative - new point is upwards relative to the current point.</param>
    /// <returns>Shifted point.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal XLSheetPoint ShiftRow(int rowShift)
    {
        return new XLSheetPoint(Row + rowShift, Column);
    }

    /// <summary>
    /// Return a new point that has its column coordinate shifted by <paramref name="columnShift"/>.
    /// </summary>
    /// <param name="columnShift">How many columns will new point be shifted. Positive - new
    ///     point is to the right, negative - new point is to the left.</param>
    /// <returns>Shifted point.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal XLSheetPoint ShiftColumn(int columnShift)
    {
        return new XLSheetPoint(Row, Column + columnShift);
    }
}
