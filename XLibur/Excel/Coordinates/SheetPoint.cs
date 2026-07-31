using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;

namespace XLibur.Excel.Coordinates;

/// <summary>
/// A single point in a workbook. The book point might point to a deleted
/// worksheet, so it might be invalid. Make sure it is checked when
/// determining the properties of the actual data of the point.
/// </summary>
/// <remarks>
/// Packed into a single <c>ulong</c> as <c>[sheetId:16][row0:20][column0:14]</c>
/// (0-based row/column, 50 bits used). This reduces size from 12 to 8 bytes
/// and gives single-instruction equality, hashing, and comparison.
/// </remarks>
[DebuggerDisplay("[{SheetId}] R{Row}C{Column}")]
internal readonly struct SheetPoint : IEquatable<SheetPoint>
{
    private const int PointBits = Point.ColumnBits + 20; // 34
    private const ulong PointMask = (1UL << PointBits) - 1;

    /// <summary>
    /// Packed representation: bits 0-33 = sheet point (0-based row/column),
    /// bits 34-49 = sheetId.
    /// </summary>
    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private readonly ulong _value;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal SheetPoint(XLWorksheet sheet, Point point)
        : this(sheet.SheetId, point)
    {
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal SheetPoint(uint sheetId, Point point)
    {
        _value = ((ulong)sheetId << PointBits) | point.PackedValue;
    }

    internal SheetPoint(uint sheetId, int row, int column)
        : this(sheetId, new Point(row, column))
    {
    }

    /// <summary>
    /// A sheet id of a point. Id of a sheet never changes during workbook
    /// lifecycle (<see cref="XLWorksheet.SheetId"/>), but the sheet may be
    /// deleted, making the sheetId and thus book point invalid.
    /// </summary>
    public uint SheetId
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => (uint)(_value >> PointBits);
    }

    /// <inheritdoc cref="Point.Row"/>
    public int Row
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => Point.Row;
    }

    /// <inheritdoc cref="Point.Column"/>
    public int Column
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => Point.Column;
    }

    /// <summary>
    /// A point in the sheet (without the sheet identifier).
    /// </summary>
    public Point Point
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get
        {
            var pointPacked = _value & PointMask;
            return Unsafe.As<ulong, Point>(ref pointPacked);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool operator ==(SheetPoint lhs, SheetPoint rhs) => lhs._value == rhs._value;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static bool operator !=(SheetPoint lhs, SheetPoint rhs) => lhs._value != rhs._value;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool Equals(SheetPoint other)
    {
        return _value == other._value;
    }

    public override bool Equals(object? obj)
    {
        return obj is SheetPoint other && Equals(other);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override int GetHashCode()
    {
        return _value.GetHashCode();
    }

    public override string ToString()
    {
        return $"[{SheetId}]{Point}";
    }
}
