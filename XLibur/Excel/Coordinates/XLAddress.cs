using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.CompilerServices;
using XLibur.Extensions;

namespace XLibur.Excel;

/// <summary>
/// Address of a single cell, with an optional worksheet and absolute/relative flags.
/// </summary>
/// <remarks>
/// <para>
/// This is what <see cref="IXLCell.Address"/>, <see cref="IXLRangeAddress.FirstAddress"/> and
/// <see cref="IXLRangeAddress.LastAddress"/> return. It is a struct, so reading its members
/// allocates nothing. Assigning it to an <see cref="IXLAddress"/> boxes a copy.
/// </para>
/// <para>
/// Two addresses are equal (<c>==</c>, <see cref="Equals(XLAddress)"/>) when their row, column
/// and <c>$</c> flags are equal. The worksheet is not compared, so <c>A1</c> on one sheet equals
/// <c>A1</c> on another.
/// </para>
/// <para>
/// Row, column, fixedRow and fixedColumn are packed into a single <c>ulong</c> to eliminate
/// alignment padding. Layout of <see cref="_packed"/>:
/// </para>
/// <list type="bullet">
///   <item>bits  0-14: column stored as (column + 1) so that -1 maps to 0 (15 bits)</item>
///   <item>bits 15-35: row stored as (row + 1) so that -1 maps to 0 (21 bits)</item>
///   <item>bit     36: fixedRow flag</item>
///   <item>bit     37: fixedColumn flag</item>
/// </list>
/// </remarks>
public readonly struct XLAddress : IXLAddress, IEquatable<XLAddress>
{
    private const string InvalidRef = "#REF!";
    private const int ColumnBits = 15;  // 15 bits: max stored value 16385 (16384+1 offset)
    private const int RowBits = 21;    // 21 bits: max stored value 1048577 (1048576+1 offset)
    private const int FixedRowBit = ColumnBits + RowBits;       // 36
    private const int FixedColumnBit = FixedRowBit + 1;          // 37
    private const ulong ColumnMask = (1UL << ColumnBits) - 1;   // 0x7FFF
    private const ulong RowMask = (1UL << RowBits) - 1;         // 0x1FFFFF

    #region Static
    /// <summary>
    /// Create an address without a worksheet. For calculation only!
    /// </summary>
    /// <param name="cellAddressString"></param>
    internal static XLAddress Create(string cellAddressString)
    {
        return Create(null, cellAddressString);
    }

    internal static XLAddress Create(XLWorksheet? worksheet, string cellAddressString)
    {
        var fixedColumn = cellAddressString[0] == '$';
        var startPos = fixedColumn ? 1 : 0;

        var rowPos = startPos;
        while (cellAddressString[rowPos] > '9')
        {
            rowPos++;
        }

        var fixedRow = cellAddressString[rowPos] == '$';

        // Column letters occupy [startPos, rowPos); decode them without allocating a substring.
        var columnNumber = XLHelper.GetColumnNumberFromLetter(cellAddressString.AsSpan(startPos, rowPos - startPos));
        var rowStart = fixedRow ? rowPos + 1 : rowPos;
        var rowNumber = int.Parse(cellAddressString.AsSpan(rowStart), XLHelper.NumberStyle, XLHelper.ParseCulture);

        return new XLAddress(worksheet, rowNumber, columnNumber, fixedRow, fixedColumn);
    }

    #endregion Static

    #region Private fields

    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private readonly ulong _packed;

    #endregion Private fields

    #region Constructors

    /// <summary>
    /// Initializes a new <see cref = "XLAddress" /> struct using a mixed notation.  Attention: without worksheet for calculation only!
    /// </summary>
    /// <param name = "rowNumber">The row number of the cell address.</param>
    /// <param name = "columnLetter">The column letter of the cell address.</param>
    /// <param name = "fixedRow"></param>
    /// <param name = "fixedColumn"></param>
    internal XLAddress(int rowNumber, string columnLetter, bool fixedRow, bool fixedColumn)
        : this(null, rowNumber, columnLetter, fixedRow, fixedColumn)
    {
    }

    /// <summary>
    /// Initializes a new <see cref = "XLAddress" /> struct using a mixed notation.
    /// </summary>
    /// <param name = "worksheet"></param>
    /// <param name = "rowNumber">The row number of the cell address.</param>
    /// <param name = "columnLetter">The column letter of the cell address.</param>
    /// <param name = "fixedRow"></param>
    /// <param name = "fixedColumn"></param>
    internal XLAddress(XLWorksheet? worksheet, int rowNumber, string columnLetter, bool fixedRow, bool fixedColumn)
        : this(worksheet, rowNumber, XLHelper.GetColumnNumberFromLetter(columnLetter), fixedRow, fixedColumn)
    {
    }

    /// <summary>
    /// Initializes a new <see cref = "XLAddress" /> struct using R1C1 notation. Attention: without worksheet for calculation only!
    /// </summary>
    /// <param name = "rowNumber">The row number of the cell address.</param>
    /// <param name = "columnNumber">The column number of the cell address.</param>
    /// <param name = "fixedRow"></param>
    /// <param name = "fixedColumn"></param>
    internal XLAddress(int rowNumber, int columnNumber, bool fixedRow, bool fixedColumn)
        : this(null, rowNumber, columnNumber, fixedRow, fixedColumn)
    {
    }

    /// <summary>
    /// Initializes a new <see cref = "XLAddress" /> struct using R1C1 notation.
    /// </summary>
    /// <param name = "worksheet"></param>
    /// <param name = "rowNumber">The row number of the cell address.</param>
    /// <param name = "columnNumber">The column number of the cell address.</param>
    /// <param name = "fixedRow"></param>
    /// <param name = "fixedColumn"></param>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal XLAddress(XLWorksheet? worksheet, int rowNumber, int columnNumber, bool fixedRow, bool fixedColumn)
    {
        Sheet = worksheet;

        // Store row and column with +1 offset so that -1 (invalid sentinel) maps to 0.
        _packed = ((uint)(columnNumber + 1) & ColumnMask)
                | (((uint)(rowNumber + 1) & RowMask) << ColumnBits)
                | (fixedRow ? 1UL << FixedRowBit : 0UL)
                | (fixedColumn ? 1UL << FixedColumnBit : 0UL);
    }

    #endregion Constructors

    #region Properties

    /// <summary>
    /// The worksheet of the address, as the internal type. Null for an address without a worksheet.
    /// </summary>
    internal XLWorksheet? Sheet { get; }

    /// <inheritdoc/>
    public IXLWorksheet? Worksheet
    {
        [DebuggerStepThrough]
        get => Sheet;
    }

    internal bool HasWorksheet
    {
        [DebuggerStepThrough]
        get => Sheet != null;
    }

    /// <inheritdoc/>
    public bool FixedRow
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => (_packed & (1UL << FixedRowBit)) != 0;
    }

    /// <inheritdoc/>
    public bool FixedColumn
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => (_packed & (1UL << FixedColumnBit)) != 0;
    }

    /// <summary>
    /// Gets the row number of this address.
    /// </summary>
    public int RowNumber
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => (int)((_packed >> ColumnBits) & RowMask) - 1;
    }

    /// <summary>
    /// Gets the column number of this address.
    /// </summary>
    public int ColumnNumber
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get => (int)(_packed & ColumnMask) - 1;
    }

    /// <summary>
    /// Gets the column letter(s) of this address.
    /// </summary>
    public string ColumnLetter => XLHelper.GetColumnLetterFromNumber(ColumnNumber);

    #endregion Properties

    #region Overrides

    public override string ToString()
    {
        if (!IsValid)
            return InvalidRef;

        return FormatA1(FixedColumn, FixedRow);
    }

    public string ToString(XLReferenceStyle referenceStyle)
    {
        return ToString(referenceStyle, false);
    }

    public string ToString(XLReferenceStyle referenceStyle, bool includeSheet)
    {
        string address;
        if (!IsValid)
            address = InvalidRef;
        else if (referenceStyle == XLReferenceStyle.A1)
            address = GetTrimmedAddress();
        else if (referenceStyle == XLReferenceStyle.R1C1
                 || HasWorksheet && Sheet!.Workbook.ReferenceStyle == XLReferenceStyle.R1C1)
            address = "R" + RowNumber.ToInvariantString() + "C" + ColumnNumber.ToInvariantString();
        else
            address = GetTrimmedAddress();

        if (includeSheet)
            return string.Concat(
                WorksheetIsDeleted ? "#REF" : Sheet!.Name.EscapeSheetName(),
                '!',
                address);

        return address;
    }

    #endregion Overrides

    #region Methods

    internal string GetTrimmedAddress()
    {
        return FormatA1(fixedColumn: false, fixedRow: false);
    }

    /// <summary>
    /// Formats the address in A1 style as exactly one string. The struct is readonly and cannot
    /// cache the result, so each call allocates, and it allocates nothing but the result.
    /// </summary>
    private string FormatA1(bool fixedColumn, bool fixedRow)
    {
        var columnLetter = ColumnLetter;

        // Max layout: '$' + 3 column letters + '$' + 7 row digits = 12 chars.
        Span<char> buffer = stackalloc char[16];
        var pos = 0;
        if (fixedColumn)
            buffer[pos++] = '$';

        columnLetter.AsSpan().CopyTo(buffer[pos..]);
        pos += columnLetter.Length;

        if (fixedRow)
            buffer[pos++] = '$';

        RowNumber.TryFormat(buffer[pos..], out var written, provider: CultureInfo.InvariantCulture);
        pos += written;

        return new string(buffer[..pos]);
    }

    #endregion Methods

    #region Operator Overloads

    /// <summary>
    /// Compares row, column and the <c>$</c> flags. The worksheet is not compared.
    /// </summary>
    public static bool operator ==(XLAddress left, XLAddress right)
    {
        return left.Equals(right);
    }

    /// <summary>
    /// Compares row, column and the <c>$</c> flags. The worksheet is not compared.
    /// </summary>
    public static bool operator !=(XLAddress left, XLAddress right)
    {
        return !(left == right);
    }

    #endregion Operator Overloads

    #region Interface Requirements

    #region IEqualityComparer<IXLAddress> Members

    bool IEqualityComparer<IXLAddress>.Equals(IXLAddress? x, IXLAddress? y)
    {
        if (x is null) return y is null;
        return x.Equals(y);
    }

    int IEqualityComparer<IXLAddress>.GetHashCode(IXLAddress obj)
    {
        return ((XLAddress)obj).GetHashCode();
    }

    #endregion IEqualityComparer<IXLAddress> Members

    #region IEquatable Members

    /// <summary>
    /// Compares row, column and the <c>$</c> flags. The worksheet is not compared.
    /// </summary>
    public bool Equals(IXLAddress? other)
    {
        if (other == null)
            return false;

        return RowNumber == other.RowNumber &&
               ColumnNumber == other.ColumnNumber &&
               FixedRow == other.FixedRow &&
               FixedColumn == other.FixedColumn;
    }

    /// <summary>
    /// Compares row, column and the <c>$</c> flags. The worksheet is not compared.
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public bool Equals(XLAddress other)
    {
        return _packed == other._packed;
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as IXLAddress);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override int GetHashCode()
    {
        return _packed.GetHashCode();
    }

    #endregion IEquatable Members

    #endregion Interface Requirements

    public string ToStringRelative()
    {
        return ToStringRelative(false);
    }

    public string ToStringRelative(bool includeSheet)
    {
        var address = IsValid ? GetTrimmedAddress() : InvalidRef;

        if (includeSheet)
            return string.Concat(
                WorksheetIsDeleted ? "#REF" : Sheet!.Name.EscapeSheetName(),
                '!',
                address
            );

        return address;
    }

    public string ToStringFixed()
    {
        return ToStringFixed(XLReferenceStyle.Default);
    }

    public string ToStringFixed(XLReferenceStyle referenceStyle)
    {
        return ToStringFixed(referenceStyle, false);
    }

    public string ToStringFixed(XLReferenceStyle referenceStyle, bool includeSheet)
    {
        string address;

        if (referenceStyle == XLReferenceStyle.Default && HasWorksheet)
            referenceStyle = Sheet!.Workbook.ReferenceStyle;

        if (referenceStyle == XLReferenceStyle.Default)
            referenceStyle = XLReferenceStyle.A1;

        Debug.Assert(referenceStyle != XLReferenceStyle.Default);

        if (!IsValid)
        {
            address = InvalidRef;
        }
        else
        {
            address = referenceStyle switch
            {
                XLReferenceStyle.A1 => FormatA1(fixedColumn: true, fixedRow: true),
                XLReferenceStyle.R1C1 => string.Concat('R', RowNumber.ToInvariantString(), 'C', ColumnNumber),
                _ => throw new NotImplementedException(),
            };
        }

        if (includeSheet)
            return string.Concat(
                WorksheetIsDeleted ? "#REF" : Sheet!.Name.EscapeSheetName(),
                '!',
                address);

        return address;
    }

    internal XLAddress WithoutWorksheet()
    {
        return new XLAddress(RowNumber, ColumnNumber, FixedRow, FixedColumn);
    }

    internal XLAddress WithWorksheet(XLWorksheet worksheet)
    {
        return new XLAddress(worksheet, RowNumber, ColumnNumber, FixedRow, FixedColumn);
    }

    public string UniqueId => RowNumber.ToString("0000000") + ColumnNumber.ToString("00000");

    internal bool IsValid => RowNumber is > 0 and <= XLHelper.MaxRowNumber &&
                             ColumnNumber is > 0 and <= XLHelper.MaxColumnNumber;

    private bool WorksheetIsDeleted => Sheet?.IsDeleted == true;
}
