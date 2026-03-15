using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using XLibur.Extensions;

namespace XLibur.Excel;

/// <summary>
/// Cell address with optional worksheet, absolute/relative flags, and
/// a cached trimmed-address string. Row, column, fixedRow, and fixedColumn
/// are packed into a single <c>ulong</c> to eliminate alignment padding.
/// </summary>
/// <remarks>
/// Layout of <see cref="_packed"/>:
/// <list type="bullet">
///   <item>bits  0-14: column stored as (column + 1) so that -1 maps to 0 (15 bits)</item>
///   <item>bits 15-35: row stored as (row + 1) so that -1 maps to 0 (21 bits)</item>
///   <item>bit     36: fixedRow flag</item>
///   <item>bit     37: fixedColumn flag</item>
/// </list>
/// </remarks>
internal struct XLAddress : IXLAddress, IEquatable<XLAddress>
{
    private const int ColumnBits = 15;  // 15 bits: max stored value 16385 (16384+1 offset)
    private const int RowBits = 21;    // 21 bits: max stored value 1048577 (1048576+1 offset)
    private const int FixedRowBit = ColumnBits + RowBits;       // 36
    private const int FixedColumnBit = FixedRowBit + 1;          // 37
    private const ulong ColumnMask = (1UL << ColumnBits) - 1;   // 0x7FFF
    private const ulong RowMask = (1UL << RowBits) - 1;         // 0x1FFFFF

    #region Static
    /// <summary>
    /// Create address without worksheet. For calculation only!
    /// </summary>
    /// <param name="cellAddressString"></param>
    public static XLAddress Create(string cellAddressString)
    {
        return Create(null, cellAddressString);
    }

    public static XLAddress Create(XLWorksheet? worksheet, string cellAddressString)
    {
        var fixedColumn = cellAddressString[0] == '$';
        int startPos;
        if (fixedColumn)
        {
            startPos = 1;
        }
        else
        {
            startPos = 0;
        }

        int rowPos = startPos;
        while (cellAddressString[rowPos] > '9')
        {
            rowPos++;
        }

        var fixedRow = cellAddressString[rowPos] == '$';
        string columnLetter;
        int rowNumber;
        if (fixedRow)
        {
            if (fixedColumn)
            {
                columnLetter = cellAddressString.Substring(startPos, rowPos - 1);
            }
            else
            {
                columnLetter = cellAddressString.Substring(startPos, rowPos);
            }

            rowNumber = int.Parse(cellAddressString.Substring(rowPos + 1), XLHelper.NumberStyle, XLHelper.ParseCulture);
        }
        else
        {
            if (fixedColumn)
            {
                columnLetter = cellAddressString.Substring(startPos, rowPos - 1);
            }
            else
            {
                columnLetter = cellAddressString.Substring(startPos, rowPos);
            }

            rowNumber = int.Parse(cellAddressString.Substring(rowPos), XLHelper.NumberStyle, XLHelper.ParseCulture);
        }
        return new XLAddress(worksheet, rowNumber, columnLetter, fixedRow, fixedColumn);
    }

    #endregion Static

    #region Private fields

    [DebuggerBrowsable(DebuggerBrowsableState.Never)]
    private readonly ulong _packed;

    private string? _trimmedAddress;

    #endregion Private fields

    #region Constructors

    /// <summary>
    /// Initializes a new <see cref = "XLAddress" /> struct using a mixed notation.  Attention: without worksheet for calculation only!
    /// </summary>
    /// <param name = "rowNumber">The row number of the cell address.</param>
    /// <param name = "columnLetter">The column letter of the cell address.</param>
    /// <param name = "fixedRow"></param>
    /// <param name = "fixedColumn"></param>
    public XLAddress(int rowNumber, string columnLetter, bool fixedRow, bool fixedColumn)
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
    public XLAddress(XLWorksheet? worksheet, int rowNumber, string columnLetter, bool fixedRow, bool fixedColumn)
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
    public XLAddress(int rowNumber, int columnNumber, bool fixedRow, bool fixedColumn)
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
    public XLAddress(XLWorksheet? worksheet, int rowNumber, int columnNumber, bool fixedRow, bool fixedColumn) : this()
    {
        Worksheet = worksheet;

        // Store row and column with +1 offset so that -1 (invalid sentinel) maps to 0.
        _packed = ((ulong)(uint)(columnNumber + 1) & ColumnMask)
                | (((ulong)(uint)(rowNumber + 1) & RowMask) << ColumnBits)
                | (fixedRow ? 1UL << FixedRowBit : 0UL)
                | (fixedColumn ? 1UL << FixedColumnBit : 0UL);
    }

    #endregion Constructors

    #region Properties

    public XLWorksheet? Worksheet { get; internal set; }

    IXLWorksheet? IXLAddress.Worksheet
    {
        [DebuggerStepThrough]
        get { return Worksheet; }
    }

    public bool HasWorksheet
    {
        [DebuggerStepThrough]
        get { return Worksheet != null; }
    }

    public bool FixedRow
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get { return (_packed & (1UL << FixedRowBit)) != 0; }
    }

    public bool FixedColumn
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get { return (_packed & (1UL << FixedColumnBit)) != 0; }
    }

    /// <summary>
    /// Gets the row number of this address.
    /// </summary>
    public int RowNumber
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get { return (int)((_packed >> ColumnBits) & RowMask) - 1; }
    }

    /// <summary>
    /// Gets the column number of this address.
    /// </summary>
    public int ColumnNumber
    {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        get { return (int)(_packed & ColumnMask) - 1; }
    }

    /// <summary>
    /// Gets the column letter(s) of this address.
    /// </summary>
    public string ColumnLetter
    {
        get { return XLHelper.GetColumnLetterFromNumber(ColumnNumber); }
    }

    #endregion Properties

    #region Overrides

    public override string ToString()
    {
        if (!IsValid)
            return "#REF!";

        string retVal = ColumnLetter;
        if (FixedColumn)
        {
            retVal = "$" + retVal;
        }
        if (FixedRow)
        {
            retVal += "$";
        }
        retVal += RowNumber.ToInvariantString();
        return retVal;
    }

    public string ToString(XLReferenceStyle referenceStyle)
    {
        return ToString(referenceStyle, false);
    }

    public string ToString(XLReferenceStyle referenceStyle, bool includeSheet)
    {
        string address;
        if (!IsValid)
            address = "#REF!";
        else if (referenceStyle == XLReferenceStyle.A1)
            address = GetTrimmedAddress();
        else if (referenceStyle == XLReferenceStyle.R1C1
                 || HasWorksheet && Worksheet!.Workbook.ReferenceStyle == XLReferenceStyle.R1C1)
            address = "R" + RowNumber.ToInvariantString() + "C" + ColumnNumber.ToInvariantString();
        else
            address = GetTrimmedAddress();

        if (includeSheet)
            return string.Concat(
                WorksheetIsDeleted ? "#REF" : Worksheet!.Name.EscapeSheetName(),
                '!',
                address);

        return address;
    }

    #endregion Overrides

    #region Methods

    public string GetTrimmedAddress()
    {
        return _trimmedAddress ??= ColumnLetter + RowNumber.ToInvariantString();
    }

    #endregion Methods

    #region Operator Overloads

    public static XLAddress operator +(XLAddress left, XLAddress right)
    {
        return new XLAddress(left.Worksheet,
            left.RowNumber + right.RowNumber,
            left.ColumnNumber + right.ColumnNumber,
            left.FixedRow,
            left.FixedColumn);
    }

    public static XLAddress operator -(XLAddress left, XLAddress right)
    {
        return new XLAddress(left.Worksheet,
            left.RowNumber - right.RowNumber,
            left.ColumnNumber - right.ColumnNumber,
            left.FixedRow,
            left.FixedColumn);
    }

    public static XLAddress operator +(XLAddress left, int right)
    {
        return new XLAddress(left.Worksheet,
            left.RowNumber + right,
            left.ColumnNumber + right,
            left.FixedRow,
            left.FixedColumn);
    }

    public static XLAddress operator -(XLAddress left, int right)
    {
        return new XLAddress(left.Worksheet,
            left.RowNumber - right,
            left.ColumnNumber - right,
            left.FixedRow,
            left.FixedColumn);
    }

    public static bool operator ==(XLAddress left, XLAddress right)
    {
        return left.Equals(right);
    }

    public static bool operator !=(XLAddress left, XLAddress right)
    {
        return !(left == right);
    }

    #endregion Operator Overloads

    #region Interface Requirements

    #region IEqualityComparer<XLCellAddress> Members

    public bool Equals(IXLAddress? x, IXLAddress? y)
    {
        return x == y;
    }

    public new bool Equals(object? x, object? y)
    {
        return x == y;
    }

    #endregion IEqualityComparer<XLCellAddress> Members

    #region IEquatable<XLCellAddress> Members

    public bool Equals(IXLAddress? other)
    {
        if (other == null)
            return false;

        return RowNumber == other.RowNumber &&
               ColumnNumber == other.ColumnNumber &&
               FixedRow == other.FixedRow &&
               FixedColumn == other.FixedColumn;
    }

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

    public int GetHashCode(IXLAddress obj)
    {
        return ((XLAddress)obj).GetHashCode();
    }

    #endregion IEquatable<XLCellAddress> Members

    #endregion Interface Requirements

    public string ToStringRelative()
    {
        return ToStringRelative(false);
    }

    public string ToStringFixed()
    {
        return ToStringFixed(XLReferenceStyle.Default);
    }

    public string ToStringRelative(bool includeSheet)
    {
        var address = IsValid ? GetTrimmedAddress() : "#REF!";

        if (includeSheet)
            return string.Concat(
                WorksheetIsDeleted ? "#REF" : Worksheet!.Name.EscapeSheetName(),
                '!',
                address
            );

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

    public string ToStringFixed(XLReferenceStyle referenceStyle)
    {
        return ToStringFixed(referenceStyle, false);
    }

    public string ToStringFixed(XLReferenceStyle referenceStyle, bool includeSheet)
    {
        string address;

        if (referenceStyle == XLReferenceStyle.Default && HasWorksheet)
            referenceStyle = Worksheet!.Workbook.ReferenceStyle;

        if (referenceStyle == XLReferenceStyle.Default)
            referenceStyle = XLReferenceStyle.A1;

        Debug.Assert(referenceStyle != XLReferenceStyle.Default);

        if (!IsValid)
        {
            address = "#REF!";
        }
        else
        {
            address = referenceStyle switch
            {
                XLReferenceStyle.A1 => string.Concat('$', ColumnLetter, '$', RowNumber.ToInvariantString()),
                XLReferenceStyle.R1C1 => string.Concat('R', RowNumber.ToInvariantString(), 'C', ColumnNumber),
                _ => throw new NotImplementedException(),
            };
        }

        if (includeSheet)
            return string.Concat(
                WorksheetIsDeleted ? "#REF" : Worksheet!.Name.EscapeSheetName(),
                '!',
                address);

        return address;
    }

    public string UniqueId => RowNumber.ToString("0000000") + ColumnNumber.ToString("00000");

    public bool IsValid => RowNumber is > 0 and <= XLHelper.MaxRowNumber &&
                           ColumnNumber is > 0 and <= XLHelper.MaxColumnNumber;

    private bool WorksheetIsDeleted => Worksheet?.IsDeleted == true;
}
