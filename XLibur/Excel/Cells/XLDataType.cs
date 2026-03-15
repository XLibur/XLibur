namespace XLibur.Excel;

/// <summary>
/// A value that is in the cell.
/// </summary>
public enum XLDataType
{
    /// <summary>
    /// The value is a blank (either blank cells or the omitted optional argument of a function, e.g. <c>IF(TRUE,,)</c>.
    /// </summary>
    /// <remarks>Keep as the first, so the default values are blank.</remarks>
    Blank = 0,

    /// <summary>
    /// The value is a logical value.
    /// </summary>
    Boolean = 1,

    /// <summary>
    /// The value is a double-precision floating points number, excluding <see cref="double.NaN"/>,
    /// <see cref="double.PositiveInfinity"/> or <see cref="double.NegativeInfinity"/>.
    /// </summary>
    Number = 2,

    /// <summary>
    /// A text or a rich text. Can't be <c>null</c> and can be at most 32767 characters long.
    /// </summary>
    Text = 3,

    /// <summary>
    /// The value is one of <see cref="XLError"/>.
    /// </summary>
    Error = 4,

    /// <summary>
    /// The value is a <see cref="DateTime"/>, represented as a serial date time number.
    /// </summary>
    /// <remarks>
    /// Serial date time 60 is a 1900-02-29, nonexistent day kept for compatibility,
    /// but unrepresentable by <c>DateTime</c>. Don't use.
    /// </remarks>
    DateTime = 5,

    /// <summary>
    /// The value is a <see cref="TimeSpan"/>, represented in a serial date time (24 hours is 1, 36 hours is 1.5 ect.).
    /// </summary>
    TimeSpan = 6,
}
