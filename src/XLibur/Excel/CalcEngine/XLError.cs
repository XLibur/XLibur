namespace XLibur.Excel;

/// <summary>
/// A formula error.
/// </summary>
/// <remarks>
/// <para>
/// A member's value is the <c>errorType</c> that [MS-XLSX] 2.3.6.1.3 gives the error's <c>_error</c>
/// rich value, which is one less than the number <c>ERROR.TYPE</c> returns for it. Keep it that way:
/// errors are compared by value in some places (e.g. AutoFilter), and being off by one lets
/// <c>default</c> be a valid error.
/// </para>
/// <para>
/// The spec leaves gaps: 15 and 16 name no error, and 17 is a second form of <c>#BUSY!</c>.
/// <see cref="Python"/> is the one member outside the spec's numbering, because the spec gives
/// <c>#PYTHON!</c> the same <c>errorType</c> as <c>#EXTERNAL!</c>.
/// </para>
/// <para>
/// The members from <see cref="GettingData"/> on, apart from <see cref="SpillRange"/>, come from Excel
/// features XLibur doesn't have, such as linked data types and Python in Excel. XLibur reads, evaluates
/// and writes them like any other error, but never produces one itself.
/// </para>
/// </remarks>
public enum XLError
{
    /// <summary>
    /// <c>#NULL!</c> - Intended to indicate when two areas are required to intersect, but do not.
    /// </summary>
    /// <remarks>The space is an intersection operator.</remarks>
    /// <example><c>SUM(B1 C1)</c> tries to intersect <c>B1:B1</c> area and <c>C1:C1</c> area, but since there are no intersecting cells, the result is <c>#NULL</c>.</example>
    NullValue = 0,

    /// <summary>
    /// <c>#DIV/0!</c> - Intended to indicate when any number (including zero) or any error code is divided by zero.
    /// </summary>
    DivisionByZero = 1,

    /// <summary>
    /// <c>#VALUE!</c> - Intended to indicate when an incompatible type argument is passed to a function, or an incompatible type operand is used with an operator.
    /// </summary>
    /// <example>Passing a non-number text to a function that requires a number, trying to get an area from non-contiguous reference. Creating an area from different sheets <c>Sheet1!A1:Sheet2!A2</c></example>
    IncompatibleValue = 2,

    /// <summary>
    /// <c>#REF!</c> - a formula refers to a cell that's not valid.
    /// </summary>
    /// <example>When unable to find a sheet or a cell.</example>
    CellReference = 3,

    /// <summary>
    /// <c>#NAME?</c> - Intended to indicate when what looks like a name is used, but no such name has been defined.
    /// </summary>
    /// <remarks>Only for named ranges, not sheets.</remarks>
    /// <example><c>TestRange*10</c> when the named range doesn't exist will result in an error.</example>
    NameNotRecognized = 4,

    /// <summary>
    /// <c>#NUM!</c> - Intended to indicate when an argument to a function has a compatible type, but has a value that is outside the domain over which that function is defined.
    /// </summary>
    /// <remarks>This is known as a domain error.</remarks>
    /// <example>ASIN(10) - the ASIN accepts only argument -1..1 (an output of SIN), so the resulting value is <c>#NUM!</c>.</example>
    NumberInvalid = 5,

    /// <summary>
    /// <c>#N/A</c> - Intended to indicate when a designated value is not available.
    /// </summary>
    /// <example>The value is used for extra cells of an array formula that is applied on an array of a smaller size that the array formula.</example>
    NoValueAvailable = 6,

    /// <summary>
    /// <c>#GETTING_DATA</c> - the value is still being retrieved, e.g. by a cube function from an
    /// OLAP data source.
    /// </summary>
    GettingData = 7,

    /// <summary>
    /// <c>#SPILL!</c> - a dynamic array formula's result can't be written because the
    /// spill range isn't empty (blocked by other content) or would fall outside the sheet.
    /// </summary>
    SpillRange = 8,

    /// <summary>
    /// <c>#CONNECT!</c> - an attempt to connect to a service the formula needs failed.
    /// </summary>
    Connect = 9,

    /// <summary>
    /// <c>#BLOCKED!</c> - the connection to a service the formula needs was blocked.
    /// </summary>
    Blocked = 10,

    /// <summary>
    /// <c>#UNKNOWN!</c> - the value has a data type this version of Excel does not support.
    /// </summary>
    Unknown = 11,

    /// <summary>
    /// <c>#FIELD!</c> - a formula refers to a field that a value, such as a linked data type, does not have.
    /// </summary>
    Field = 12,

    /// <summary>
    /// <c>#CALC!</c> - the calculation engine met a case it does not support, e.g. an array
    /// function whose result would be empty.
    /// </summary>
    Calc = 13,

    /// <summary>
    /// <c>#BUSY!</c> - the formula is waiting on data from a service.
    /// </summary>
    Busy = 14,

    /// <summary>
    /// <c>#EXTERNAL!</c> - an external code service the formula depends on returned an error.
    /// </summary>
    External = 18,

    /// <summary>
    /// <c>#TIMEOUT!</c> - the formula ran for longer than the time it is allowed.
    /// </summary>
    Timeout = 19,

    /// <summary>
    /// <c>#PYTHON!</c> - the Python code in the formula returned an error.
    /// </summary>
    /// <remarks>
    /// [MS-XLSX] records this as the <c>#EXTERNAL!</c> error of the Python service, so it has no
    /// <c>errorType</c> of its own. It has a member of its own so that its text survives a round trip,
    /// and <c>ERROR.TYPE</c> returns the same number for it as for <see cref="External"/>.
    /// </remarks>
    Python = 20
}
