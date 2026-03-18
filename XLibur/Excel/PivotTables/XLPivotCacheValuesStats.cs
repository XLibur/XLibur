using System;

namespace XLibur.Excel;

/// <summary>
/// Statistics about a <see cref="XLPivotCacheValues">pivot cache field
/// value</see>. These statistics are available, even if the cache field
/// doesn't have any record values.
/// </summary>
internal readonly struct XLPivotCacheValuesStats
{
    internal bool ContainsBlank { get; init; }

    internal bool ContainsNumber { get; init; }

    /// <summary>
    /// Are all numbers in the field integers? Doesn't have to fit into int32/64, just no fractions.
    /// </summary>
    internal bool ContainsInteger { get; init; }

    internal double? MinValue { get; init; }

    internal double? MaxValue { get; init; }

    /// <summary>
    /// Does the field contain any string, boolean, or error?
    /// </summary>
    internal bool ContainsString { get; init; }

    /// <summary>
    /// Is any text longer than 255 chars?
    /// </summary>
    internal bool LongText { get; init; }

    /// <summary>
    /// Is any value <c>DateTime</c> or <c>TimeSpan</c>? TimeSpan is converted to <em>1899-12-31TXX:XX:XX</em> date.
    /// </summary>
    internal bool ContainsDate { get; init; }

    internal DateTime? MinDate { get; init; }

    internal DateTime? MaxDate { get; init; }
}
