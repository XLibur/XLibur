using System;
using System.Collections.Generic;

namespace XLibur.Excel;

/// <summary>
/// Reference to a single cell in a workbook. Reference can be absolute, relative or mixed.
/// Reference can be with or without a worksheet.
/// </summary>
public interface IXLAddress : IEqualityComparer<IXLAddress>, IEquatable<IXLAddress>
{
    string ColumnLetter { get; }
    int ColumnNumber { get; }
    bool FixedColumn { get; }
    bool FixedRow { get; }
    int RowNumber { get; }
    string UniqueId { get; }

    /// <summary>
    /// Worksheet of the reference. Value is null for address without a worksheet.
    /// </summary>
    IXLWorksheet? Worksheet { get; }

    string ToString(XLReferenceStyle referenceStyle);

    string ToString(XLReferenceStyle referenceStyle, bool includeSheet);

    string ToStringFixed();

    string ToStringFixed(XLReferenceStyle referenceStyle);

    string ToStringFixed(XLReferenceStyle referenceStyle, bool includeSheet);

    string ToStringRelative();

    string ToStringRelative(bool includeSheet);
}
