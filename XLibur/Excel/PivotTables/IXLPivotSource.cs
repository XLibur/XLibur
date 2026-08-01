using System;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>
/// An abstraction of source data for a <see cref="XLPivotCache"/>. Implementations must correctly
/// implement equals.
/// </summary>
internal interface IXLPivotSource : IEquatable<IXLPivotSource>
{
    /// <summary>
    /// Which kind of source this is. Surfaced publicly as
    /// <see cref="IXLPivotCache.SourceKind"/>, so that a caller can tell a source XLibur cannot
    /// read from one that it can read but that no longer resolves — both of which have no
    /// worksheet, for entirely different reasons.
    /// </summary>
    XLPivotSourceKind Kind { get; }

    /// <summary>
    /// Try to determine actual area of the source reference in the
    /// workbook. Source reference might not be valid in the workbook, some might
    /// not be supported.
    /// </summary>
    bool TryGetSource(XLWorkbook workbook, out XLWorksheet? sheet, out Area? sheetArea);
}
