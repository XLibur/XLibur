using System;

namespace XLibur.Excel;

/// <summary>
/// Which of a slicer's properties the caller has actually assigned.
/// </summary>
/// <remarks>
/// The same device <see cref="XLChartSeriesFormat"/> uses, and for the same reason. A slicer read
/// from a file is never regenerated — that is what keeps the attributes XLibur has no model for
/// intact — so an edit has to be patched into the element the reader saw. Seeding a value while
/// reading leaves these flags clear, which is how a slicer nobody touched is left alone entirely.
/// </remarks>
[Flags]
internal enum XLSlicerFormat
{
    None = 0,
    Caption = 1 << 0,
    ShowCaption = 1 << 1,
    Style = 1 << 2,
    ColumnCount = 1 << 3,
    RowHeight = 1 << 4,
}
