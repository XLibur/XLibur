using System;

namespace XLibur.Excel;

/// <summary>
/// Which of a timeline's properties the caller has actually assigned.
/// </summary>
/// <remarks>
/// The same device <see cref="XLSlicerFormat"/> uses, and for the same reason. A timeline read from
/// a file is never regenerated — that is what keeps the parts of its XML XLibur has no model for
/// intact — so an edit has to be patched into the element the reader saw. Seeding a value while
/// reading leaves these flags clear, which is how a timeline nobody touched is left alone entirely.
/// </remarks>
[Flags]
internal enum XLTimelineFormat
{
    None = 0,
    Caption = 1 << 0,
    ShowHeader = 1 << 1,
    ShowSelectionLabel = 1 << 2,
    ShowTimeLevel = 1 << 3,
    ShowHorizontalScrollbar = 1 << 4,
    Style = 1 << 5,
    Level = 1 << 6,

    /// <summary>
    /// The timeline has been moved. Unlike the others this is patched into the drawing part rather
    /// than the timelines part, because that is where a timeline's anchor lives.
    /// </summary>
    Position = 1 << 7,
}
