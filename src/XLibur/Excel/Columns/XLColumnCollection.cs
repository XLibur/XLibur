using System.Collections.Generic;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel;

/// <summary>The materialised columns of a worksheet, keyed by column number.</summary>
internal sealed class XLColumnsCollection : XLLineCollection<XLColumn, ColumnAxis>
{
    /// <summary>
    /// The style the load read from each <c>&lt;cellXfs&gt;</c> index that a <c>&lt;col&gt;</c> of the
    /// part named, or null when it named none.
    /// </summary>
    private Dictionary<uint, XLStyleValue>? _loadedStyles;

    /// <inheritdoc cref="XLLineCollection{TLine, TAxis}.ShiftLines"/>
    public void ShiftColumnsRight(int startingColumn, int columnsToShift) => ShiftLines(startingColumn, columnsToShift);

    /// <summary>
    /// Records that the load read <paramref name="style"/> from the <c>&lt;cellXfs&gt;</c> index
    /// <paramref name="styleIndex"/>, which a <c>&lt;col&gt;</c> of the part named.
    /// </summary>
    internal void RecordLoadedStyle(uint styleIndex, XLStyleValue style)
    {
        _loadedStyles ??= new Dictionary<uint, XLStyleValue>();
        _loadedStyles[styleIndex] = style;
    }

    /// <summary>
    /// Is <paramref name="style"/> the style the load read from <paramref name="styleIndex"/>? A save
    /// writes a loaded column's index back while the column has the style it was loaded with, though
    /// the save would find that style first at another index (#722). The styles part of a loaded
    /// workbook is only appended to, so the index still names that style.
    /// </summary>
    internal bool IsLoadedStyle(uint styleIndex, XLStyleValue style) =>
        _loadedStyles is not null && _loadedStyles.TryGetValue(styleIndex, out var loaded) && loaded.Equals(style);
}
