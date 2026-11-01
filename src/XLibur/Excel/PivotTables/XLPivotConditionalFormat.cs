using System.Collections.Generic;
using XLibur.Excel.ConditionalFormats;

namespace XLibur.Excel;

/// <summary>
/// Specification of conditional formatting of a pivot table.
/// </summary>
internal sealed class XLPivotConditionalFormat
{
    private readonly List<XLPivotArea> _area = new();

    internal XLPivotConditionalFormat(XLConditionalFormat format)
    {
        Format = format;
    }

    /// <summary>
    /// An option to display in GUI on how to update <see cref="Areas"/>.
    /// </summary>
    internal XLPivotCfScope Scope { get; init; } = XLPivotCfScope.SelectedCells;

    /// <summary>
    /// A rule that determines how CF should be applied to <see cref="Areas"/>.
    /// </summary>
    /// <remarks>Doesn't seem to work, Excel has no dialogue, nothing found on web, and Excel tries
    /// to repair on row/column values. Avoid it if possible.</remarks>
    internal XLPivotCfRuleType Type { get; init; } = XLPivotCfRuleType.None;

    /// <summary>
    /// Areas of pivot table the rule should be applied. The areas are projected to the sheet
    /// <see cref="XLConditionalFormat.Ranges"/> that Excel actually uses to display CF.
    /// </summary>
    internal IReadOnlyList<XLPivotArea> Areas => _area;

    /// <summary>
    /// Conditional format applied to the <see cref="Areas"/>.
    /// </summary>
    /// <remarks>
    /// The <see cref="XLConditionalFormat.Priority"/> of the format is used as an identifier used
    /// to connect pivot CF element and sheet CF element. Pivot CF is ultimately part of sheet CFs,
    /// and the priority determines order of CF application (note that CF has
    /// <see cref="XLConditionalFormat.StopIfTrue"/> flag).
    /// </remarks>
    internal XLConditionalFormat Format { get; }

    internal void AddArea(XLPivotArea pivotArea)
    {
        _area.Add(pivotArea);
    }

    /// <summary>
    /// Renumber the 'data' field positions of every area around the removal of the value at
    /// <paramref name="removedPosition"/> (#585, following #577's <see cref="XLPivotTable.
    /// RemoveValueFromFormats"/>), dropping an area left naming no value. The areas are this
    /// format's own (<see cref="XLPivotTable.CopyConditionalFormatsTo"/> clones them for a copy), so
    /// renumbering in place cannot corrupt another table's format.
    /// </summary>
    /// <returns>
    /// <c>true</c> when this format is left with no area at all, meaning it applies to nothing and
    /// has to go.
    /// </returns>
    internal bool RemoveEmptiedAreas(uint removedPosition, uint remainingValueCount)
    {
        var hadAreas = _area.Count > 0;
        _area.RemoveAll(area => XLPivotTable.RenumberDataFieldPositions(area, removedPosition, remainingValueCount));
        return hadAreas && _area.Count == 0;
    }
}
