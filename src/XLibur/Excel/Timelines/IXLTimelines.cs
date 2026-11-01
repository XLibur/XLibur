using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace XLibur.Excel;

/// <summary>
/// The timelines drawn on a worksheet.
/// </summary>
/// <remarks>
/// The worksheet owns its timelines: this is the collection a timeline is added to and removed from.
/// What a timeline filters is a separate relationship held by its cache, exposed as a view on
/// <see cref="IXLPivotTable.Timelines"/>.
/// </remarks>
public interface IXLTimelines : IEnumerable<IXLTimeline>
{
    /// <summary>The number of timelines on the worksheet.</summary>
    int Count { get; }

    /// <summary>The timeline with the given <see cref="IXLTimeline.Name"/>.</summary>
    /// <param name="name">The timeline's internal name, which is not its caption.</param>
    /// <exception cref="System.Collections.Generic.KeyNotFoundException">
    /// The worksheet has no timeline with that name.
    /// </exception>
    IXLTimeline Timeline(string name);

    /// <summary>Finds the timeline with the given <see cref="IXLTimeline.Name"/>.</summary>
    /// <param name="name">The timeline's internal name, which is not its caption.</param>
    /// <param name="timeline">The timeline, when one was found.</param>
    /// <returns>Whether a timeline with that name is on the worksheet.</returns>
    bool TryGetTimeline(string name, [NotNullWhen(true)] out IXLTimeline? timeline);

    /// <summary>
    /// Adds a timeline that filters a pivot table on one of its date fields.
    /// </summary>
    /// <param name="pivotTable">The pivot table to filter. It need not be on this worksheet.</param>
    /// <param name="dateFieldName">The pivot cache field the band is drawn from.</param>
    /// <returns>The new timeline, showing every date and so filtering nothing.</returns>
    /// <exception cref="System.ArgumentException">
    /// The pivot table's cache has no field of that name, or that field holds no dates.
    /// </exception>
    /// <remarks>
    /// The timeline is placed to the right of the pivot table. Use <see cref="IXLTimeline.Position"/>
    /// to move it.
    /// </remarks>
    IXLTimeline Add(IXLPivotTable pivotTable, string dateFieldName);
}
