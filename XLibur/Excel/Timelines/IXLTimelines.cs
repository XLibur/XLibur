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
}
