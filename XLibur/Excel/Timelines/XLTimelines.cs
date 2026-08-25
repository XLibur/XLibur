using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace XLibur.Excel;

internal sealed class XLTimelines : IXLTimelines
{
    private readonly List<XLTimeline> _timelines = [];
    private readonly XLWorksheet _worksheet;

    internal XLTimelines(XLWorksheet worksheet)
    {
        _worksheet = worksheet;
    }

    public int Count => _timelines.Count;

    internal IReadOnlyList<XLTimeline> Items => _timelines;

    public IEnumerator<IXLTimeline> GetEnumerator() => _timelines.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IXLTimeline Timeline(string name)
    {
        if (!TryGetTimeline(name, out var timeline))
            throw new KeyNotFoundException($"The worksheet has no timeline named '{name}'.");

        return timeline;
    }

    public bool TryGetTimeline(string name, [NotNullWhen(true)] out IXLTimeline? timeline)
    {
        foreach (var candidate in _timelines)
        {
            if (XLHelper.NameComparer.Equals(candidate.Name, name))
            {
                timeline = candidate;
                return true;
            }
        }

        timeline = null;
        return false;
    }

    internal void Add(XLTimeline timeline) => _timelines.Add(timeline);

    /// <summary>
    /// Drops a timeline from the worksheet and records what the save path has to unpick.
    /// </summary>
    /// <remarks>
    /// Removing a timeline is not a matter of dropping one element. Its cache part, the workbook's
    /// registration of that cache, the <c>#N/A</c> defined name written for it, the worksheet's
    /// <c>extLst</c> reference and the drawing anchor all have to go with it, or the saved file has
    /// an orphan Excel will offer to repair.
    /// </remarks>
    internal void Remove(XLTimeline timeline)
    {
        if (_timelines.Remove(timeline) && !timeline.IsNew)
            Removed.Add(timeline);
    }

    /// <summary>
    /// Timelines removed since the workbook was loaded, still holding the relationship ids and cache
    /// names the save path needs to clean up after them. Cleared once a save has consumed it.
    /// </summary>
    internal List<XLTimeline> Removed { get; } = [];
}
