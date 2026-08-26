using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel;

/// <summary>
/// Tracks outline level counts for rows and columns in a worksheet.
/// </summary>
internal sealed class XLOutlineTracker
{
    private readonly Dictionary<int, int> _columnOutlineCount = new();
    private readonly Dictionary<int, int> _rowOutlineCount = new();

    public void IncrementColumnOutline(int level)
    {
        if (level <= 0) return;
        if (_columnOutlineCount.TryGetValue(level, out var value))
            _columnOutlineCount[level] = value + 1;
        else
            _columnOutlineCount.Add(level, 1);
    }

    public void DecrementColumnOutline(int level)
    {
        if (level <= 0) return;
        if (_columnOutlineCount.TryGetValue(level, out var value))
        {
            if (value > 0)
                _columnOutlineCount[level] = value - 1;
        }
        else
            _columnOutlineCount.Add(level, 0);
    }

    public int GetMaxColumnOutline()
    {
        var list = _columnOutlineCount.Where(kp => kp.Value > 0).ToList();
        return list.Count == 0 ? 0 : list.Max(kp => kp.Key);
    }

    public void IncrementRowOutline(int level)
    {
        if (level <= 0) return;
        if (_rowOutlineCount.TryGetValue(level, out var value))
            _rowOutlineCount[level] = value + 1;
        else
            _rowOutlineCount.Add(level, 1);
    }

    public void DecrementRowOutline(int level)
    {
        if (level <= 0) return;
        if (_rowOutlineCount.TryGetValue(level, out var value))
        {
            if (value > 0)
                _rowOutlineCount[level] = value - 1;
        }
        else
            _rowOutlineCount.Add(level, 0);
    }

    /// <summary>
    /// The filtered sequence is what needs the emptiness guard, not the dictionary: Decrement leaves
    /// a zero count behind, so a sheet whose rows were all ungrouped holds a non-empty dictionary of
    /// zeroes and Max() would throw on an empty sequence.
    /// </summary>
    public int GetMaxRowOutline()
    {
        var list = _rowOutlineCount.Where(kp => kp.Value > 0).ToList();
        return list.Count == 0 ? 0 : list.Max(kp => kp.Key);
    }
}
