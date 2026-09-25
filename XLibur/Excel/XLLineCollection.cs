using System;
using System.Buffers;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using XLibur.Excel.Coordinates;
using XLibur.Extensions;

namespace XLibur.Excel;

/// <summary>
/// The materialised rows or columns of a worksheet, keyed by their number on
/// <typeparamref name="TAxis"/>. <see cref="Rows.XLRowsCollection"/> and
/// <see cref="XLColumnsCollection"/> bind it to one axis each.
/// </summary>
/// <typeparam name="TLine"><see cref="Rows.XLRow"/> or <see cref="XLColumn"/>.</typeparam>
/// <typeparam name="TAxis">The axis the keys count along. A struct type argument, so the axis calls
/// in <see cref="ShiftLines"/> are specialised per axis rather than dispatched.</typeparam>
internal abstract class XLLineCollection<TLine, TAxis> : IDictionary<int, TLine>
    where TLine : XLRangeBase
    where TAxis : struct, IGridAxis
{
    private readonly Dictionary<int, TLine> _dictionary = new();

    /// <summary>
    /// Renumbers every materialised line at or after <paramref name="startingIndex"/> by
    /// <paramref name="shift"/>. Lines pushed past the last line of the sheet are dropped.
    /// </summary>
    /// <remarks>
    /// Every affected line is detached before any of them is renumbered, which is what lets the sort
    /// go. The earlier implementations walked the keys in descending order so that each line's
    /// destination was guaranteed free, and paid for that ordering with a LINQ chain and an
    /// O(n log n) sort over a materialised key list — on every single-line insert, not once per
    /// batch. Emptying the affected keys first makes every destination free by construction, so the
    /// order lines are re-added in stops mattering. Inserting one line at a time into a sheet of n
    /// lines is still O(n) per insert, since each line after it genuinely has to be renumbered; this
    /// removes the sort and the allocation on top of it.
    /// </remarks>
    public void ShiftLines(int startingIndex, int shift)
    {
        if (_dictionary.Count == 0)
            return;

        var moving = ArrayPool<KeyValuePair<int, TLine>>.Shared.Rent(_dictionary.Count);
        try
        {
            var count = 0;
            foreach (var pair in _dictionary)
            {
                if (pair.Key >= startingIndex)
                    moving[count++] = pair;
            }

            for (var i = 0; i < count; i++)
                _dictionary.Remove(moving[i].Key);

            var maxIndex = default(TAxis).MaxIndex;
            for (var i = 0; i < count; i++)
            {
                var newIndex = moving[i].Key + shift;
                if (newIndex > maxIndex)
                    continue;

                var line = moving[i].Value;
                default(TAxis).SetLineNumber(line, newIndex);
                _dictionary.Add(newIndex, line);
            }
        }
        finally
        {
            // Cleared on return: the pool hands the buffer on, and stale line references in it would
            // keep whole worksheets alive.
            ArrayPool<KeyValuePair<int, TLine>>.Shared.Return(moving, clearArray: true);
        }
    }

    public void RemoveAll(Func<TLine, bool> predicate) => _dictionary.RemoveAll(predicate);

    #region IDictionary<int, TLine> Members

    public void Add(int key, TLine value) => _dictionary.Add(key, value);

    public bool ContainsKey(int key) => _dictionary.ContainsKey(key);

    public ICollection<int> Keys => _dictionary.Keys;

    public bool Remove(int key) => _dictionary.Remove(key);

    public bool TryGetValue(int key, [MaybeNullWhen(false)] out TLine value)
        => _dictionary.TryGetValue(key, out value);

    public ICollection<TLine> Values => _dictionary.Values;

    public TLine this[int key]
    {
        get => _dictionary[key];
        set => _dictionary[key] = value;
    }

    public void Add(KeyValuePair<int, TLine> item) => _dictionary.Add(item.Key, item.Value);

    public void Clear() => _dictionary.Clear();

    public bool Contains(KeyValuePair<int, TLine> item)
        => ((ICollection<KeyValuePair<int, TLine>>)_dictionary).Contains(item);

    public void CopyTo(KeyValuePair<int, TLine>[] array, int arrayIndex)
    {
        throw new NotImplementedException();
    }

    public int Count => _dictionary.Count;

    public bool IsReadOnly => false;

    public bool Remove(KeyValuePair<int, TLine> item) => _dictionary.Remove(item.Key);

    public IEnumerator<KeyValuePair<int, TLine>> GetEnumerator() => _dictionary.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => _dictionary.GetEnumerator();

    #endregion IDictionary<int, TLine> Members
}
