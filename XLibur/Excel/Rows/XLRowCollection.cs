using System;
using System.Buffers;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace XLibur.Excel.Rows;

internal sealed class XLRowsCollection : IDictionary<int, XLRow>
{
    private readonly Dictionary<int, XLRow> _dictionary = new();

    private Dictionary<int, XLRow> Deleted { get; } = new();

    private int _maxRowUsed;

    #region IDictionary<int,XLRow> Members

    public void Add(int key, XLRow value)
    {
        if (key > _maxRowUsed) _maxRowUsed = key;

        Deleted.Remove(key);
        _dictionary.Add(key, value);
    }

    public bool ContainsKey(int key)
    {
        return _dictionary.ContainsKey(key);
    }

    public ICollection<int> Keys => _dictionary.Keys;

    public bool Remove(int key)
    {
        if (!Deleted.ContainsKey(key))
            Deleted.Add(key, _dictionary[key]);

        return _dictionary.Remove(key);
    }

    public bool TryGetValue(int key, [MaybeNullWhen(false)] out XLRow value)
    {
        return _dictionary.TryGetValue(key, out value);
    }

    public ICollection<XLRow> Values => _dictionary.Values;

    public XLRow this[int key]
    {
        get => _dictionary[key];
        set => _dictionary[key] = value;
    }

    public void Add(KeyValuePair<int, XLRow> item)
    {
        if (item.Key > _maxRowUsed) _maxRowUsed = item.Key;

        Deleted.Remove(item.Key);
        _dictionary.Add(item.Key, item.Value);
    }

    public void Clear()
    {
        _dictionary.Clear();
    }

    public bool Contains(KeyValuePair<int, XLRow> item)
    {
        return _dictionary.Contains(item);
    }

    public void CopyTo(KeyValuePair<int, XLRow>[] array, int arrayIndex)
    {
        throw new NotImplementedException();
    }

    public int Count => _dictionary.Count;

    public bool IsReadOnly => false;

    public bool Remove(KeyValuePair<int, XLRow> item)
    {
        if (!Deleted.ContainsKey(item.Key))
            Deleted.Add(item.Key, _dictionary[item.Key]);

        return _dictionary.Remove(item.Key);
    }

    public IEnumerator<KeyValuePair<int, XLRow>> GetEnumerator()
    {
        return _dictionary.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return _dictionary.GetEnumerator();
    }

    #endregion IDictionary<int,XLRow> Members

    /// <summary>
    /// Renumbers every materialised row at or below <paramref name="startingRow"/> by
    /// <paramref name="rowsToShift"/>. Rows pushed past the last row of the sheet are dropped.
    /// </summary>
    /// <remarks>
    /// Every affected row is detached before any of them is renumbered, which is what lets the sort go.
    /// The previous implementation walked the keys in descending order so that each row's destination
    /// was guaranteed free, and paid for that ordering with a LINQ chain and an O(n log n) sort over a
    /// materialised key list — on every single-row insert, not once per batch. Emptying the affected
    /// keys first makes every destination free by construction, so the order rows are re-added in stops
    /// mattering. Inserting one row at a time into a sheet of n rows is still O(n) per insert, since
    /// each row below genuinely has to be renumbered; this removes the sort and the allocation on top
    /// of it.
    /// <para>
    /// Rows are added straight to the backing dictionary, not through <see cref="Add(int, XLRow)"/>:
    /// moving a row must not touch <c>_maxRowUsed</c> or clear the deleted-row record, and the previous
    /// implementation did not either.
    /// </para>
    /// </remarks>
    public void ShiftRowsDown(int startingRow, int rowsToShift)
    {
        if (_dictionary.Count == 0)
            return;

        var moving = ArrayPool<KeyValuePair<int, XLRow>>.Shared.Rent(_dictionary.Count);
        try
        {
            var count = 0;
            foreach (var pair in _dictionary)
            {
                if (pair.Key >= startingRow)
                    moving[count++] = pair;
            }

            for (var i = 0; i < count; i++)
                _dictionary.Remove(moving[i].Key);

            for (var i = 0; i < count; i++)
            {
                var newRowNumber = moving[i].Key + rowsToShift;
                if (newRowNumber > XLHelper.MaxRowNumber)
                    continue;

                var row = moving[i].Value;
                row.SetRowNumber(newRowNumber);
                _dictionary.Add(newRowNumber, row);
            }
        }
        finally
        {
            // Cleared on return: the pool hands the buffer on, and stale XLRow references in it would
            // keep whole worksheets alive.
            ArrayPool<KeyValuePair<int, XLRow>>.Shared.Return(moving, clearArray: true);
        }
    }
}
