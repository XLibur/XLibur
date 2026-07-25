using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;

namespace XLibur.Excel.Caching;

internal abstract class XLRepositoryBase : IXLRepository
{
    public abstract void Clear();
}

internal class XLRepositoryBase<Tkey, Tvalue> : XLRepositoryBase, IXLRepository<Tkey, Tvalue>
    where Tkey : struct, IEquatable<Tkey>
    where Tvalue : class
{
    const int CONCURRENCY_LEVEL = 4;
    const int INITIAL_CAPACITY = 1000;

    /// <summary>
    /// Number of lookups that found a collected entry before an opportunistic prune is triggered.
    /// The repositories are process-wide and hold one entry per distinct key ever seen, so without
    /// pruning a long-lived process that builds many workbooks accumulates dead shells forever.
    /// Sweeping is O(n), hence the counter rather than sweeping on every miss.
    /// </summary>
    private const int DeadEntriesBeforePrune = 512;

    private readonly ConcurrentDictionary<Tkey, WeakReference<Tvalue>> _storage;
    private readonly Func<Tkey, Tvalue> _createNew;
    private int _deadEntriesSeen;

    internal XLRepositoryBase(Func<Tkey, Tvalue> createNew)
        : this(createNew, EqualityComparer<Tkey>.Default)
    {
    }

    internal XLRepositoryBase(Func<Tkey, Tvalue> createNew, IEqualityComparer<Tkey> comparer)
    {
        _storage = new ConcurrentDictionary<Tkey, WeakReference<Tvalue>>(CONCURRENCY_LEVEL, INITIAL_CAPACITY, comparer);
        _createNew = createNew;
    }

    /// <summary>
    /// Check if the specified key is presented in the repository.
    /// </summary>
    /// <param name="key">Key to look for.</param>
    /// <param name="value">Value from the repository stored under the specified key or null if the key does
    /// not exist or the entry under this key has already been GCed.</param>
    /// <returns>True if entry exists and alive, false otherwise.</returns>
    public bool ContainsKey(ref Tkey key, out Tvalue? value)
    {
        if (_storage.TryGetValue(key, out var cachedReference))
        {
            if (cachedReference.TryGetTarget(out var target))
            {
                value = target;
                return true;
            }

            NoteDeadEntry();
        }

        value = null;
        return false;
    }

    /// <summary>
    /// Put the entity into the repository under the specified key if no other entity with
    /// the same key is presented.
    /// </summary>
    /// <param name="key">Key to identify the entity.</param>
    /// <param name="value">Entity to store.</param>
    /// <returns>Entity that is stored in the repository under the specified key
    /// (it can be either the <paramref name="value"/> or another entity that has been added to
    /// the repository before.)</returns>
    public Tvalue? Store(ref Tkey key, Tvalue value)
    {
        if (value is null)
            return null;

        while (true)
        {
            if (_storage.TryGetValue(key, out var cachedReference))
            {
                if (cachedReference.TryGetTarget(out var storedValue))
                    return storedValue;

                // The entry is a collected shell. Swap it atomically so that a racing reviver or
                // a concurrent prune cannot make two callers walk away with different instances
                // for the same key: whoever loses the swap loops and picks up the winner's value.
                if (_storage.TryUpdate(key, new WeakReference<Tvalue>(value), cachedReference))
                    return value;

                continue;
            }

            if (_storage.TryAdd(key, new WeakReference<Tvalue>(value)))
                return value;
        }
    }

    public Tvalue GetOrCreate(ref Tkey key)
    {
        if (_storage.TryGetValue(key, out var cachedReference))
        {
            if (cachedReference.TryGetTarget(out var storedValue))
                return storedValue;

            NoteDeadEntry();
        }

        var value = _createNew(key);
        return Store(ref key, value)!;
    }

    public Tvalue? Replace(ref Tkey oldKey, ref Tkey newKey)
    {
        if (_storage.TryRemove(oldKey, out var cachedReference))
        {
            _storage.TryAdd(newKey, cachedReference);
            return GetOrCreate(ref newKey);
        }

        return null;
    }

    public void Remove(ref Tkey key)
    {
        _storage.TryRemove(key, out _);
    }

    public override void Clear()
    {
        _storage.Clear();
        _deadEntriesSeen = 0;
    }

    /// <summary>
    /// List items in the repository removing "dead" entries.
    /// </summary>
    public IEnumerator<Tvalue> GetEnumerator()
    {
        foreach (var pair in _storage)
        {
            if (pair.Value.TryGetTarget(out var value))
                yield return value;
            else
                RemoveIfDead(pair.Key, pair.Value);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    /// <summary>
    /// Count a lookup that hit a collected entry and sweep the whole repository once enough of
    /// them have accumulated.
    /// </summary>
    private void NoteDeadEntry()
    {
        if (Interlocked.Increment(ref _deadEntriesSeen) < DeadEntriesBeforePrune)
            return;

        Interlocked.Exchange(ref _deadEntriesSeen, 0);
        Prune();
    }

    private void Prune()
    {
        foreach (var pair in _storage)
        {
            if (!pair.Value.TryGetTarget(out _))
                RemoveIfDead(pair.Key, pair.Value);
        }
    }

    /// <summary>
    /// Remove an entry only if it is still the same collected shell we observed, so a value
    /// stored concurrently under that key is never dropped.
    /// </summary>
    private void RemoveIfDead(Tkey key, WeakReference<Tvalue> observed)
    {
        if (observed.TryGetTarget(out _))
            return;

        ((ICollection<KeyValuePair<Tkey, WeakReference<Tvalue>>>)_storage)
            .Remove(new KeyValuePair<Tkey, WeakReference<Tvalue>>(key, observed));
    }
}
