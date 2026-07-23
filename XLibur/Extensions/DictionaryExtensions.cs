using System;
using System.Collections.Generic;

namespace XLibur.Extensions;

internal static class DictionaryExtensions
{
    /// <summary>
    /// Removes every entry whose value matches <paramref name="predicate"/>.
    /// </summary>
    public static void RemoveAll<TKey, TValue>(this Dictionary<TKey, TValue> dic,
        Func<TValue, bool> predicate)
        where TKey : notnull
    {
        // Matching keys must be buffered first: the dictionary cannot be mutated
        // while it is being enumerated. The buffer is allocated lazily so the
        // common "nothing matches" case stays allocation-free.
        List<TKey>? keysToRemove = null;
        foreach (var pair in dic)
        {
            if (predicate(pair.Value))
            {
                (keysToRemove ??= new List<TKey>()).Add(pair.Key);
            }
        }

        if (keysToRemove is null)
        {
            return;
        }

        foreach (var key in keysToRemove)
        {
            dic.Remove(key);
        }
    }
}
