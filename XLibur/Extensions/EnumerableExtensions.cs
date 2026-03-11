
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel;

internal static class EnumerableExtensions
{
    public static void ForEach<T>(this IEnumerable<T> source, Action<T> action)
    {
        foreach (var item in source)
            action(item);
    }

    public static Type? GetItemType(this IEnumerable source)
    {
        return GetGenericArgument(source.GetType());

        Type? GetGenericArgument(Type collectionType)
        {
            var ienumerable = collectionType.GetInterfaces()
                .SingleOrDefault(i => i.GetGenericArguments().Length == 1 &&
                                      i.Name == "IEnumerable`1");

            return ienumerable?.GetGenericArguments().FirstOrDefault();
        }
    }

    extension<T>(IEnumerable<T> source)
    {
        public HashSet<T> ToHashSet()
        {
            return [.. source];
        }

        /// <summary>
        /// Skip the last element of a sequence.
        /// </summary>
        public IEnumerable<T> SkipLast()
        {
            using var enumerator = source.GetEnumerator();
            if (!enumerator.MoveNext())
                yield break;

            var prev = enumerator.Current;
            while (enumerator.MoveNext())
            {
                yield return prev;
                prev = enumerator.Current;
            }
        }

        public bool HasDuplicates()
        {
            HashSet<T> distinctItems = [];
            return source.Any(item => !distinctItems.Add(item));
        }

        /// <summary>
        /// Select all <typeparamref name="TItem"/> that are not null.
        /// </summary>
        public IEnumerable<TItem> WhereNotNull<TItem>(Func<T, TItem?> property)
            where TItem : struct
        {
            return source.Select(property).Where(x => x.HasValue).Select(x => x!.Value);
        }
    }
}
