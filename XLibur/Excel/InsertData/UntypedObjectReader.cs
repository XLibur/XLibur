using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel.InsertData;

internal sealed class UntypedObjectReader : IInsertDataReader
{
    private readonly IEnumerable<IInsertDataReader> _readers;

    public UntypedObjectReader(IEnumerable data)
    {
        var data1 = (data ?? Array.Empty<object>()).Cast<object>();
        _readers = CreateReaders(data1).ToList();
    }

    /// <summary>
    /// Splits <paramref name="data"/> into runs of items of the same type and yields one reader
    /// per run.
    /// </summary>
    private static IEnumerable<IInsertDataReader> CreateReaders(IEnumerable<object> data)
    {
        if (!data.Any())
            yield break;

        List<object> itemsOfSameType = new List<object>();
        Type? previousType = null;

        foreach (var item in data)
        {
            var currentType = item?.GetType();

            if (previousType != currentType && itemsOfSameType.Count > 0)
            {
                yield return CreateReader(itemsOfSameType, previousType);
                itemsOfSameType.Clear();
            }
            itemsOfSameType.Add(item!);
            previousType = currentType;
        }

        if (itemsOfSameType.Count > 0)
        {
            yield return CreateReader(itemsOfSameType, previousType);
        }
    }

    private static IInsertDataReader CreateReader(List<object> itemsOfSameType, Type? itemType)
    {
        if (itemType == null)
            return new NullDataReader(itemsOfSameType);

        var items = Array.CreateInstance(itemType, itemsOfSameType.Count);
        Array.Copy(itemsOfSameType.ToArray(), items, items.Length);

        return InsertDataReaderFactory.CreateReader(items);
    }

    public IEnumerable<IEnumerable<XLCellValue>> GetRecords()
    {
        foreach (var reader in _readers)
        {
            foreach (var item in reader.GetRecords())
            {
                yield return item;
            }
        }
    }

    public int GetPropertiesCount()
    {
        return GetFirstNonNullReader()?.GetPropertiesCount() ?? 0;
    }

    public string? GetPropertyName(int propertyIndex)
    {
        return GetFirstNonNullReader()?.GetPropertyName(propertyIndex);
    }

    private IInsertDataReader? GetFirstNonNullReader()
    {
        return _readers.FirstOrDefault(r => r is not NullDataReader);
    }
}
