using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace XLibur.Excel.InsertData;

internal sealed class InsertDataReaderFactory
{
    private static readonly Lazy<InsertDataReaderFactory> _instance = new(() => new InsertDataReaderFactory());

    public static InsertDataReaderFactory Instance => _instance.Value;

    public IInsertDataReader CreateReader(IEnumerable data)
    {
        ArgumentNullException.ThrowIfNull(data);

        var itemType = data.GetItemType();

        if (itemType == null || itemType == typeof(object))
            return new UntypedObjectReader(data);
        if (itemType.IsNullableType() && itemType.GetUnderlyingType().IsSimpleType())
            return new SimpleNullableTypeReader(data);
        if (itemType.IsSimpleType())
            return new SimpleTypeReader(data);
        if (typeof(IDataRecord).IsAssignableFrom(itemType))
            return new DataRecordReader(data.OfType<IDataRecord>());
        if (itemType.IsArray || typeof(IEnumerable).IsAssignableFrom(itemType))
            return new ArrayReader(data.Cast<IEnumerable>());
        if (itemType == typeof(DataRow))
            return new DataTableReader(data.Cast<DataRow>());

        return new ObjectReader(data);
    }

    public IInsertDataReader CreateReader<T>(IEnumerable<T[]> data)
    {
        return data == null ? throw new ArgumentNullException(nameof(data)) : new ArrayReader(data);
    }

    public IInsertDataReader CreateReader(IEnumerable<IEnumerable> data)
    {
        ArgumentNullException.ThrowIfNull(data);

        if (data.GetType().GetElementType() == typeof(string))
            return new SimpleTypeReader(data);

        return new ArrayReader(data);
    }

    public IInsertDataReader CreateReader(DataTable dataTable)
    {
        return dataTable == null ? throw new ArgumentNullException(nameof(dataTable)) : new DataTableReader(dataTable);
    }
}
