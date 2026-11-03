using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using XLibur.Extensions;

namespace XLibur.Excel.InsertData;

internal sealed class InsertDataReaderFactory
{
    private static readonly Lazy<InsertDataReaderFactory> _instance = new(() => new InsertDataReaderFactory());

    public static InsertDataReaderFactory Instance => _instance.Value;

    [RequiresUnreferencedCode("Reads the public fields and properties of the item type by reflection, which trimming can remove.")]
    [RequiresDynamicCode("Builds arrays of the items' runtime types, which native AOT cannot always precompile.")]
    public static IInsertDataReader CreateReader(IEnumerable data)
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

    public static IInsertDataReader CreateReader<T>(IEnumerable<T[]> data)
    {
        ArgumentNullException.ThrowIfNull(data);
        return new ArrayReader(data);
    }

    [RequiresUnreferencedCode("Enumerates the interfaces of a runtime type, and trimming can remove an interface implementation that is never used.")]
    public static IInsertDataReader CreateReader(IEnumerable<IEnumerable> data)
    {
        ArgumentNullException.ThrowIfNull(data);

        if (data.GetType().GetElementType() == typeof(string))
            return new SimpleTypeReader(data);

        return new ArrayReader(data);
    }

    public static IInsertDataReader CreateReader(DataTable dataTable)
    {
        ArgumentNullException.ThrowIfNull(dataTable);
        return new DataTableReader(dataTable);
    }
}
