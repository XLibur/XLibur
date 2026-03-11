using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

namespace XLibur.Excel;

using System.Collections;

internal sealed class XLTables : IXLTables, IEnumerable<XLTable>
{
    private readonly Dictionary<string, XLTable> _tables;

    public XLTables()
    {
        _tables = new Dictionary<string, XLTable>(StringComparer.OrdinalIgnoreCase);
        Deleted = new HashSet<string>();
    }

    internal ICollection<string> Deleted { get; }

    #region IXLTables Members

    bool IXLTables.TryGetTable(string tableName, out IXLTable? table)
    {
        if (TryGetTable(tableName, out var foundTable))
        {
            table = foundTable;
            return true;
        }

        table = default;
        return false;
    }

    public void Add(IXLTable table)
    {
        var xlTable = (XLTable)table;
        _tables.Add(table.Name, xlTable);
        xlTable.OnAddedToTables();
    }

    public IXLTables Clear(XLClearOptions clearOptions = XLClearOptions.All)
    {
        _tables.Values.ForEach(t => t.Clear(clearOptions));
        return this;
    }

    public bool Contains(string name)
    {
        return _tables.ContainsKey(name);
    }

    public Dictionary<string, XLTable>.ValueCollection.Enumerator GetEnumerator()
    {
        return _tables.Values.GetEnumerator();
    }

    IEnumerator<XLTable> IEnumerable<XLTable>.GetEnumerator() => GetEnumerator();

    IEnumerator<IXLTable> IEnumerable<IXLTable>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public void Remove(int index)
    {
        Remove(_tables.ElementAt(index).Key);
    }

    public void Remove(string name)
    {
        if (!_tables.Remove(name, out var table))
            throw new ArgumentOutOfRangeException(nameof(name), $"Unable to delete table because the table name {name} could not be found.");

        var relId = table.RelId;

        if (relId is not null)
            Deleted.Add(relId);
    }

    public IXLTable Table(int index)
    {
        return _tables.ElementAt(index).Value;
    }

    public IXLTable Table(string name)
    {
        if (TryGetTable(name, out var table))
            return table;

        throw new ArgumentOutOfRangeException(nameof(name), $"Table {name} was not found.");
    }

    internal bool TryGetTable(string tableName, [MaybeNullWhen(false)] out XLTable table)
    {
        return _tables.TryGetValue(tableName, out table);
    }

    #endregion IXLTables Members
}
