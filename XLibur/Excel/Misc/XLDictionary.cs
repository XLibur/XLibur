using System.Collections.Generic;
using XLibur.Extensions;

namespace XLibur.Excel;

public class XLDictionary<T> : Dictionary<int, T>
    where T : notnull
{
    public XLDictionary()
    {

    }
    public XLDictionary(XLDictionary<T> other)
    {
        other.Values.ForEach(Add);
    }

    public void Initialize(T value)
    {
        if (Count > 0)
            Clear();

        Add(value);
    }

    public void Add(T value)
    {
        Add(Count + 1, value);
    }
}
