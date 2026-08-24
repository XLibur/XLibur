using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace XLibur.Excel;

internal sealed class XLSlicers : IXLSlicers
{
    private readonly List<XLSlicer> _slicers = [];

    public int Count => _slicers.Count;

    internal IReadOnlyList<XLSlicer> Items => _slicers;

    public IEnumerator<IXLSlicer> GetEnumerator() => _slicers.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    public IXLSlicer Slicer(string name)
    {
        if (!TryGetSlicer(name, out var slicer))
            throw new KeyNotFoundException($"The worksheet has no slicer named '{name}'.");

        return slicer;
    }

    public bool TryGetSlicer(string name, [NotNullWhen(true)] out IXLSlicer? slicer)
    {
        foreach (var candidate in _slicers)
        {
            if (XLHelper.NameComparer.Equals(candidate.Name, name))
            {
                slicer = candidate;
                return true;
            }
        }

        slicer = null;
        return false;
    }

    internal void Add(XLSlicer slicer) => _slicers.Add(slicer);
}
