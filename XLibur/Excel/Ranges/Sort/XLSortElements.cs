using System.Collections.Generic;

namespace XLibur.Excel;

internal sealed class XLSortElements : IXLSortElements
{
    readonly List<IXLSortElement> elements = new List<IXLSortElement>();
    public void Add(int elementNumber)
    {
        Add(elementNumber, XLSortOrder.Ascending);
    }
    public void Add(int elementNumber, XLSortOrder sortOrder)
    {
        Add(elementNumber, sortOrder, true);
    }
    public void Add(int elementNumber, XLSortOrder sortOrder, bool ignoreBlanks)
    {
        Add(elementNumber, sortOrder, ignoreBlanks, false);
    }
    public void Add(int elementNumber, XLSortOrder sortOrder, bool ignoreBlanks, bool matchCase)
    {
        elements.Add(new XLSortElement(
            elementNumber,
            sortOrder,
            ignoreBlanks,
            matchCase));
    }

    public void Add(string elementNumber)
    {
        Add(elementNumber, XLSortOrder.Ascending);
    }
    public void Add(string elementNumber, XLSortOrder sortOrder)
    {
        Add(elementNumber, sortOrder, true);
    }
    public void Add(string elementNumber, XLSortOrder sortOrder, bool ignoreBlanks)
    {
        Add(elementNumber, sortOrder, ignoreBlanks, false);
    }
    public void Add(string elementNumber, XLSortOrder sortOrder, bool ignoreBlanks, bool matchCase)
    {
        elements.Add(new XLSortElement(
            XLHelper.GetColumnNumberFromLetter(elementNumber),
            sortOrder,
            ignoreBlanks,
            matchCase));
    }

    public IEnumerator<IXLSortElement> GetEnumerator()
    {
        return elements.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public void Clear()
    {
        elements.Clear();
    }

    public void Remove(int elementNumber)
    {
        elements.RemoveAt(elementNumber - 1);
    }

    internal void AddRange(IEnumerable<XLSortElement> sortElements) => elements.AddRange(sortElements);
}
