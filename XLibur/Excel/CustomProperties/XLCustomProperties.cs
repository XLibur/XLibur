using System.Collections.Generic;

namespace XLibur.Excel;

internal sealed class XLCustomProperties : IXLCustomProperties
{
    readonly XLWorkbook workbook;
    public XLCustomProperties(XLWorkbook workbook)
    {
        this.workbook = workbook;
    }

    private readonly Dictionary<string, IXLCustomProperty> customProperties = new();
    public void Add(IXLCustomProperty customProperty)
    {
        customProperties.Add(customProperty.Name, customProperty);
    }
    public void Add<T>(string name, T value)
    {
        var cp = new XLCustomProperty(workbook) { Name = name, Value = value! };
        Add(cp);
    }

    public void Delete(string name)
    {
        customProperties.Remove(name);
    }
    public IXLCustomProperty CustomProperty(string name)
    {
        return customProperties[name];
    }

    public IEnumerator<IXLCustomProperty> GetEnumerator()
    {
        return customProperties.Values.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }


}
