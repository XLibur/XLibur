using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;

namespace XLibur.Excel;

internal sealed class XLCustomProperties : IXLCustomProperties
{
    readonly XLWorkbook workbook;
    public XLCustomProperties(XLWorkbook workbook)
    {
        this.workbook = workbook;
    }

    private readonly Dictionary<string, IXLCustomProperty> customProperties = new();

    private readonly Dictionary<string, CustomDocumentProperty> unmodeledProperties = new();

    public void Add(IXLCustomProperty customProperty)
    {
        customProperties.Add(customProperty.Name, customProperty);
    }
    public void Add<T>(string name, T value)
    {
        var cp = new XLCustomProperty(workbook) { Name = name, Value = value! };
        Add(cp);
    }

    internal void AddUnmodeled(string name, CustomDocumentProperty property)
    {
        unmodeledProperties[name] = property;
    }

    internal IEnumerable<CustomDocumentProperty> UnmodeledPropertiesToSave =>
        unmodeledProperties.Where(p => !customProperties.Values.Any(cp => cp.Name == p.Key)).Select(p => p.Value);

    internal bool HasAnyToSave => customProperties.Count > 0 || UnmodeledPropertiesToSave.Any();

    public void Delete(string name)
    {
        customProperties.Remove(name);
        unmodeledProperties.Remove(name);
    }
    public IXLCustomProperty CustomProperty(string name)
    {
        return customProperties[name];
    }

    public IEnumerator<IXLCustomProperty> GetEnumerator()
    {
        return customProperties.Values.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }


}
