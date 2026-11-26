using System.Collections;
using System.Collections.Generic;
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
        unmodeledProperties.Remove(customProperty.Name);
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

    /// <summary>A modelled property that takes the name of an unmodeled one replaces it.</summary>
    internal void RemoveUnmodeled(string name)
    {
        unmodeledProperties.Remove(name);
    }

    internal IEnumerable<CustomDocumentProperty> UnmodeledPropertiesToSave => unmodeledProperties.Values;

    internal bool HasAnyToSave => customProperties.Count > 0 || unmodeledProperties.Count > 0;

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
