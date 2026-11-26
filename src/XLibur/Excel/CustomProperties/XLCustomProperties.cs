using System;
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

    /// <summary>
    /// The modelled properties in insertion order. A property's name lives only on the property, so renaming one
    /// leaves this list untouched.
    /// </summary>
    private readonly List<IXLCustomProperty> customProperties = new();

    private readonly Dictionary<string, CustomDocumentProperty> unmodeledProperties = new();

    public void Add(IXLCustomProperty customProperty)
    {
        if (Find(customProperty.Name) is not null)
            throw new ArgumentException(
                $"This workbook already contains a custom property named '{customProperty.Name}'");

        customProperties.Add(customProperty);
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
        customProperties.RemoveAll(p => p.Name == name);
        unmodeledProperties.Remove(name);
    }
    public IXLCustomProperty CustomProperty(string name)
    {
        return Find(name) ?? throw new KeyNotFoundException($"The workbook has no custom property named '{name}'.");
    }

    private IXLCustomProperty? Find(string name) => customProperties.Find(p => p.Name == name);

    /// <summary>
    /// Enumerates a snapshot, so a caller can rename, delete or add properties inside a <c>foreach</c>.
    /// </summary>
    public IEnumerator<IXLCustomProperty> GetEnumerator()
    {
        return ((IEnumerable<IXLCustomProperty>)customProperties.ToArray()).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }


}
