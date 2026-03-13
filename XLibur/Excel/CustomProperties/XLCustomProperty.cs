using System;
using System.Linq;

namespace XLibur.Excel;

internal sealed class XLCustomProperty : IXLCustomProperty
{
    private readonly XLWorkbook _workbook;

    private string name = string.Empty;

    public XLCustomProperty(XLWorkbook workbook)
    {
        _workbook = workbook;
    }

    #region IXLCustomProperty Members

    public string Name
    {
        get { return name; }
        set
        {
            if (name == value) return;

            if (_workbook.CustomProperties.Any(t => t.Name == value))
                throw new ArgumentException(
                    $"This workbook already contains a custom property named '{value}'");

            name = value;
        }
    }

    public XLCustomPropertyType Type
    {
        get
        {
            if (Value is DateTime)
                return XLCustomPropertyType.Date;

            if (Value is bool)
                return XLCustomPropertyType.Boolean;

            if (Value is double or int or long or float or decimal or short or byte or sbyte or ushort or uint or ulong)
                return XLCustomPropertyType.Number;

            return XLCustomPropertyType.Text;
        }
    }

    public object Value { get; set; } = null!;

    public T GetValue<T>()
    {
        return (T)Convert.ChangeType(Value, typeof(T));
    }

    #endregion
}
