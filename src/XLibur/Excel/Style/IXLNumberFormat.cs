using System;

namespace XLibur.Excel;

public interface IXLNumberFormat : IXLNumberFormatBase, IEquatable<IXLNumberFormatBase>
{
    IXLStyle SetNumberFormatId(int value);

    IXLStyle SetFormat(string value);
}
