#nullable disable

using System;
using System.Collections.Generic;

namespace XLibur.Excel;

public interface IXLConditionalFormats : IEnumerable<IXLConditionalFormat>
{
    void Add(IXLConditionalFormat conditionalFormat);

    void RemoveAll();

    void Remove(Predicate<IXLConditionalFormat> predicate);
}
