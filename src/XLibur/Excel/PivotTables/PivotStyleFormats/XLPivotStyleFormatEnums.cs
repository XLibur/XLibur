using System;

namespace XLibur.Excel;

[Flags]
public enum XLPivotStyleFormatElement
{
    None = 0,
    Label = 1 << 1,
    Data = 1 << 2,

    All = Label | Data
}
