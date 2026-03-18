using System;
using System.Collections.Generic;

namespace XLibur.Excel.Caching;

/// <summary>
/// Base repository for <see cref="XLWorkbook"/> elements.
/// </summary>
internal abstract class XLWorkbookElementRepositoryBase<Tkey, Tvalue> : XLRepositoryBase<Tkey, Tvalue>
    where Tkey : struct, IEquatable<Tkey>
    where Tvalue : class
{
    public XLWorkbook Workbook { get; private set; }

    protected XLWorkbookElementRepositoryBase(XLWorkbook workbook, Func<Tkey, Tvalue> createNew)
        : this(workbook, createNew, EqualityComparer<Tkey>.Default)
    {
    }

    protected XLWorkbookElementRepositoryBase(XLWorkbook workbook, Func<Tkey, Tvalue> createNew, IEqualityComparer<Tkey> comparer)
        : base(createNew, comparer)
    {
        Workbook = workbook;
    }
}
