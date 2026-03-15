using System;
using System.Collections.Generic;

namespace XLibur.Excel.Caching;

internal sealed class XLFillRepository : XLRepositoryBase<XLFillKey, XLFillValue>
{
    #region Constructors

    public XLFillRepository(Func<XLFillKey, XLFillValue> createNew)
        : base(createNew)
    {
    }

    public XLFillRepository(Func<XLFillKey, XLFillValue> createNew, IEqualityComparer<XLFillKey> comparer)
        : base(createNew, comparer)
    {
    }

    #endregion Constructors
}
