#nullable disable

using System;
using System.Collections.Generic;

namespace XLibur.Excel.Caching;

internal sealed class XLNumberFormatRepository : XLRepositoryBase<XLNumberFormatKey, XLNumberFormatValue>
{
    #region Constructors

    public XLNumberFormatRepository(Func<XLNumberFormatKey, XLNumberFormatValue> createNew)
        : base(createNew)
    {
    }

    public XLNumberFormatRepository(Func<XLNumberFormatKey, XLNumberFormatValue> createNew, IEqualityComparer<XLNumberFormatKey> comparer)
        : base(createNew, comparer)
    {
    }

    #endregion Constructors
}
