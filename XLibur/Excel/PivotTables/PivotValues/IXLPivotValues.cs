using System.Collections.Generic;

namespace XLibur.Excel;

public interface IXLPivotValues : IEnumerable<IXLPivotValue>
{
    /// <summary>
    /// Add a new value field to the pivot table. If addition would cause, the
    /// <see cref="XLConstants.PivotTable.ValuesSentinalLabel"/> field is added to the
    /// <see cref="IXLPivotTable.ColumnLabels"/>. The added field will use passed
    /// <paramref name="sourceName"/> as the <see cref="IXLPivotField.CustomName"/>.
    /// </summary>
    /// <param name="sourceName">The <see cref="IXLPivotField.SourceName"/> that is used as a
    ///     data. Multiple data fields can use same source (e.g. sum and count).</param>
    /// <returns>Newly added field.</returns>
    IXLPivotValue Add(string sourceName);

    /// <summary>
    /// Add a new value field to the pivot table. If addition would cause, the
    /// <see cref="XLConstants.PivotTable.ValuesSentinalLabel"/> field is added to the
    /// <see cref="IXLPivotTable.ColumnLabels"/>.
    /// </summary>
    /// <param name="sourceName">The <see cref="IXLPivotField.SourceName"/> that is used as a
    ///     data. Multiple data fields can use same source (e.g. sum and count).</param>
    /// <param name="customName">The added data field <see cref="IXLPivotField.CustomName"/>.</param>
    /// <returns>Newly added field.</returns>
    IXLPivotValue Add(string sourceName, string customName);

    void Clear();

    bool Contains(string customName);

    bool Contains(IXLPivotValue pivotValue);

    IXLPivotValue Get(string customName);

    IXLPivotValue Get(int index);

    int IndexOf(string customName);

    int IndexOf(IXLPivotValue pivotValue);

    void Remove(string customName);
}
