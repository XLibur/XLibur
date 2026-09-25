using System;
using System.Diagnostics;

namespace XLibur.Excel;

/// <summary>
/// A list of pivot cache values that takes one value at a time: the shared items of a field
/// (<see cref="XLPivotCacheSharedItems"/>) or its records (<see cref="XLPivotCacheValues"/>).
/// </summary>
/// <remarks>
/// Both lists are filled from the same sources — cell values on a cache refresh, cache items on
/// load — so the dispatch from a source to an <c>Add*</c> method is written once, against this
/// interface, instead of once per list.
/// </remarks>
internal interface IXLPivotCacheValueSink
{
    void AddMissing();

    void AddNumber(double number);

    void AddBoolean(bool boolean);

    void AddError(XLError error);

    void AddString(string text);

    void AddDateTime(DateTime dateTime);
}

internal static class XLPivotCacheValueSinkExtensions
{
    /// <summary>
    /// Add a cell value to the sink, using the cache representation of its type.
    /// </summary>
    /// <remarks>
    /// A <see cref="TimeSpan"/> has no cache type of its own, so it is stored as a date, see
    /// <see cref="XLPivotCacheValue.ToCacheDateTime"/>.
    /// </remarks>
    internal static void AddCellValue(this IXLPivotCacheValueSink sink, XLCellValue value)
    {
        switch (value.Type)
        {
            case XLDataType.Blank:
                sink.AddMissing();
                break;
            case XLDataType.Boolean:
                sink.AddBoolean(value.GetBoolean());
                break;
            case XLDataType.Number:
                sink.AddNumber(value.GetNumber());
                break;
            case XLDataType.Text:
                sink.AddString(value.GetText());
                break;
            case XLDataType.Error:
                sink.AddError(value.GetError());
                break;
            case XLDataType.DateTime:
                sink.AddDateTime(value.GetDateTime());
                break;
            case XLDataType.TimeSpan:
                sink.AddDateTime(XLPivotCacheValue.ToCacheDateTime(value.GetTimeSpan()));
                break;
            default:
                throw new UnreachableException();
        }
    }
}
