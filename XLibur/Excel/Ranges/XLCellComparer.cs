using System;

namespace XLibur.Excel.Ranges;

internal static class XLCellComparer
{
    internal static int CompareCells(XLCell thisCell, XLCell otherCell, IXLSortElement e)
    {
        var thisCellIsBlank = thisCell.IsEmpty();
        var otherCellIsBlank = otherCell.IsEmpty();

        if (e.IgnoreBlanks && (thisCellIsBlank || otherCellIsBlank))
            return CompareBlanks(thisCellIsBlank, otherCellIsBlank, e.SortOrder);

        return CompareSameOrMixedTypes(thisCell, otherCell, e.MatchCase);
    }

    private static int CompareBlanks(bool thisCellIsBlank, bool otherCellIsBlank, XLSortOrder sortOrder)
    {
        return thisCellIsBlank switch
        {
            true when otherCellIsBlank => 0,
            true => sortOrder == XLSortOrder.Ascending ? 1 : -1,
            _ => sortOrder == XLSortOrder.Ascending ? -1 : 1
        };
    }

    private static int CompareSameOrMixedTypes(XLCell thisCell, XLCell otherCell, bool matchCase)
    {
        return thisCell.DataType != otherCell.DataType
            ? CompareMixedTypes(thisCell, otherCell, matchCase)
            : CompareSameType(thisCell, otherCell, matchCase);
    }

    private static int CompareMixedTypes(XLCell thisCell, XLCell otherCell, bool matchCase)
    {
        if (thisCell.Value.IsUnifiedNumber && otherCell.Value.IsUnifiedNumber)
            return thisCell.Value.GetUnifiedNumber().CompareTo(otherCell.Value.GetUnifiedNumber());

        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        return string.Compare(thisCell.GetString(), otherCell.GetString(), comparison);
    }

    private static int CompareSameType(XLCell thisCell, XLCell otherCell, bool matchCase)
    {
        switch (thisCell.DataType)
        {
            case XLDataType.Blank:
            case XLDataType.Error:
                return 0;
            case XLDataType.Boolean:
                return thisCell.GetBoolean().CompareTo(otherCell.GetBoolean());
            case XLDataType.Text:
                var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                return string.Compare(thisCell.GetText(), otherCell.GetText(), comparison);
            case XLDataType.TimeSpan:
                return thisCell.GetTimeSpan().CompareTo(otherCell.GetTimeSpan());
            case XLDataType.DateTime:
                return thisCell.GetDateTime().CompareTo(otherCell.GetDateTime());
            case XLDataType.Number:
                return thisCell.GetDouble().CompareTo(otherCell.GetDouble());
            default:
                throw new NotImplementedException(
                    $"Cell comparison not implemented for data type {thisCell.DataType}.");
        }
    }
}
