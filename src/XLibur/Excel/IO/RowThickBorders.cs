using System.Collections.Generic;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.IO;

/// <summary>
/// The <c>thickTop</c> and <c>thickBot</c> flags the borders of a sheet's cells give its rows (#1031).
/// Excel draws a row with either flag a little taller, and does not work the flags out on load, so the
/// save writes them from the borders as they are now.
/// </summary>
/// <remarks>
/// The rule is the one Excel follows in the files it writes (every row of the test fixtures that carry
/// the flags, 1013 rows, bar one <c>thickTop</c>). A medium line grows down only and a thick or double
/// line grows both ways:
/// <list type="bullet">
/// <item><c>thickBot</c>: a cell of the row has a medium, thick or double bottom border, or a cell of
/// the row below has such a top border. Medium includes the dashed medium styles and slant dash dot.</item>
/// <item><c>thickTop</c>: a cell of the row has a thick or double top border, or a cell of the row
/// above has such a bottom border.</item>
/// </list>
/// The cells read are those with a style record and, for a formatted row, the row's own style. A cell
/// without a record that takes a column's border is not read.
/// </remarks>
internal sealed class RowThickBorders
{
    private enum Weight : byte
    {
        Thin,
        Medium,
        Thick,
    }

    private readonly Dictionary<int, (Weight Top, Weight Bottom)> _edges = [];

    private RowThickBorders()
    {
    }

    internal static RowThickBorders Of(XLWorksheet worksheet)
    {
        var borders = new RowThickBorders();
        using (var styles = new Slice<XLStyleValue?>.Enumerator(worksheet.Internals.CellsCollection.StyleSlice, Area.Full))
        {
            while (styles.MoveNext())
            {
                if (styles.Current is { } style)
                    borders.Add(styles.Point.Row, style);
            }
        }

        foreach (var (rowNumber, row) in worksheet.Internals.RowsCollection)
        {
            if (row.IsFormatted)
                borders.Add(rowNumber, row.StyleValue);
        }

        return borders;
    }

    /// <summary>The flags of a row: <c>thickTop</c>, then <c>thickBot</c>.</summary>
    internal (bool Top, bool Bottom) FlagsOf(int rowNumber)
    {
        var (top, bottom) = EdgesOf(rowNumber);
        return (
            top == Weight.Thick || EdgesOf(rowNumber - 1).Bottom == Weight.Thick,
            bottom >= Weight.Medium || EdgesOf(rowNumber + 1).Top >= Weight.Medium);
    }

    /// <summary>
    /// Every row that has a flag: a row with a heavy border, and its neighbour that a heavy border
    /// reaches into. Excel writes the neighbour even when it has nothing else.
    /// </summary>
    internal IEnumerable<int> RowsWithFlags()
    {
        foreach (var rowNumber in _edges.Keys)
        {
            for (var candidate = rowNumber - 1; candidate <= rowNumber + 1; candidate++)
            {
                if (candidate is >= 1 and <= XLHelper.MaxRowNumber && FlagsOf(candidate) is (true, _) or (_, true))
                    yield return candidate;
            }
        }
    }

    private (Weight Top, Weight Bottom) EdgesOf(int rowNumber) =>
        _edges.TryGetValue(rowNumber, out var edges) ? edges : default;

    private void Add(int rowNumber, XLStyleValue style)
    {
        var top = WeightOf(style.Border.TopBorder);
        var bottom = WeightOf(style.Border.BottomBorder);
        if (top == Weight.Thin && bottom == Weight.Thin)
            return;

        var (oldTop, oldBottom) = EdgesOf(rowNumber);
        _edges[rowNumber] = (Max(oldTop, top), Max(oldBottom, bottom));
    }

    private static Weight Max(Weight a, Weight b) => a > b ? a : b;

    private static Weight WeightOf(XLBorderStyleValues style) => style switch
    {
        XLBorderStyleValues.Thick or XLBorderStyleValues.Double => Weight.Thick,
        XLBorderStyleValues.Medium or XLBorderStyleValues.MediumDashed or XLBorderStyleValues.MediumDashDot
            or XLBorderStyleValues.MediumDashDotDot or XLBorderStyleValues.SlantDashDot => Weight.Medium,
        _ => Weight.Thin,
    };
}
