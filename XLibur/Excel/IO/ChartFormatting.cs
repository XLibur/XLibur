using System;
using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XLibur.Excel.IO;

/// <summary>
/// Schema child ordering for chart elements. The OpenXML SDK does not order the children of an
/// element that is built by hand, so every insertion the concept modules under
/// <c>XLibur.Excel.IO.Charts</c> make goes through here.
/// </summary>
internal static class ChartFormatting
{
    /// <summary>
    /// The schema order of the children of <c>c:ser</c> the concept modules write. Only the elements
    /// they touch need to be listed; anything else keeps its place.
    /// </summary>
    internal static readonly Type[] SeriesChildOrder =
    [
        typeof(C.Index), typeof(C.Order), typeof(C.SeriesText), typeof(C.ChartShapeProperties),
        typeof(C.Marker), typeof(C.DataLabels), typeof(C.CategoryAxisData), typeof(C.XValues),
        typeof(C.Values), typeof(C.YValues), typeof(C.Smooth), typeof(C.ExtensionList)
    ];

    /// <summary>
    /// The schema order of the children of <c>c:chart</c>.
    /// </summary>
    /// <remarks>
    /// The 3D elements between <c>c:autoTitleDeleted</c> and <c>c:plotArea</c> are listed even though
    /// nothing writes them: <see cref="InsertOrdered"/> steps over any child it cannot rank, so
    /// leaving them out put a newly inserted title after the <c>c:view3D</c> block of a 3D chart —
    /// which the schema rejects outright.
    /// </remarks>
    internal static readonly Type[] ChartChildOrder =
    [
        typeof(C.Title), typeof(C.AutoTitleDeleted), typeof(C.PivotFormats), typeof(C.View3D),
        typeof(C.Floor), typeof(C.SideWall), typeof(C.BackWall), typeof(C.PlotArea),
        typeof(C.Legend), typeof(C.PlotVisibleOnly), typeof(C.DisplayBlanksAs),
        typeof(C.ExtensionList)
    ];

    /// <summary>The schema order of the children of <c>c:title</c>.</summary>
    internal static readonly Type[] TitleChildOrder =
    [
        typeof(C.ChartText), typeof(C.Layout), typeof(C.Overlay),
        typeof(C.ChartShapeProperties), typeof(C.TextProperties), typeof(C.ExtensionList)
    ];

    /// <summary>The schema order of the children of <c>c:legend</c>.</summary>
    internal static readonly Type[] LegendChildOrder =
    [
        typeof(C.LegendPosition), typeof(C.LegendEntry), typeof(C.Layout), typeof(C.Overlay),
        typeof(C.ChartShapeProperties), typeof(C.TextProperties), typeof(C.ExtensionList)
    ];

    /// <summary>
    /// The schema order of the children of <c>c:catAx</c> and <c>c:valAx</c>. The two types agree on
    /// the elements they share; the unit elements exist on <c>c:valAx</c> only.
    /// </summary>
    internal static readonly Type[] AxisChildOrder =
    [
        typeof(C.AxisId), typeof(C.Scaling), typeof(C.Delete), typeof(C.AxisPosition),
        typeof(C.MajorGridlines), typeof(C.MinorGridlines), typeof(C.Title),
        typeof(C.NumberingFormat), typeof(C.MajorTickMark), typeof(C.MinorTickMark),
        typeof(C.TickLabelPosition), typeof(C.ChartShapeProperties), typeof(C.TextProperties),
        typeof(C.CrossingAxis), typeof(C.Crosses), typeof(C.CrossesAt), typeof(C.CrossBetween),
        typeof(C.MajorUnit), typeof(C.MinorUnit), typeof(C.DisplayUnits), typeof(C.ExtensionList)
    ];

    /// <summary>The schema order of the children of <c>c:scaling</c>.</summary>
    internal static readonly Type[] ScalingChildOrder =
    [
        typeof(C.LogBase), typeof(C.Orientation), typeof(C.MaxAxisValue), typeof(C.MinAxisValue),
        typeof(C.ExtensionList)
    ];

    /// <summary>The schema order of the children of <c>c:dLbls</c>.</summary>
    internal static readonly Type[] DataLabelsChildOrder =
    [
        typeof(C.DataLabel), typeof(C.Delete), typeof(C.NumberingFormat),
        typeof(C.ChartShapeProperties), typeof(C.TextProperties), typeof(C.DataLabelPosition),
        typeof(C.ShowLegendKey), typeof(C.ShowValue), typeof(C.ShowCategoryName),
        typeof(C.ShowSeriesName), typeof(C.ShowPercent), typeof(C.ShowBubbleSize),
        typeof(C.Separator), typeof(C.ExtensionList)
    ];

    /// <summary>
    /// Inserts <paramref name="element"/> at the position the schema order gives it: before the first
    /// child that must come after it, or at the end when there is none.
    /// </summary>
    internal static void InsertOrdered(OpenXmlCompositeElement parent, OpenXmlElement element, Type[] order)
    {
        var rank = Array.IndexOf(order, element.GetType());
        if (rank < 0)
        {
            parent.Append(element);
            return;
        }

        foreach (var child in parent.ChildElements)
        {
            var childRank = Array.IndexOf(order, child.GetType());
            if (childRank > rank)
            {
                parent.InsertBefore(element, child);
                return;
            }
        }

        parent.Append(element);
    }
}
