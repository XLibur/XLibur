using System;
using System.Linq;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace XLibur.Excel.IO.Charts;

/// <summary>
/// The <c>c:legend</c> element: how it is read into <see cref="XLChartLegend"/>, and how an assigned
/// model is written back.
/// </summary>
/// <remarks>
/// <see cref="Apply"/> covers both a chart being created and a chart loaded from a file. The two used
/// to be separate functions — <c>BuildLegend</c> and <c>PatchLegend</c> — that had to agree by hand;
/// creating the element is simply the branch <see cref="Apply"/> takes when there is none.
/// </remarks>
internal static class ChartLegendXml
{
    /// <summary>
    /// Seeds the model from the chart's <c>c:legend</c>. A chart with no legend element seeds
    /// <see cref="IXLChartLegend.Visible"/> as <c>false</c>.
    /// </summary>
    internal static void Read(C.Chart chart, XLChartLegend legend)
    {
        var element = chart.Elements<C.Legend>().FirstOrDefault();
        if (element == null)
        {
            legend.SeedLoaded(visible: false, XLLegendPosition.Right, overlay: false);
            return;
        }

        legend.SeedLoaded(
            visible: true,
            position: ReadPosition(element),
            overlay: element.Elements<C.Overlay>().FirstOrDefault()?.Val?.Value ?? false);
    }

    /// <summary>
    /// Writes the assigned legend properties into <paramref name="chart"/>, creating, editing or
    /// removing the <c>c:legend</c> child as the model requires. A chart with no assigned legend
    /// properties is not modified at all, which is what lets an untouched legend round-trip byte for
    /// byte. The legend's own text and shape properties are left alone.
    /// </summary>
    internal static void Apply(C.Chart chart, XLChartLegend legend)
    {
        var assigned = legend.AssignedFormat;
        if (assigned == XLChartLegendFormat.None)
            return;

        var element = chart.Elements<C.Legend>().FirstOrDefault();

        if ((assigned & XLChartLegendFormat.Visible) != 0 && !legend.Visible)
        {
            element?.Remove();
            return;
        }

        if (element == null)
        {
            // Position and Overlay are ignored while the legend is hidden, so assigning one of them
            // on a chart that has no legend must not conjure one.
            if (!legend.Visible)
                return;

            element = new C.Legend();
            element.Append(new C.LegendPosition { Val = MapPosition(legend.Position) });
            element.Append(new C.Overlay { Val = legend.Overlay });
            ChartFormatting.InsertOrdered(chart, element, ChartFormatting.ChartChildOrder);
            return;
        }

        if ((assigned & XLChartLegendFormat.Position) != 0)
        {
            foreach (var existing in element.Elements<C.LegendPosition>().ToList())
                existing.Remove();
            ChartFormatting.InsertOrdered(element,
                new C.LegendPosition { Val = MapPosition(legend.Position) },
                ChartFormatting.LegendChildOrder);
        }

        if ((assigned & XLChartLegendFormat.Overlay) != 0)
        {
            foreach (var existing in element.Elements<C.Overlay>().ToList())
                existing.Remove();
            ChartFormatting.InsertOrdered(element, new C.Overlay { Val = legend.Overlay },
                ChartFormatting.LegendChildOrder);
        }
    }

    private static XLLegendPosition ReadPosition(C.Legend element)
    {
        var position = element.Elements<C.LegendPosition>().FirstOrDefault()?.Val;
        if (position == null)
            return XLLegendPosition.Right;

        var value = position.Value;
        if (value == C.LegendPositionValues.Bottom) return XLLegendPosition.Bottom;
        if (value == C.LegendPositionValues.Left) return XLLegendPosition.Left;
        if (value == C.LegendPositionValues.Top) return XLLegendPosition.Top;
        if (value == C.LegendPositionValues.TopRight) return XLLegendPosition.TopRight;
        return XLLegendPosition.Right;
    }

    private static C.LegendPositionValues MapPosition(XLLegendPosition position) => position switch
    {
        XLLegendPosition.Right => C.LegendPositionValues.Right,
        XLLegendPosition.Bottom => C.LegendPositionValues.Bottom,
        XLLegendPosition.Left => C.LegendPositionValues.Left,
        XLLegendPosition.Top => C.LegendPositionValues.Top,
        XLLegendPosition.TopRight => C.LegendPositionValues.TopRight,
        _ => throw new ArgumentOutOfRangeException(nameof(position), position,
            "Unknown legend position.")
    };
}
