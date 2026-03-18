using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads chart definitions from an OpenXML worksheet part and populates the worksheet's chart collection.
/// Supports standard ChartPart charts and extended ExtendedChartPart charts (Office 2016+).
/// </summary>
internal static class ChartReader
{
    internal static void LoadCharts(WorksheetPart worksheetPart, XLWorksheet ws)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart?.WorksheetDrawing == null)
            return;

        foreach (var anchor in drawingsPart.WorksheetDrawing.Elements<Xdr.TwoCellAnchor>())
        {
            var xlChart = TryLoadChartFromAnchor(drawingsPart, anchor, ws);
            if (xlChart != null)
            {
                ReadPositions(anchor, xlChart);
                ws.Charts.Add(xlChart);
            }
        }
    }

    private static XLChart? TryLoadChartFromAnchor(
        DrawingsPart drawingsPart, Xdr.TwoCellAnchor anchor, XLWorksheet ws)
    {
        // GraphicFrame may be direct child or inside mc:AlternateContent > mc:Choice
        var graphicFrame = anchor.Elements<Xdr.GraphicFrame>().FirstOrDefault()
            ?? anchor.Descendants<Xdr.GraphicFrame>().FirstOrDefault();

        var graphicData = graphicFrame?.Graphic?.GraphicData;
        if (graphicData == null)
            return null;

        // Try standard chart reference
        var chartRef = graphicData.Elements<C.ChartReference>().FirstOrDefault();
        if (chartRef?.Id?.Value != null)
            return LoadStandardChart(drawingsPart, chartRef.Id.Value, ws);

        // Try extended chart reference (cx namespace)
        var cxRefId = ResolveExtendedChartRelId(graphicData);
        if (cxRefId != null)
            return LoadExtendedChart(drawingsPart, cxRefId, ws);

        return null;
    }

    private static string? ResolveExtendedChartRelId(A.GraphicData graphicData)
    {
        // GraphicData may deserialize cx:chart as OpenXmlUnknownElement, so also check by URI + r:id
        var cxRef = graphicData.Elements<Cx.RelId>().FirstOrDefault();
        var cxRefId = cxRef?.Id?.Value;

        if (cxRefId == null && graphicData.Uri == "http://schemas.microsoft.com/office/drawing/2014/chartex")
        {
            // Fallback: find the cx:chart element as unknown element and extract r:id
            var unknownEl = graphicData.ChildElements.Count > 0 ? graphicData.ChildElements[0] : null;
            if (unknownEl != null)
            {
                var attr = unknownEl.GetAttribute("id",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                if (!string.IsNullOrWhiteSpace(attr.Value))
                    cxRefId = attr.Value;
            }
        }

        return cxRefId;
    }

    // ── Standard chart loading ──────────────────────────────────────────

    private static XLChart? LoadStandardChart(DrawingsPart drawingsPart, string relId, XLWorksheet ws)
    {
        var chartPart = (ChartPart)drawingsPart.GetPartById(relId);
        var chartSpace = chartPart.ChartSpace;
        if (chartSpace == null) return null;

        var chart = chartSpace.Elements<C.Chart>().FirstOrDefault();
        if (chart == null) return null;

        var xlChart = new XLChart(ws) { IsNew = false, RelId = relId };
        ReadTitle(chart, xlChart);

        var plotArea = chart.PlotArea;
        if (plotArea != null)
            ReadPlotArea(plotArea, xlChart);

        return xlChart;
    }

    private static void ReadTitle(C.Chart chart, XLChart xlChart)
    {
        var title = chart.Title;
        if (title == null) return;

        var chartText = title.Elements<ChartText>().FirstOrDefault();
        var richText = chartText?.Elements<C.RichText>().FirstOrDefault();
        if (richText != null)
        {
            var text = string.Join("", richText.Descendants<A.Text>().Select(t => t.Text));
            if (!string.IsNullOrEmpty(text))
                xlChart.Title = text;
        }
    }

    private static void ReadPlotArea(PlotArea plotArea, XLChart xlChart)
    {
        var primarySet = false;

        // Primary-only chart types (cannot be secondary in a combo chart)
        primarySet |= TryReadPrimaryChart<BarChart, BarChartSeries>(plotArea, xlChart, ref primarySet, DetermineBarChartType);
        primarySet |= TryReadPrimaryChart<Bar3DChart, BarChartSeries>(plotArea, xlChart, ref primarySet, DetermineBar3DChartType);
        primarySet |= TryReadPrimaryChart<PieChart, PieChartSeries>(plotArea, xlChart, ref primarySet, _ => XLChartType.Pie);
        primarySet |= TryReadPrimaryChart<DoughnutChart, PieChartSeries>(plotArea, xlChart, ref primarySet, _ => XLChartType.Doughnut);

        // Chart types that can appear as primary or secondary (combo charts)
        TryReadComboChart<AreaChart, AreaChartSeries>(plotArea, xlChart, ref primarySet, DetermineAreaChartType);
        TryReadComboChart<LineChart, LineChartSeries>(plotArea, xlChart, ref primarySet, DetermineLineChartType);
        TryReadComboChart<RadarChart, RadarChartSeries>(plotArea, xlChart, ref primarySet, DetermineRadarChartType);

        // Primary-only chart types with custom series readers
        TryReadPrimaryChartCustom(plotArea, xlChart, ref primarySet);
    }

    private static bool TryReadPrimaryChart<TChart, TSeries>(
        PlotArea plotArea, XLChart xlChart, ref bool primarySet,
        System.Func<TChart, XLChartType> determineType)
        where TChart : OpenXmlCompositeElement
        where TSeries : OpenXmlCompositeElement
    {
        if (primarySet) return false;
        var chart = plotArea.Elements<TChart>().FirstOrDefault();
        if (chart == null) return false;

        xlChart.ChartType = determineType(chart);
        ReadSeriesFromElements<TSeries>(chart, xlChart.Series);
        return true;
    }

    private static void TryReadComboChart<TChart, TSeries>(
        PlotArea plotArea, XLChart xlChart, ref bool primarySet,
        System.Func<TChart, XLChartType> determineType)
        where TChart : OpenXmlCompositeElement
        where TSeries : OpenXmlCompositeElement
    {
        var chart = plotArea.Elements<TChart>().FirstOrDefault();
        if (chart == null) return;

        var chartType = determineType(chart);
        if (!primarySet)
        {
            xlChart.ChartType = chartType;
            ReadSeriesFromElements<TSeries>(chart, xlChart.Series);
            primarySet = true;
        }
        else
        {
            xlChart.SecondaryChartType = chartType;
            ReadSeriesFromElements<TSeries>(chart, xlChart.SecondarySeries);
        }
    }

    private static void TryReadPrimaryChartCustom(
        PlotArea plotArea, XLChart xlChart, ref bool primarySet)
    {
        // Bubble
        if (!primarySet)
        {
            var bubbleChart = plotArea.Elements<BubbleChart>().FirstOrDefault();
            if (bubbleChart != null)
            {
                xlChart.ChartType = XLChartType.Bubble;
                ReadBubbleSeries(bubbleChart, xlChart.Series);
                primarySet = true;
            }
        }

        // Scatter
        if (!primarySet)
        {
            var scatterChart = plotArea.Elements<ScatterChart>().FirstOrDefault();
            if (scatterChart != null)
            {
                xlChart.ChartType = DetermineScatterChartType(scatterChart);
                ReadScatterSeries(scatterChart, xlChart.Series);
                primarySet = true;
            }
        }

        // Stock
        if (!primarySet)
        {
            var stockChart = plotArea.Elements<StockChart>().FirstOrDefault();
            if (stockChart != null)
            {
                xlChart.ChartType = XLChartType.StockHighLowClose;
                ReadSeriesFromElements<LineChartSeries>(stockChart, xlChart.Series);
                primarySet = true;
            }
        }

        // Surface
        if (!primarySet)
        {
            var surfaceChart = plotArea.Elements<SurfaceChart>().FirstOrDefault();
            if (surfaceChart != null)
            {
                var wireframe = surfaceChart.Elements<Wireframe>().FirstOrDefault()?.Val?.Value ?? false;
                xlChart.ChartType = wireframe ? XLChartType.SurfaceWireframe : XLChartType.Surface;
                ReadSeriesFromElements<SurfaceChartSeries>(surfaceChart, xlChart.Series);
            }
        }
    }

    /// <summary>
    /// Generic reader for standard series types that use SeriesText + CategoryAxisData + Values.
    /// </summary>
    private static void ReadSeriesFromElements<TSeries>(
        OpenXmlCompositeElement parent, IXLChartSeriesCollection target)
        where TSeries : OpenXmlCompositeElement
    {
        foreach (var series in parent.Elements<TSeries>())
        {
            var (name, catRef, valRef) = ExtractSeriesData(
                series.Elements<SeriesText>().FirstOrDefault(),
                series.Elements<CategoryAxisData>().FirstOrDefault(),
                series.Elements<C.Values>().FirstOrDefault());
            target.Add(name, valRef, catRef);
        }
    }

    private static void ReadScatterSeries(ScatterChart scatterChart, IXLChartSeriesCollection target)
    {
        foreach (var series in scatterChart.Elements<ScatterChartSeries>())
        {
            var name = ExtractSeriesName(series.Elements<SeriesText>().FirstOrDefault());

            string? xRef = null;
            var xValues = series.Elements<XValues>().FirstOrDefault();
            if (xValues != null)
            {
                var numRef = xValues.Elements<NumberReference>().FirstOrDefault();
                xRef = numRef?.Formula?.Text;
                if (xRef == null)
                {
                    var strRef = xValues.Elements<StringReference>().FirstOrDefault();
                    xRef = strRef?.Formula?.Text;
                }
            }

            var yRef = string.Empty;
            var yValues = series.Elements<YValues>().FirstOrDefault();
            if (yValues != null)
            {
                var numRef = yValues.Elements<NumberReference>().FirstOrDefault();
                yRef = numRef?.Formula?.Text ?? string.Empty;
            }

            target.Add(name, yRef, xRef);
        }
    }

    // ── Extended chart loading ──────────────────────────────────────────

    private static XLChart? LoadExtendedChart(DrawingsPart drawingsPart, string relId, XLWorksheet ws)
    {
        var extPart = (ExtendedChartPart)drawingsPart.GetPartById(relId);
        var chartSpace = extPart.ChartSpace;
        if (chartSpace == null) return null;

        var chartType = ReadExtendedChartType(chartSpace);
        if (chartType == null) return null;

        var xlChart = new XLChart(ws) { IsNew = false, RelId = relId, ChartType = chartType.Value };

        ReadExtendedTitle(chartSpace, xlChart);
        ReadExtendedSeries(chartSpace, xlChart);

        return xlChart;
    }

    private static void ReadExtendedTitle(Cx.ChartSpace chartSpace, XLChart xlChart)
    {
        var cxTitle = chartSpace.Descendants<Cx.ChartTitle>().FirstOrDefault();
        if (cxTitle == null) return;

        var titleText = string.Join("", cxTitle.Descendants<A.Text>().Select(t => t.Text));
        if (!string.IsNullOrEmpty(titleText))
            xlChart.Title = titleText;
    }

    private static XLChartType? ReadExtendedChartType(Cx.ChartSpace chartSpace)
    {
        var firstSeries = chartSpace.Descendants<Cx.Series>().FirstOrDefault();
        if (firstSeries == null) return null;

        var layoutId = firstSeries.GetAttribute("layoutId", string.Empty).Value ?? string.Empty;
        return layoutId switch
        {
            "sunburst" => XLChartType.Sunburst,
            "treemap" => XLChartType.Treemap,
            "waterfall" => XLChartType.Waterfall,
            "funnel" => XLChartType.Funnel,
            "boxWhisker" => XLChartType.BoxWhisker,
            _ => null
        };
    }

    private static void ReadExtendedSeries(Cx.ChartSpace chartSpace, XLChart xlChart)
    {
        var chartData = chartSpace.Descendants<Cx.ChartData>().FirstOrDefault();

        foreach (var cxSeries in chartSpace.Descendants<Cx.Series>())
        {
            var name = ReadExtendedSeriesName(cxSeries);
            var (catRef, valRef) = ReadExtendedSeriesRefs(cxSeries, chartData);
            xlChart.Series.Add(name, valRef, catRef);
        }
    }

    private static string ReadExtendedSeriesName(Cx.Series cxSeries)
    {
        var txData = cxSeries.Descendants<Cx.TextData>().FirstOrDefault();
        if (txData == null) return string.Empty;

        return txData.Descendants<Cx.VXsdstring>().FirstOrDefault()?.Text ?? string.Empty;
    }

    private static (string? catRef, string valRef) ReadExtendedSeriesRefs(
        Cx.Series cxSeries, Cx.ChartData? chartData)
    {
        var dataId = cxSeries.Descendants<Cx.DataId>().FirstOrDefault();
        if (dataId == null || chartData == null)
            return (null, string.Empty);

        var data = chartData.Elements<Cx.Data>()
            .FirstOrDefault(d => d.Id?.Value == dataId.Val?.Value);
        if (data == null)
            return (null, string.Empty);

        var catRef = data.Elements<Cx.StringDimension>().FirstOrDefault()
            ?.Elements<Cx.Formula>().FirstOrDefault()?.Text;

        var valRef = data.Elements<Cx.NumericDimension>().FirstOrDefault()
            ?.Elements<Cx.Formula>().FirstOrDefault()?.Text ?? string.Empty;

        return (catRef, valRef);
    }

    // ── Type determination helpers ──────────────────────────────────────

    private static XLChartType DetermineBarChartType(BarChart barChart)
    {
        var direction = barChart.BarDirection?.Val?.Value ?? BarDirectionValues.Column;
        var grouping = barChart.BarGrouping?.Val?.Value ?? BarGroupingValues.Clustered;

        if (direction == BarDirectionValues.Bar)
        {
            if (grouping == BarGroupingValues.Stacked) return XLChartType.BarStacked;
            if (grouping == BarGroupingValues.PercentStacked) return XLChartType.BarStacked100Percent;
            return XLChartType.BarClustered;
        }

        if (grouping == BarGroupingValues.Stacked) return XLChartType.ColumnStacked;
        if (grouping == BarGroupingValues.PercentStacked) return XLChartType.ColumnStacked100Percent;
        return XLChartType.ColumnClustered;
    }

    private static XLChartType DetermineLineChartType(LineChart lineChart)
    {
        var grouping = lineChart.Grouping?.Val?.Value;
        var hasMarkers = lineChart.Elements<LineChartSeries>().Any(s => s.Elements<Marker>().Any());

        if (grouping == GroupingValues.Stacked)
            return hasMarkers ? XLChartType.LineWithMarkersStacked : XLChartType.LineStacked;
        if (grouping == GroupingValues.PercentStacked)
            return hasMarkers ? XLChartType.LineWithMarkersStacked100Percent : XLChartType.LineStacked100Percent;

        return hasMarkers ? XLChartType.LineWithMarkers : XLChartType.Line;
    }

    private static XLChartType DetermineRadarChartType(RadarChart radarChart) =>
        radarChart.RadarStyle?.Val?.Value == RadarStyleValues.Filled
            ? XLChartType.RadarFilled : XLChartType.Radar;

    private static XLChartType DetermineBar3DChartType(Bar3DChart bar3DChart)
    {
        var direction = bar3DChart.BarDirection?.Val?.Value ?? BarDirectionValues.Column;
        var grouping = bar3DChart.BarGrouping?.Val?.Value ?? BarGroupingValues.Clustered;
        var shape = bar3DChart.Elements<Shape>().FirstOrDefault()?.Val?.Value;
        var isHorizontal = direction == BarDirectionValues.Bar;

        if (shape == ShapeValues.Cone || shape == ShapeValues.ConeToMax)
            return ResolveBar3DGrouping(isHorizontal, grouping,
                horizontal: (XLChartType.ConeHorizontalClustered, XLChartType.ConeHorizontalStacked, XLChartType.ConeHorizontalStacked100Percent),
                vertical: (XLChartType.ConeClustered, XLChartType.ConeStacked, XLChartType.ConeStacked100Percent),
                verticalStandard: XLChartType.Cone);

        if (shape == ShapeValues.Cylinder)
            return ResolveBar3DGrouping(isHorizontal, grouping,
                horizontal: (XLChartType.CylinderHorizontalClustered, XLChartType.CylinderHorizontalStacked, XLChartType.CylinderHorizontalStacked100Percent),
                vertical: (XLChartType.CylinderClustered, XLChartType.CylinderStacked, XLChartType.CylinderStacked100Percent),
                verticalStandard: XLChartType.Cylinder);

        if (shape == ShapeValues.Pyramid || shape == ShapeValues.PyramidToMaximum)
            return ResolveBar3DGrouping(isHorizontal, grouping,
                horizontal: (XLChartType.PyramidHorizontalClustered, XLChartType.PyramidHorizontalStacked, XLChartType.PyramidHorizontalStacked100Percent),
                vertical: (XLChartType.PyramidClustered, XLChartType.PyramidStacked, XLChartType.PyramidStacked100Percent),
                verticalStandard: XLChartType.Pyramid);

        // Default: Box shape = standard 3D bar/column
        return ResolveBar3DBoxGrouping(isHorizontal, grouping);
    }

    private static XLChartType ResolveBar3DGrouping(
        bool isHorizontal, BarGroupingValues grouping,
        (XLChartType Clustered, XLChartType Stacked, XLChartType Stacked100) horizontal,
        (XLChartType Clustered, XLChartType Stacked, XLChartType Stacked100) vertical,
        XLChartType verticalStandard)
    {
        if (isHorizontal)
        {
            if (grouping == BarGroupingValues.Stacked) return horizontal.Stacked;
            if (grouping == BarGroupingValues.PercentStacked) return horizontal.Stacked100;
            return horizontal.Clustered;
        }

        if (grouping == BarGroupingValues.Stacked) return vertical.Stacked;
        if (grouping == BarGroupingValues.PercentStacked) return vertical.Stacked100;
        if (grouping == BarGroupingValues.Standard) return verticalStandard;
        return vertical.Clustered;
    }

    private static XLChartType ResolveBar3DBoxGrouping(bool isHorizontal, BarGroupingValues grouping)
    {
        if (isHorizontal)
        {
            if (grouping == BarGroupingValues.Stacked) return XLChartType.BarStacked3D;
            if (grouping == BarGroupingValues.PercentStacked) return XLChartType.BarStacked100Percent3D;
            return XLChartType.BarClustered3D;
        }

        if (grouping == BarGroupingValues.Stacked) return XLChartType.ColumnStacked3D;
        if (grouping == BarGroupingValues.PercentStacked) return XLChartType.ColumnStacked100Percent3D;
        if (grouping == BarGroupingValues.Standard) return XLChartType.Column3D;
        return XLChartType.ColumnClustered3D;
    }

    private static XLChartType DetermineAreaChartType(AreaChart areaChart)
    {
        var grouping = areaChart.Grouping?.Val?.Value;
        if (grouping == GroupingValues.Stacked) return XLChartType.AreaStacked;
        if (grouping == GroupingValues.PercentStacked) return XLChartType.AreaStacked100Percent;
        return XLChartType.Area;
    }

    private static void ReadBubbleSeries(BubbleChart bubbleChart, IXLChartSeriesCollection target)
    {
        foreach (var series in bubbleChart.Elements<BubbleChartSeries>())
        {
            var name = ExtractSeriesName(series.Elements<SeriesText>().FirstOrDefault());

            string? xRef = null;
            var xValues = series.Elements<XValues>().FirstOrDefault();
            if (xValues != null)
            {
                xRef = xValues.Elements<NumberReference>().FirstOrDefault()?.Formula?.Text;
                xRef ??= xValues.Elements<StringReference>().FirstOrDefault()?.Formula?.Text;
            }

            var yRef = string.Empty;
            var yValues = series.Elements<YValues>().FirstOrDefault();
            if (yValues != null)
                yRef = yValues.Elements<NumberReference>().FirstOrDefault()?.Formula?.Text ?? string.Empty;

            target.Add(name, yRef, xRef);
        }
    }

    private static XLChartType DetermineScatterChartType(ScatterChart scatterChart)
    {
        var style = scatterChart.ScatterStyle?.Val?.Value;
        if (style == ScatterStyleValues.SmoothMarker)
            return XLChartType.XYScatterSmoothLinesWithMarkers;
        return XLChartType.XYScatterMarkers;
    }

    // ── Shared extraction helpers ───────────────────────────────────────

    private static (string name, string? catRef, string valRef) ExtractSeriesData(
        SeriesText? seriesText, CategoryAxisData? catData, C.Values? valData)
    {
        var name = ExtractSeriesName(seriesText);

        string? catRef = null;
        if (catData != null)
        {
            catRef = catData.Elements<StringReference>().FirstOrDefault()?.Formula?.Text;
            catRef ??= catData.Elements<NumberReference>().FirstOrDefault()?.Formula?.Text;
        }

        var valRef = string.Empty;
        if (valData != null)
        {
            valRef = valData.Elements<NumberReference>().FirstOrDefault()?.Formula?.Text ?? string.Empty;
        }

        return (name, catRef, valRef);
    }

    private static string ExtractSeriesName(SeriesText? seriesText)
    {
        if (seriesText == null) return string.Empty;
        var strRef = seriesText.Elements<StringReference>().FirstOrDefault();
        var strCache = strRef?.Elements<StringCache>().FirstOrDefault();
        var pt = strCache?.Elements<StringPoint>().FirstOrDefault();
        return pt?.Elements<NumericValue>().FirstOrDefault()?.Text ?? string.Empty;
    }

    // ── Position reading ────────────────────────────────────────────────

    private static void ReadPositions(Xdr.TwoCellAnchor anchor, XLChart xlChart)
    {
        var from = anchor.FromMarker;
        if (from != null)
        {
            if (int.TryParse(from.ColumnId?.Text, out var col)) xlChart.Position.Column = col;
            if (int.TryParse(from.RowId?.Text, out var row)) xlChart.Position.Row = row;
            if (long.TryParse(from.ColumnOffset?.Text, out var colOff)) xlChart.Position.ColumnOffset = colOff / 9525.0;
            if (long.TryParse(from.RowOffset?.Text, out var rowOff)) xlChart.Position.RowOffset = rowOff / 9525.0;
        }

        var to = anchor.ToMarker;
        if (to != null)
        {
            if (int.TryParse(to.ColumnId?.Text, out var col)) xlChart.SecondPosition.Column = col;
            if (int.TryParse(to.RowId?.Text, out var row)) xlChart.SecondPosition.Row = row;
            if (long.TryParse(to.ColumnOffset?.Text, out var colOff)) xlChart.SecondPosition.ColumnOffset = colOff / 9525.0;
            if (long.TryParse(to.RowOffset?.Text, out var rowOff)) xlChart.SecondPosition.RowOffset = rowOff / 9525.0;
        }
    }
}
