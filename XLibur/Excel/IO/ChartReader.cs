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
            // GraphicFrame may be direct child or inside mc:AlternateContent > mc:Choice
            var graphicFrame = anchor.Elements<Xdr.GraphicFrame>().FirstOrDefault()
                ?? anchor.Descendants<Xdr.GraphicFrame>().FirstOrDefault();
            if (graphicFrame == null)
                continue;

            var graphicData = graphicFrame.Graphic?.GraphicData;
            if (graphicData == null)
                continue;

            // Try standard chart reference
            var chartRef = graphicData.Elements<C.ChartReference>().FirstOrDefault();
            if (chartRef?.Id?.Value != null)
            {
                var xlChart = LoadStandardChart(drawingsPart, chartRef.Id.Value, ws);
                if (xlChart != null)
                {
                    ReadPositions(anchor, xlChart);
                    ws.Charts.Add(xlChart);
                }
                continue;
            }

            // Try extended chart reference (cx namespace)
            // GraphicData may deserialize cx:chart as OpenXmlUnknownElement, so also check by URI + r:id
            var cxRef = graphicData.Elements<Cx.RelId>().FirstOrDefault();
            var cxRefId = cxRef?.Id?.Value;

            if (cxRefId == null && graphicData.Uri == "http://schemas.microsoft.com/office/drawing/2014/chartex")
            {
                // Fallback: find the cx:chart element as unknown element and extract r:id
                var unknownEl = graphicData.ChildElements.FirstOrDefault();
                if (unknownEl != null)
                {
                    cxRefId = unknownEl.GetAttribute("id",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;
                }
            }

            if (cxRefId != null)
            {
                var xlChart = LoadExtendedChart(drawingsPart, cxRefId, ws);
                if (xlChart != null)
                {
                    ReadPositions(anchor, xlChart);
                    ws.Charts.Add(xlChart);
                }
            }
        }
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

        // Bar/Column
        var barChart = plotArea.Elements<BarChart>().FirstOrDefault();
        if (barChart != null)
        {
            xlChart.ChartType = DetermineBarChartType(barChart);
            ReadSeriesFromElements<BarChartSeries>(barChart, xlChart.Series);
            primarySet = true;
        }

        // Pie
        var pieChart = plotArea.Elements<PieChart>().FirstOrDefault();
        if (pieChart != null && !primarySet)
        {
            xlChart.ChartType = XLChartType.Pie;
            ReadSeriesFromElements<PieChartSeries>(pieChart, xlChart.Series);
            primarySet = true;
        }

        // Line
        var lineChart = plotArea.Elements<LineChart>().FirstOrDefault();
        if (lineChart != null)
        {
            var lineType = DetermineLineChartType(lineChart);
            if (!primarySet) { xlChart.ChartType = lineType; ReadSeriesFromElements<LineChartSeries>(lineChart, xlChart.Series); primarySet = true; }
            else { xlChart.SecondaryChartType = lineType; ReadSeriesFromElements<LineChartSeries>(lineChart, xlChart.SecondarySeries); }
        }

        // Radar
        var radarChart = plotArea.Elements<RadarChart>().FirstOrDefault();
        if (radarChart != null)
        {
            var radarType = DetermineRadarChartType(radarChart);
            if (!primarySet) { xlChart.ChartType = radarType; ReadSeriesFromElements<RadarChartSeries>(radarChart, xlChart.Series); primarySet = true; }
            else { xlChart.SecondaryChartType = radarType; ReadSeriesFromElements<RadarChartSeries>(radarChart, xlChart.SecondarySeries); }
        }

        // Scatter
        var scatterChart = plotArea.Elements<ScatterChart>().FirstOrDefault();
        if (scatterChart != null && !primarySet)
        {
            xlChart.ChartType = DetermineScatterChartType(scatterChart);
            ReadScatterSeries(scatterChart, xlChart.Series);
            primarySet = true;
        }

        // Stock
        var stockChart = plotArea.Elements<StockChart>().FirstOrDefault();
        if (stockChart != null && !primarySet)
        {
            xlChart.ChartType = XLChartType.StockHighLowClose;
            ReadSeriesFromElements<LineChartSeries>(stockChart, xlChart.Series);
            primarySet = true;
        }

        // Surface
        var surfaceChart = plotArea.Elements<SurfaceChart>().FirstOrDefault();
        if (surfaceChart != null && !primarySet)
        {
            var wireframe = surfaceChart.Elements<Wireframe>().FirstOrDefault()?.Val?.Value ?? false;
            xlChart.ChartType = wireframe ? XLChartType.SurfaceWireframe : XLChartType.Surface;
            ReadSeriesFromElements<SurfaceChartSeries>(surfaceChart, xlChart.Series);
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

        var xlChart = new XLChart(ws) { IsNew = false, RelId = relId };

        // Read title from cx:chart > cx:title > cx:tx > cx:rich
        var cxChart = chartSpace.Descendants<Cx.Chart>().FirstOrDefault();
        if (cxChart != null)
        {
            var cxTitle = cxChart.Descendants<Cx.ChartTitle>().FirstOrDefault();
            if (cxTitle != null)
            {
                var titleText = string.Join("", cxTitle.Descendants<A.Text>().Select(t => t.Text));
                if (!string.IsNullOrEmpty(titleText))
                    xlChart.Title = titleText;
            }
        }

        // Read series layout to determine chart type
        var firstSeries = chartSpace.Descendants<Cx.Series>().FirstOrDefault();
        if (firstSeries != null)
        {
            var layoutId = firstSeries.GetAttribute("layoutId", string.Empty).Value ?? string.Empty;
            xlChart.ChartType = layoutId switch
            {
                "sunburst" => XLChartType.Sunburst,
                "treemap" => XLChartType.Treemap,
                "waterfall" => XLChartType.Waterfall,
                "funnel" => XLChartType.Funnel,
                "boxWhisker" => XLChartType.BoxWhisker,
                _ => XLChartType.Waterfall
            };
        }

        // Read series data from cx:chartData
        var chartData = chartSpace.Descendants<Cx.ChartData>().FirstOrDefault();
        var seriesList = chartSpace.Descendants<Cx.Series>().ToList();

        foreach (var cxSeries in seriesList)
        {
            var name = string.Empty;
            var txData = cxSeries.Descendants<Cx.TextData>().FirstOrDefault();
            if (txData != null)
            {
                var v = txData.Descendants<Cx.VXsdstring>().FirstOrDefault();
                name = v?.Text ?? string.Empty;
            }

            string? catRef = null;
            var valRef = string.Empty;

            var dataId = cxSeries.Descendants<Cx.DataId>().FirstOrDefault();
            if (dataId != null && chartData != null)
            {
                var data = chartData.Elements<Cx.Data>()
                    .FirstOrDefault(d => d.Id?.Value == dataId.Val?.Value);
                if (data != null)
                {
                    var strDim = data.Elements<Cx.StringDimension>().FirstOrDefault();
                    if (strDim != null)
                        catRef = strDim.Elements<Cx.Formula>().FirstOrDefault()?.Text;

                    var numDim = data.Elements<Cx.NumericDimension>().FirstOrDefault();
                    if (numDim != null)
                        valRef = numDim.Elements<Cx.Formula>().FirstOrDefault()?.Text ?? string.Empty;
                }
            }

            xlChart.Series.Add(name, valRef, catRef);
        }

        return xlChart;
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
        if (grouping == GroupingValues.Stacked) return XLChartType.LineStacked;
        if (grouping == GroupingValues.PercentStacked) return XLChartType.LineStacked100Percent;
        var hasMarkers = lineChart.Elements<LineChartSeries>().Any(s => s.Elements<Marker>().Any());
        return hasMarkers ? XLChartType.LineWithMarkers : XLChartType.Line;
    }

    private static XLChartType DetermineRadarChartType(RadarChart radarChart) =>
        radarChart.RadarStyle?.Val?.Value == RadarStyleValues.Filled
            ? XLChartType.RadarFilled : XLChartType.Radar;

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
