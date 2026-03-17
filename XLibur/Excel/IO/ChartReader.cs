using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Reads chart definitions from an OpenXML worksheet part and populates the worksheet's chart collection.
/// Iterates TwoCellAnchor elements looking for GraphicFrame references to ChartParts,
/// then parses chart type, title, series data, and anchor positions.
/// </summary>
internal static class ChartReader
{
    /// <summary>
    /// Loads all charts from the worksheet part's DrawingsPart into the worksheet's Charts collection.
    /// Each loaded chart has <see cref="XLChart.IsNew"/> set to <c>false</c> so it is skipped during save.
    /// </summary>
    /// <param name="worksheetPart">The OpenXML worksheet part to read from.</param>
    /// <param name="ws">The target worksheet to populate with loaded charts.</param>
    internal static void LoadCharts(WorksheetPart worksheetPart, XLWorksheet ws)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart?.WorksheetDrawing == null)
            return;

        foreach (var anchor in drawingsPart.WorksheetDrawing.Elements<Xdr.TwoCellAnchor>())
        {
            var graphicFrame = anchor.Elements<Xdr.GraphicFrame>().FirstOrDefault();
            if (graphicFrame == null)
                continue;

            var graphicData = graphicFrame.Graphic?.GraphicData;
            if (graphicData == null)
                continue;

            var chartRef = graphicData.Elements<C.ChartReference>().FirstOrDefault();
            if (chartRef?.Id?.Value == null)
                continue;

            var chartPart = (ChartPart)drawingsPart.GetPartById(chartRef.Id.Value);
            var chartSpace = chartPart.ChartSpace;
            if (chartSpace == null)
                continue;

            var chart = chartSpace.Elements<C.Chart>().FirstOrDefault();
            if (chart == null)
                continue;

            var xlChart = new XLChart((XLWorksheet)ws);
            xlChart.IsNew = false;
            xlChart.RelId = chartRef.Id.Value;

            // Read title
            var title = chart.Title;
            if (title != null)
            {
                var chartText = title.Elements<ChartText>().FirstOrDefault();
                var richText = chartText?.Elements<C.RichText>().FirstOrDefault();
                if (richText != null)
                {
                    var text = string.Join("", richText.Descendants<A.Text>()
                        .Select(t => t.Text));
                    if (!string.IsNullOrEmpty(text))
                        xlChart.Title = text;
                }
            }

            // Read chart type and series from PlotArea
            var plotArea = chart.PlotArea;
            if (plotArea != null)
            {
                ReadPlotArea(plotArea, xlChart);
            }

            // Read positions from anchor
            ReadPositions(anchor, xlChart);

            ws.Charts.Add(xlChart);
        }
    }

    /// <summary>
    /// Maps an OpenXML BarChart's direction and grouping to the corresponding <see cref="XLChartType"/>.
    /// </summary>
    private static XLChartType DetermineChartType(BarChart barChart)
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

    /// <summary>
    /// Reads all <see cref="BarChartSeries"/> elements from a BarChart and adds them to the chart's series collection.
    /// Extracts series name from StringCache, category references, and value references.
    /// </summary>
    private static void ReadBarChartSeries(BarChart barChart, XLChart xlChart)
    {
        foreach (var series in barChart.Elements<BarChartSeries>())
        {
            var name = string.Empty;
            var seriesText = series.SeriesText;
            if (seriesText != null)
            {
                var strRef = seriesText.Elements<StringReference>().FirstOrDefault();
                var strCache = strRef?.Elements<StringCache>().FirstOrDefault();
                var pt = strCache?.Elements<StringPoint>().FirstOrDefault();
                name = pt?.Elements<NumericValue>().FirstOrDefault()?.Text ?? string.Empty;
            }

            string? catRef = null;
            var catData = series.Elements<CategoryAxisData>().FirstOrDefault();
            if (catData != null)
            {
                var strRefCat = catData.Elements<StringReference>().FirstOrDefault();
                catRef = strRefCat?.Formula?.Text;

                if (catRef == null)
                {
                    var numRefCat = catData.Elements<NumberReference>().FirstOrDefault();
                    catRef = numRefCat?.Formula?.Text;
                }
            }

            var valRef = string.Empty;
            var valData = series.Elements<C.Values>().FirstOrDefault();
            if (valData != null)
            {
                var numRef = valData.Elements<NumberReference>().FirstOrDefault();
                valRef = numRef?.Formula?.Text ?? string.Empty;
            }

            xlChart.Series.Add(name, valRef, catRef);
        }
    }

    /// <summary>
    /// Reads all <see cref="PieChartSeries"/> elements from a PieChart and adds them to the chart's series collection.
    /// </summary>
    private static void ReadPieChartSeries(PieChart pieChart, XLChart xlChart)
    {
        foreach (var series in pieChart.Elements<PieChartSeries>())
        {
            var name = string.Empty;
            var seriesText = series.SeriesText;
            if (seriesText != null)
            {
                var strRef = seriesText.Elements<StringReference>().FirstOrDefault();
                var strCache = strRef?.Elements<StringCache>().FirstOrDefault();
                var pt = strCache?.Elements<StringPoint>().FirstOrDefault();
                name = pt?.Elements<NumericValue>().FirstOrDefault()?.Text ?? string.Empty;
            }

            string? catRef = null;
            var catData = series.Elements<CategoryAxisData>().FirstOrDefault();
            if (catData != null)
            {
                var strRefCat = catData.Elements<StringReference>().FirstOrDefault();
                catRef = strRefCat?.Formula?.Text;

                if (catRef == null)
                {
                    var numRefCat = catData.Elements<NumberReference>().FirstOrDefault();
                    catRef = numRefCat?.Formula?.Text;
                }
            }

            var valRef = string.Empty;
            var valData = series.Elements<C.Values>().FirstOrDefault();
            if (valData != null)
            {
                var numRef = valData.Elements<NumberReference>().FirstOrDefault();
                valRef = numRef?.Formula?.Text ?? string.Empty;
            }

            xlChart.Series.Add(name, valRef, catRef);
        }
    }

    /// <summary>
    /// Reads all chart elements from a PlotArea. The first recognized element sets the primary
    /// chart type and series; a second recognized element (combo chart) sets the secondary.
    /// </summary>
    private static void ReadPlotArea(PlotArea plotArea, XLChart xlChart)
    {
        var primarySet = false;

        var barChart = plotArea.Elements<BarChart>().FirstOrDefault();
        if (barChart != null)
        {
            xlChart.ChartType = DetermineChartType(barChart);
            ReadBarChartSeries(barChart, xlChart);
            primarySet = true;
        }

        var pieChart = plotArea.Elements<PieChart>().FirstOrDefault();
        if (pieChart != null)
        {
            if (!primarySet)
            {
                xlChart.ChartType = XLChartType.Pie;
                ReadPieChartSeries(pieChart, xlChart);
                primarySet = true;
            }
        }

        var lineChart = plotArea.Elements<LineChart>().FirstOrDefault();
        if (lineChart != null)
        {
            var lineType = DetermineLineChartType(lineChart);
            if (!primarySet)
            {
                xlChart.ChartType = lineType;
                ReadLineChartSeries(lineChart, xlChart.Series);
                primarySet = true;
            }
            else
            {
                xlChart.SecondaryChartType = lineType;
                ReadLineChartSeries(lineChart, xlChart.SecondarySeries);
            }
        }

        var radarChart = plotArea.Elements<RadarChart>().FirstOrDefault();
        if (radarChart != null)
        {
            var radarType = DetermineRadarChartType(radarChart);
            if (!primarySet)
            {
                xlChart.ChartType = radarType;
                ReadRadarChartSeries(radarChart, xlChart.Series);
            }
            else
            {
                xlChart.SecondaryChartType = radarType;
                ReadRadarChartSeries(radarChart, xlChart.SecondarySeries);
            }
        }

        // Handle combo: if bar was primary and line was secondary, check the reverse too
        if (primarySet && barChart != null && lineChart == null)
        {
            // Check if there's a second BarChart element for combo bar types
            // (not common, but handle gracefully)
        }
    }

    private static XLChartType DetermineLineChartType(LineChart lineChart)
    {
        var grouping = lineChart.Grouping?.Val?.Value;
        if (grouping == GroupingValues.Stacked) return XLChartType.LineStacked;
        if (grouping == GroupingValues.PercentStacked) return XLChartType.LineStacked100Percent;

        // Check if series have markers
        var hasMarkers = lineChart.Elements<LineChartSeries>()
            .Any(s => s.Elements<Marker>().Any());
        return hasMarkers ? XLChartType.LineWithMarkers : XLChartType.Line;
    }

    private static XLChartType DetermineRadarChartType(RadarChart radarChart)
    {
        var style = radarChart.RadarStyle?.Val?.Value;
        if (style == RadarStyleValues.Filled) return XLChartType.RadarFilled;
        return XLChartType.Radar;
    }

    /// <summary>
    /// Reads series from a LineChart element into the specified series collection.
    /// </summary>
    private static void ReadLineChartSeries(LineChart lineChart, IXLChartSeriesCollection target)
    {
        foreach (var series in lineChart.Elements<LineChartSeries>())
        {
            var (name, catRef, valRef) = ExtractSeriesData(series.SeriesText,
                series.Elements<CategoryAxisData>().FirstOrDefault(),
                series.Elements<C.Values>().FirstOrDefault());
            target.Add(name, valRef, catRef);
        }
    }

    /// <summary>
    /// Reads series from a RadarChart element into the specified series collection.
    /// </summary>
    private static void ReadRadarChartSeries(RadarChart radarChart, IXLChartSeriesCollection target)
    {
        foreach (var series in radarChart.Elements<RadarChartSeries>())
        {
            var (name, catRef, valRef) = ExtractSeriesData(series.SeriesText,
                series.Elements<CategoryAxisData>().FirstOrDefault(),
                series.Elements<C.Values>().FirstOrDefault());
            target.Add(name, valRef, catRef);
        }
    }

    /// <summary>
    /// Extracts name, category reference, and value reference from common series child elements.
    /// </summary>
    private static (string name, string? catRef, string valRef) ExtractSeriesData(
        SeriesText? seriesText, CategoryAxisData? catData, C.Values? valData)
    {
        var name = string.Empty;
        if (seriesText != null)
        {
            var strRef = seriesText.Elements<StringReference>().FirstOrDefault();
            var strCache = strRef?.Elements<StringCache>().FirstOrDefault();
            var pt = strCache?.Elements<StringPoint>().FirstOrDefault();
            name = pt?.Elements<NumericValue>().FirstOrDefault()?.Text ?? string.Empty;
        }

        string? catRef = null;
        if (catData != null)
        {
            var strRefCat = catData.Elements<StringReference>().FirstOrDefault();
            catRef = strRefCat?.Formula?.Text;
            if (catRef == null)
            {
                var numRefCat = catData.Elements<NumberReference>().FirstOrDefault();
                catRef = numRefCat?.Formula?.Text;
            }
        }

        var valRef = string.Empty;
        if (valData != null)
        {
            var numRef = valData.Elements<NumberReference>().FirstOrDefault();
            valRef = numRef?.Formula?.Text ?? string.Empty;
        }

        return (name, catRef, valRef);
    }

    /// <summary>
    /// Reads the FromMarker and ToMarker of a TwoCellAnchor into the chart's
    /// Position and <see cref="XLChart.SecondPosition"/>.
    /// Offsets are converted from EMUs (English Metric Units) to pixels (÷ 9525).
    /// </summary>
    private static void ReadPositions(Xdr.TwoCellAnchor anchor, XLChart xlChart)
    {
        var from = anchor.FromMarker;
        if (from != null)
        {
            if (int.TryParse(from.ColumnId?.Text, out var col))
                xlChart.Position.Column = col;
            if (int.TryParse(from.RowId?.Text, out var row))
                xlChart.Position.Row = row;
            if (long.TryParse(from.ColumnOffset?.Text, out var colOff))
                xlChart.Position.ColumnOffset = colOff / 9525.0;
            if (long.TryParse(from.RowOffset?.Text, out var rowOff))
                xlChart.Position.RowOffset = rowOff / 9525.0;
        }

        var to = anchor.ToMarker;
        if (to != null)
        {
            if (int.TryParse(to.ColumnId?.Text, out var col))
                xlChart.SecondPosition.Column = col;
            if (int.TryParse(to.RowId?.Text, out var row))
                xlChart.SecondPosition.Row = row;
            if (long.TryParse(to.ColumnOffset?.Text, out var colOff))
                xlChart.SecondPosition.ColumnOffset = colOff / 9525.0;
            if (long.TryParse(to.RowOffset?.Text, out var rowOff))
                xlChart.SecondPosition.RowOffset = rowOff / 9525.0;
        }
    }
}
