using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

internal static class ChartReader
{
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
                var barChart = plotArea.Elements<BarChart>().FirstOrDefault();
                if (barChart != null)
                {
                    xlChart.ChartType = DetermineChartType(barChart);
                    ReadBarChartSeries(barChart, xlChart);
                }
            }

            // Read positions from anchor
            ReadPositions(anchor, xlChart);

            ws.Charts.Add(xlChart);
        }
    }

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
