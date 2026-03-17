using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using XLibur.Excel.ContentManagers;
using static XLibur.Excel.XLWorkbook;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes newly created charts to OpenXML. For each chart with <see cref="XLChart.IsNew"/> == <c>true</c>,
/// creates a ChartPart containing the ChartSpace DOM and a TwoCellAnchor with a GraphicFrame
/// referencing the chart. Ensures the worksheet's <c>&lt;drawing&gt;</c> element exists.
/// Charts loaded from an existing file (<see cref="XLChart.IsNew"/> == <c>false</c>) are skipped
/// and preserved untouched by OpenXML.
/// </summary>
internal static class ChartWriter
{
    /// <summary>
    /// Writes all new charts from the worksheet to the OpenXML worksheet part.
    /// Skips charts that were loaded from an existing file.
    /// </summary>
    internal static void WriteCharts(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        XLWorksheet xlWorksheet,
        WorksheetPart worksheetPart,
        SaveContext context)
    {
        foreach (var chart in xlWorksheet.Charts)
        {
            var xlChart = (XLChart)chart;
            if (!xlChart.IsNew)
                continue;

            WriteChart(worksheet, cm, worksheetPart, xlChart, context);
        }
    }

    /// <summary>
    /// Writes a single chart: creates the ChartPart, builds the ChartSpace DOM,
    /// appends a TwoCellAnchor/GraphicFrame to the DrawingsPart, and ensures
    /// the worksheet XML contains a <c>&lt;drawing&gt;</c> reference.
    /// </summary>
    private static void WriteChart(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        WorksheetPart worksheetPart,
        XLChart xlChart,
        SaveContext context)
    {
        var drawingsPart = worksheetPart.DrawingsPart ??
                           worksheetPart.AddNewPart<DrawingsPart>(context.RelIdGenerator.GetNext(RelType.Workbook));

        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

        var worksheetDrawing = drawingsPart.WorksheetDrawing;

        EnsureNamespaces(worksheetDrawing);

        // Create chart part
        var chartRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
        var chartPart = drawingsPart.AddNewPart<ChartPart>(chartRelId);

        // Build chart DOM
        chartPart.ChartSpace = BuildChartSpace(xlChart);

        // Create TwoCellAnchor with GraphicFrame
        var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
        var nvpId = nvps.Any()
            ? (UInt32Value)nvps.Max(p => p.Id!.Value) + 1
            : 1U;

        var fromPos = xlChart.Position;
        var toPos = xlChart.SecondPosition;

        var anchor = new Xdr.TwoCellAnchor(
            new Xdr.FromMarker
            {
                ColumnId = new Xdr.ColumnId(fromPos.Column.ToString()),
                RowId = new Xdr.RowId(fromPos.Row.ToString()),
                ColumnOffset = new Xdr.ColumnOffset(((long)(fromPos.ColumnOffset * 9525)).ToString()),
                RowOffset = new Xdr.RowOffset(((long)(fromPos.RowOffset * 9525)).ToString())
            },
            new Xdr.ToMarker
            {
                ColumnId = new Xdr.ColumnId(toPos.Column.ToString()),
                RowId = new Xdr.RowId(toPos.Row.ToString()),
                ColumnOffset = new Xdr.ColumnOffset(((long)(toPos.ColumnOffset * 9525)).ToString()),
                RowOffset = new Xdr.RowOffset(((long)(toPos.RowOffset * 9525)).ToString())
            },
            new Xdr.GraphicFrame(
                new Xdr.NonVisualGraphicFrameProperties(
                    new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = xlChart.Name },
                    new Xdr.NonVisualGraphicFrameDrawingProperties()
                ),
                new Xdr.Transform(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = 0, Cy = 0 }
                ),
                new A.Graphic(
                    new A.GraphicData(
                        new C.ChartReference { Id = chartRelId }
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
                )
            ),
            new Xdr.ClientData()
        );

        worksheetDrawing.Append(anchor);

        // Ensure <drawing> element in worksheet XML
        if (!worksheet.OfType<Drawing>().Any())
        {
            var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
            var worksheetDrawingRef = new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) };
            worksheetDrawingRef.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            if (tableParts != null)
                worksheet.InsertBefore(worksheetDrawingRef, tableParts);
            else
                worksheet.AppendChild(worksheetDrawingRef);

            cm.SetElement(XLWorksheetContents.Drawing, worksheet.Elements<Drawing>().First());
        }
    }

    /// <summary>
    /// Builds the complete OpenXML ChartSpace DOM for a chart.
    /// Dispatches to pie or bar/column chart builders based on the chart type.
    /// </summary>
    private static ChartSpace BuildChartSpace(XLChart xlChart)
    {
        var chart = new C.Chart();

        if (xlChart.Title != null)
        {
            chart.Title = new Title(
                new ChartText(
                    new C.RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(
                                new A.RunProperties { Language = "en-US" },
                                new A.Text(xlChart.Title)
                            )
                        )
                    )
                ),
                new Overlay { Val = false }
            );
        }

        chart.Append(BuildPlotArea(xlChart));
        chart.Append(new PlotVisibleOnly { Val = true });

        return new ChartSpace(chart);
    }

    private static bool IsPieType(XLChartType chartType) =>
        chartType is XLChartType.Pie or XLChartType.PieExploded
            or XLChartType.Pie3D or XLChartType.PieExploded3D
            or XLChartType.PieToPie or XLChartType.PieToBar;

    private static bool IsLineType(XLChartType chartType) =>
        chartType is XLChartType.Line or XLChartType.Line3D
            or XLChartType.LineStacked or XLChartType.LineStacked100Percent
            or XLChartType.LineWithMarkers or XLChartType.LineWithMarkersStacked
            or XLChartType.LineWithMarkersStacked100Percent;

    private static bool IsRadarType(XLChartType chartType) =>
        chartType is XLChartType.Radar or XLChartType.RadarFilled
            or XLChartType.RadarWithMarkers;

    /// <summary>
    /// Builds the PlotArea, dispatching to the appropriate chart element builder.
    /// For combo charts, emits both a primary and secondary chart element sharing axes.
    /// </summary>
    private static PlotArea BuildPlotArea(XLChart xlChart)
    {
        if (IsPieType(xlChart.ChartType))
            return BuildPiePlotArea(xlChart);

        const uint catAxisId = 1u;
        const uint valAxisId = 2u;

        var plotArea = new PlotArea();
        plotArea.Append(new Layout());

        // Primary chart element
        AppendChartElement(plotArea, xlChart.ChartType, xlChart.Series, catAxisId, valAxisId);

        // Secondary chart element (combo charts)
        if (xlChart.SecondaryChartType.HasValue && xlChart.SecondarySeries.Count > 0)
        {
            AppendChartElement(plotArea, xlChart.SecondaryChartType.Value, xlChart.SecondarySeries,
                catAxisId, valAxisId);
        }

        // Shared axes
        plotArea.Append(new CategoryAxis(
            new AxisId { Val = catAxisId },
            new Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new Delete { Val = false },
            new AxisPosition { Val = AxisPositionValues.Bottom },
            new CrossingAxis { Val = valAxisId }
        ));
        plotArea.Append(new ValueAxis(
            new AxisId { Val = valAxisId },
            new Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new Delete { Val = false },
            new AxisPosition { Val = AxisPositionValues.Left },
            new CrossingAxis { Val = catAxisId }
        ));

        return plotArea;
    }

    /// <summary>
    /// Appends a typed chart element (BarChart, LineChart, or RadarChart) with its series to the PlotArea.
    /// </summary>
    private static void AppendChartElement(
        PlotArea plotArea,
        XLChartType chartType,
        IXLChartSeriesCollection seriesCollection,
        uint catAxisId,
        uint valAxisId)
    {
        if (IsLineType(chartType))
        {
            var lineChart = new LineChart
            {
                Grouping = new Grouping { Val = GetLineGrouping(chartType) }
            };

            foreach (var s in seriesCollection)
            {
                var series = new LineChartSeries
                {
                    Index = new Index { Val = s.Index },
                    Order = new Order { Val = s.Order },
                    SeriesText = BuildSeriesText(s)
                };

                if (s.CategoryReferences != null)
                {
                    series.Append(new CategoryAxisData(
                        new StringReference { Formula = new C.Formula(s.CategoryReferences) }
                    ));
                }

                series.Append(new C.Values(
                    new NumberReference { Formula = new C.Formula(s.ValueReferences) }
                ));

                if (chartType is XLChartType.LineWithMarkers
                    or XLChartType.LineWithMarkersStacked
                    or XLChartType.LineWithMarkersStacked100Percent)
                {
                    series.Append(new Marker { Symbol = new Symbol { Val = MarkerStyleValues.Auto } });
                }

                lineChart.Append(series);
            }

            lineChart.Append(new AxisId { Val = catAxisId });
            lineChart.Append(new AxisId { Val = valAxisId });
            plotArea.Append(lineChart);
        }
        else if (IsRadarType(chartType))
        {
            var radarStyle = chartType == XLChartType.RadarFilled
                ? RadarStyleValues.Filled
                : RadarStyleValues.Marker;

            var radarChart = new RadarChart
            {
                RadarStyle = new RadarStyle { Val = radarStyle }
            };

            foreach (var s in seriesCollection)
            {
                var series = new RadarChartSeries
                {
                    Index = new Index { Val = s.Index },
                    Order = new Order { Val = s.Order },
                    SeriesText = BuildSeriesText(s)
                };

                if (s.CategoryReferences != null)
                {
                    series.Append(new CategoryAxisData(
                        new StringReference { Formula = new C.Formula(s.CategoryReferences) }
                    ));
                }

                series.Append(new C.Values(
                    new NumberReference { Formula = new C.Formula(s.ValueReferences) }
                ));

                radarChart.Append(series);
            }

            radarChart.Append(new AxisId { Val = catAxisId });
            radarChart.Append(new AxisId { Val = valAxisId });
            plotArea.Append(radarChart);
        }
        else
        {
            // Bar/Column chart
            var xlChart = new XLChart_BarParams(chartType);
            var barChart = new BarChart
            {
                BarDirection = new BarDirection { Val = xlChart.Direction },
                BarGrouping = new BarGrouping { Val = xlChart.Grouping }
            };

            foreach (var s in seriesCollection)
            {
                var series = new BarChartSeries
                {
                    Index = new Index { Val = s.Index },
                    Order = new Order { Val = s.Order },
                    SeriesText = BuildSeriesText(s)
                };

                if (s.CategoryReferences != null)
                {
                    series.Append(new CategoryAxisData(
                        new StringReference { Formula = new C.Formula(s.CategoryReferences) }
                    ));
                }

                series.Append(new C.Values(
                    new NumberReference { Formula = new C.Formula(s.ValueReferences) }
                ));

                barChart.Append(series);
            }

            barChart.Append(new AxisId { Val = catAxisId });
            barChart.Append(new AxisId { Val = valAxisId });
            plotArea.Append(barChart);
        }
    }

    /// <summary>
    /// Builds a PlotArea containing a PieChart. Pie charts have no axes.
    /// </summary>
    private static PlotArea BuildPiePlotArea(XLChart xlChart)
    {
        var pieChart = new PieChart();

        foreach (var s in xlChart.Series)
        {
            var series = new PieChartSeries
            {
                Index = new Index { Val = s.Index },
                Order = new Order { Val = s.Order },
                SeriesText = BuildSeriesText(s)
            };

            if (s.CategoryReferences != null)
            {
                series.Append(new CategoryAxisData(
                    new StringReference { Formula = new C.Formula(s.CategoryReferences) }
                ));
            }

            series.Append(new C.Values(
                new NumberReference { Formula = new C.Formula(s.ValueReferences) }
            ));

            pieChart.Append(series);
        }

        return new PlotArea(new Layout(), pieChart);
    }

    private static GroupingValues GetLineGrouping(XLChartType chartType) =>
        chartType is XLChartType.LineStacked or XLChartType.LineWithMarkersStacked
            ? GroupingValues.Stacked
            : chartType is XLChartType.LineStacked100Percent or XLChartType.LineWithMarkersStacked100Percent
                ? GroupingValues.PercentStacked
                : GroupingValues.Standard;

    /// <summary>
    /// Lightweight helper that resolves bar direction/grouping from a chart type
    /// without needing a full XLChart instance.
    /// </summary>
    private readonly struct XLChart_BarParams
    {
        public BarDirectionValues Direction { get; }
        public BarGroupingValues Grouping { get; }

        public XLChart_BarParams(XLChartType chartType)
        {
            Direction = IsHorizontalBarType(chartType) ? BarDirectionValues.Bar : BarDirectionValues.Column;
            Grouping = GetBarGroupingForType(chartType);
        }

        private static bool IsHorizontalBarType(XLChartType ct) =>
            ct is XLChartType.BarClustered or XLChartType.BarClustered3D
                or XLChartType.BarStacked or XLChartType.BarStacked100Percent
                or XLChartType.BarStacked100Percent3D or XLChartType.BarStacked3D;

        private static BarGroupingValues GetBarGroupingForType(XLChartType ct) =>
            ct is XLChartType.BarClustered or XLChartType.BarClustered3D
                or XLChartType.ColumnClustered or XLChartType.ColumnClustered3D
                ? BarGroupingValues.Clustered
                : ct is XLChartType.BarStacked or XLChartType.BarStacked3D
                    or XLChartType.ColumnStacked or XLChartType.ColumnStacked3D
                    ? BarGroupingValues.Stacked
                    : ct is XLChartType.BarStacked100Percent or XLChartType.BarStacked100Percent3D
                        or XLChartType.ColumnStacked100Percent or XLChartType.ColumnStacked100Percent3D
                        ? BarGroupingValues.PercentStacked
                        : BarGroupingValues.Clustered;
    }

    /// <summary>
    /// Builds a SeriesText element containing the series name in a StringCache.
    /// </summary>
    private static SeriesText BuildSeriesText(IXLChartSeries s)
    {
        return new SeriesText(
            new StringReference(
                new StringCache(
                    new PointCount { Val = 1 },
                    new StringPoint(new NumericValue(s.Name)) { Index = 0 }
                )
            )
        );
    }

    /// <summary>
    /// Ensures the WorksheetDrawing element has the required DrawingML and relationships namespace declarations.
    /// </summary>
    private static void EnsureNamespaces(Xdr.WorksheetDrawing worksheetDrawing)
    {
        if (!worksheetDrawing.NamespaceDeclarations.Any(nd =>
                nd.Value.Equals("http://schemas.openxmlformats.org/drawingml/2006/main")))
            worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

        if (!worksheetDrawing.NamespaceDeclarations.Any(nd =>
                nd.Value.Equals("http://schemas.openxmlformats.org/officeDocument/2006/relationships")))
            worksheetDrawing.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    }
}
