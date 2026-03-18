using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using XLibur.Excel.ContentManagers;
using static XLibur.Excel.XLWorkbook;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Drawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.IO;

/// <summary>
/// Writes newly created charts to OpenXML. Supports standard chart types (bar, line, pie,
/// scatter, stock, surface, radar) via ChartPart, and extended chart types (sunburst, treemap,
/// waterfall, funnel, box &amp; whisker) via ExtendedChartPart.
/// </summary>
internal static class ChartWriter
{
    /// <summary>
    /// Writes all new charts from the worksheet to the OpenXML worksheet part.
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

            if (IsExtendedType(xlChart.ChartType))
                WriteExtendedChart(worksheet, cm, worksheetPart, xlChart, context);
            else
                WriteStandardChart(worksheet, cm, worksheetPart, xlChart, context);

            xlChart.IsNew = false;
        }
    }

    // ── Standard chart writing ──────────────────────────────────────────

    private static void WriteStandardChart(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        WorksheetPart worksheetPart,
        XLChart xlChart,
        SaveContext context)
    {
        var drawingsPart = EnsureDrawingsPart(worksheetPart, context);
        var worksheetDrawing = drawingsPart.WorksheetDrawing!;
        EnsureNamespaces(worksheetDrawing);

        var chartRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
        var chartPart = drawingsPart.AddNewPart<ChartPart>(chartRelId);
        chartPart.ChartSpace = BuildChartSpace(xlChart);

        AppendAnchor(worksheetDrawing, xlChart,
            new A.GraphicData(new C.ChartReference { Id = chartRelId })
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" });

        EnsureDrawingElement(worksheet, cm, worksheetPart, drawingsPart);
    }

    // ── Extended chart writing (Sunburst, Treemap, Waterfall, Funnel, BoxWhisker) ──

    /// <summary>
    /// Counter for generating unique extended chart part URIs.
    /// Reset per save operation via the SaveContext lifecycle.
    /// </summary>
    [ThreadStatic]
    private static int _extChartCounter;

    internal static void ResetExtendedChartCounter() => _extChartCounter = 0;

    private static void WriteExtendedChart(
        Worksheet worksheet,
        XLWorksheetContentManager cm,
        WorksheetPart worksheetPart,
        XLChart xlChart,
        SaveContext context)
    {
        var drawingsPart = EnsureDrawingsPart(worksheetPart, context);
        var worksheetDrawing = drawingsPart.WorksheetDrawing!;
        EnsureNamespaces(worksheetDrawing);

        var chartRelId = context.RelIdGenerator.GetNext(RelType.Workbook);

        // The OpenXML SDK's AddNewPart<ExtendedChartPart> places the part under
        // xl/drawings/extendedCharts/ which Excel rejects. Excel expects extended
        // charts at xl/charts/chartExN.xml. Use the IPackageFeature to access the
        // underlying System.IO.Packaging.Package and create the part at the correct URI.
        _extChartCounter++;
        var partUri = new Uri($"/xl/charts/chartEx{_extChartCounter}.xml", UriKind.Relative);

#pragma warning disable OOXML0001 // Experimental API needed to place ExtendedChartPart at xl/charts/
        var package = DocumentFormat.OpenXml.Experimental.PackageExtensions.GetPackage(worksheetPart.OpenXmlPackage);
#pragma warning restore OOXML0001

        var packagePart = package.CreatePart(
            partUri,
            "application/vnd.ms-office.chartex+xml",
            System.IO.Packaging.CompressionOption.Normal);

        // Write chart XML
        var chartSpace = BuildExtendedChartSpace(xlChart);
        using (var stream = packagePart.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write))
        {
            chartSpace.Save(stream);
        }

        // Create relationship from DrawingsPart to the chart part using relative path
        // Excel requires relative target URIs for extended chart relationships
        var relativeTarget = new Uri("../charts/chartEx" + _extChartCounter + ".xml", UriKind.Relative);
        var drawingsPackagePart = package.GetPart(drawingsPart.Uri);
        drawingsPackagePart.Relationships.Create(
            relativeTarget,
            System.IO.Packaging.TargetMode.Internal,
            "http://schemas.microsoft.com/office/2014/relationships/chartEx",
            chartRelId);

        // Excel requires chart style and color files for extended charts
        WriteExtendedChartStyleAndColor(package, packagePart, _extChartCounter);

        AppendExtendedAnchor(worksheetDrawing, xlChart, chartRelId);

        EnsureDrawingElement(worksheet, cm, worksheetPart, drawingsPart);

        // The SDK hoists mc/cx namespaces to the wsDr root element, which can confuse Excel.
        // Write the drawing XML manually to control namespace placement.
        SaveDrawingWithLocalNamespaces(drawingsPart);
    }

    /// <summary>
    /// Creates the chart style and color files required by Excel for extended charts.
    /// These are siblings of the chart part at xl/charts/ with their own content types and relationships.
    /// </summary>
#pragma warning disable OOXML0001
    private static void WriteExtendedChartStyleAndColor(
        DocumentFormat.OpenXml.Packaging.IPackage package,
        DocumentFormat.OpenXml.Packaging.IPackagePart chartPart,
        int chartIndex)
#pragma warning restore OOXML0001
    {
        var colorsUri = new Uri($"/xl/charts/colors{chartIndex}.xml", UriKind.Relative);
        var styleUri = new Uri($"/xl/charts/style{chartIndex}.xml", UriKind.Relative);

        // Create color style part
        var colorsPart = package.CreatePart(colorsUri,
            "application/vnd.ms-office.chartcolorstyle+xml",
            System.IO.Packaging.CompressionOption.Normal);
        using (var stream = colorsPart.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write))
        {
            var asm = typeof(ChartWriter).Assembly;
            using var resStream = asm.GetManifestResourceStream("XLibur.Excel.IO.ChartExDefaultColors.xml")!;
            resStream.CopyTo(stream);
        }

        // Create chart style part
        var stylePart = package.CreatePart(styleUri,
            "application/vnd.ms-office.chartstyle+xml",
            System.IO.Packaging.CompressionOption.Normal);
        using (var stream = stylePart.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write))
        {
            var asm = typeof(ChartWriter).Assembly;
            using var resStream = asm.GetManifestResourceStream("XLibur.Excel.IO.ChartExDefaultStyle.xml")!;
            resStream.CopyTo(stream);
        }

        // Create relationships from chart part to style and color parts
        var colorsRelTarget = new Uri($"colors{chartIndex}.xml", UriKind.Relative);
        var styleRelTarget = new Uri($"style{chartIndex}.xml", UriKind.Relative);

        chartPart.Relationships.Create(
            styleRelTarget,
            System.IO.Packaging.TargetMode.Internal,
            "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
            "rId1");
        chartPart.Relationships.Create(
            colorsRelTarget,
            System.IO.Packaging.TargetMode.Internal,
            "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle",
            "rId2");
    }

    /// <summary>
    /// Re-serializes the WorksheetDrawing to move mc/cx namespace declarations from the root
    /// element to local elements where they are used. Excel is strict about namespace placement
    /// on the wsDr root element for extended chart drawings.
    /// </summary>
    private static void SaveDrawingWithLocalNamespaces(DrawingsPart drawingsPart)
    {
        var xml = drawingsPart.WorksheetDrawing!.OuterXml;

        // Remove mc, cx1, cx, a16 namespace declarations from the root wsDr element
        // These will remain on the child elements where the SDK placed them originally
        var prefixesToRemove = new[] { "mc", "cx1", "cx", "a16" };
        foreach (var prefix in prefixesToRemove)
        {
            xml = System.Text.RegularExpressions.Regex.Replace(
                xml,
                $@"\s*xmlns:{prefix}=""[^""]*""",
                "",
                System.Text.RegularExpressions.RegexOptions.None,
                System.TimeSpan.FromSeconds(1));
        }

        // Re-add the namespace declarations on the elements that use them
        // mc: on mc:AlternateContent
        xml = xml.Replace(
            "<mc:AlternateContent>",
            @"<mc:AlternateContent xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"">");

        // cx1: on mc:Choice
        xml = xml.Replace(
            "<mc:Choice ",
            @"<mc:Choice xmlns:cx1=""http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"" ");

        // cx: on cx:chart
        xml = xml.Replace(
            "<cx:chart ",
            @"<cx:chart xmlns:cx=""http://schemas.microsoft.com/office/drawing/2014/chartex"" ");

        // a16: on a16:creationId
        xml = xml.Replace(
            "<a16:creationId ",
            @"<a16:creationId xmlns:a16=""http://schemas.microsoft.com/office/drawing/2014/main"" ");

        // Re-parse the fixed XML back into the SDK DOM
        drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing(xml);
    }

    private static Cx.ChartSpace BuildExtendedChartSpace(XLChart xlChart)
    {
        var layoutId = xlChart.ChartType switch
        {
            XLChartType.Sunburst => Cx.SeriesLayout.Sunburst,
            XLChartType.Treemap => Cx.SeriesLayout.Treemap,
            XLChartType.Waterfall => Cx.SeriesLayout.Waterfall,
            XLChartType.Funnel => Cx.SeriesLayout.Funnel,
            XLChartType.BoxWhisker => Cx.SeriesLayout.BoxWhisker,
            _ => throw new NotSupportedException($"Extended chart type {xlChart.ChartType} is not supported.")
        };

        var isSunburstOrTreemap = xlChart.ChartType is XLChartType.Sunburst or XLChartType.Treemap;
        var isWaterfall = xlChart.ChartType == XLChartType.Waterfall;

        var plotAreaRegion = new Cx.PlotAreaRegion();
        var chartData = new Cx.ChartData();
        uint dataIdx = 0;

        foreach (var s in xlChart.Series)
        {
            plotAreaRegion.AppendChild(BuildExtendedSeries(s, layoutId, dataIdx, isWaterfall));
            chartData.AppendChild(BuildExtendedData(s, dataIdx, isSunburstOrTreemap));
            dataIdx++;
        }

        var plotArea = BuildExtendedPlotArea(plotAreaRegion, isSunburstOrTreemap);

        var cxChart = new Cx.Chart();
        if (xlChart.Title != null)
            cxChart.AppendChild(BuildExtendedChartTitle(xlChart.Title));
        cxChart.AppendChild(plotArea);

        var chartSpace = new Cx.ChartSpace();
        chartSpace.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        chartSpace.AppendChild(chartData);
        chartSpace.AppendChild(cxChart);
        return chartSpace;
    }

    private static Cx.Series BuildExtendedSeries(
        IXLChartSeries s, Cx.SeriesLayout layoutId, uint dataIdx, bool isWaterfall)
    {
        var cxSeries = new Cx.Series
        {
            LayoutId = layoutId,
            FormatIdx = dataIdx,
            UniqueId = "{" + System.Guid.NewGuid().ToString() + "}"
        };

        if (!string.IsNullOrEmpty(s.Name))
        {
            var txData = new Cx.TextData();
            txData.AppendChild(new Cx.VXsdstring(s.Name));
            var tx = new Cx.Text();
            tx.AppendChild(txData);
            cxSeries.AppendChild(tx);
        }

        cxSeries.AppendChild(new Cx.DataId { Val = dataIdx });

        if (isWaterfall)
        {
            var layoutPr = new Cx.SeriesLayoutProperties();
            layoutPr.AppendChild(new Cx.Subtotals());
            cxSeries.AppendChild(layoutPr);
        }

        return cxSeries;
    }

    private static Cx.Data BuildExtendedData(
        IXLChartSeries s, uint dataIdx, bool isSunburstOrTreemap)
    {
        var data = new Cx.Data { Id = dataIdx };

        if (s.CategoryReferences != null)
        {
            var strDim = new Cx.StringDimension { Type = Cx.StringDimensionType.Cat };
            var catFormula = new Cx.Formula(s.CategoryReferences);
            // Sunburst/Treemap with multi-column category ranges need dir="col"
            // to indicate each column is a hierarchy level
            if (isSunburstOrTreemap && s.CategoryReferences.Contains(':'))
                catFormula.SetAttribute(new OpenXmlAttribute("dir", string.Empty, "col"));
            strDim.AppendChild(catFormula);
            data.AppendChild(strDim);
        }

        var numDimType = isSunburstOrTreemap
            ? Cx.NumericDimensionType.Size
            : Cx.NumericDimensionType.Val;
        var numDim = new Cx.NumericDimension { Type = numDimType };
        numDim.AppendChild(new Cx.Formula(s.ValueReferences));
        data.AppendChild(numDim);

        return data;
    }

    private static Cx.PlotArea BuildExtendedPlotArea(
        Cx.PlotAreaRegion plotAreaRegion, bool isSunburstOrTreemap)
    {
        var plotArea = new Cx.PlotArea();
        plotArea.AppendChild(plotAreaRegion);

        if (!isSunburstOrTreemap)
        {
            var catAxis = new Cx.Axis { Id = 0u };
            catAxis.AppendChild(new Cx.CategoryAxisScaling());
            catAxis.AppendChild(new Cx.TickLabels());
            plotArea.AppendChild(catAxis);

            var valAxis = new Cx.Axis { Id = 1u };
            valAxis.AppendChild(new Cx.ValueAxisScaling());
            valAxis.AppendChild(new Cx.MajorGridlinesGridlines());
            valAxis.AppendChild(new Cx.TickLabels());
            plotArea.AppendChild(valAxis);
        }

        return plotArea;
    }

    private static Cx.ChartTitle BuildExtendedChartTitle(string titleText)
    {
        var title = new Cx.ChartTitle
        {
            Pos = Cx.SidePos.T,
            Align = Cx.PosAlign.Ctr,
            Overlay = false
        };
        var rich = new Cx.RichTextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(
                new A.Run(
                    new A.RunProperties { Language = "en-US" },
                    new A.Text(titleText)
                )
            )
        );
        var txTitle = new Cx.Text();
        txTitle.AppendChild(rich);
        title.AppendChild(txTitle);
        return title;
    }

    // ── Standard ChartSpace building ────────────────────────────────────

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

    private static PlotArea BuildPlotArea(XLChart xlChart)
    {
        if (IsPieType(xlChart.ChartType) || IsDoughnutType(xlChart.ChartType))
            return BuildNoAxesPlotArea(xlChart);

        if (IsBubbleType(xlChart.ChartType))
            return BuildBubblePlotArea(xlChart);

        const uint catAxisId = 1u;
        const uint valAxisId = 2u;
        const uint serAxisId = 3u;

        var plotArea = new PlotArea();
        plotArea.Append(new Layout());

        AppendChartElement(plotArea, xlChart.ChartType, xlChart.Series, catAxisId, valAxisId, 0);

        if (xlChart.SecondaryChartType.HasValue && xlChart.SecondarySeries.Count > 0)
        {
            // Secondary series indices must continue from primary to avoid conflicts
            AppendChartElement(plotArea, xlChart.SecondaryChartType.Value, xlChart.SecondarySeries,
                catAxisId, valAxisId, (uint)xlChart.Series.Count);
        }

        // Axes depend on primary chart type
        if (IsScatterType(xlChart.ChartType))
        {
            // Scatter uses two ValueAxis (X and Y)
            plotArea.Append(new ValueAxis(
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
        }
        else if (IsSurfaceType(xlChart.ChartType))
        {
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
            plotArea.Append(new SeriesAxis(
                new AxisId { Val = serAxisId },
                new Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new Delete { Val = false },
                new AxisPosition { Val = AxisPositionValues.Bottom },
                new CrossingAxis { Val = valAxisId }
            ));
        }
        else
        {
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
        }

        return plotArea;
    }

    private static void AppendChartElement(
        PlotArea plotArea, XLChartType chartType,
        IXLChartSeriesCollection seriesCollection, uint catAxisId, uint valAxisId, uint indexOffset)
    {
        if (IsAreaType(chartType))
            AppendAreaChart(plotArea, chartType, seriesCollection, catAxisId, valAxisId, indexOffset);
        else if (IsLineType(chartType))
            AppendLineChart(plotArea, chartType, seriesCollection, catAxisId, valAxisId, indexOffset);
        else if (IsRadarType(chartType))
            AppendRadarChart(plotArea, chartType, seriesCollection, catAxisId, valAxisId, indexOffset);
        else if (IsScatterType(chartType))
            AppendScatterChart(plotArea, chartType, seriesCollection, catAxisId, valAxisId, indexOffset);
        else if (IsStockType(chartType))
            AppendStockChart(plotArea, seriesCollection, catAxisId, valAxisId, indexOffset);
        else if (IsSurfaceType(chartType))
            AppendSurfaceChart(plotArea, chartType, seriesCollection, catAxisId, valAxisId, indexOffset);
        else if (IsBar3DType(chartType))
            AppendBar3DChart(plotArea, chartType, seriesCollection, catAxisId, valAxisId, indexOffset);
        else
            AppendBarChart(plotArea, chartType, seriesCollection, catAxisId, valAxisId, indexOffset);
    }

    // ── Pie / Doughnut (no axes) ──

    private static PlotArea BuildNoAxesPlotArea(XLChart xlChart)
    {
        OpenXmlCompositeElement chartElement;

        if (IsDoughnutType(xlChart.ChartType))
        {
            var doughnut = new DoughnutChart();
            foreach (var s in xlChart.Series)
            {
                var series = new PieChartSeries
                {
                    Index = new C.Index { Val = s.Index },
                    Order = new Order { Val = s.Order },
                    SeriesText = BuildSeriesText(s)
                };
                AppendCatAndVal(series, s);
                doughnut.Append(series);
            }
            chartElement = doughnut;
        }
        else
        {
            var pie = new PieChart();
            foreach (var s in xlChart.Series)
            {
                var series = new PieChartSeries
                {
                    Index = new C.Index { Val = s.Index },
                    Order = new Order { Val = s.Order },
                    SeriesText = BuildSeriesText(s)
                };
                AppendCatAndVal(series, s);
                pie.Append(series);
            }
            chartElement = pie;
        }

        return new PlotArea(new Layout(), chartElement);
    }

    // ── Area ──

    private static void AppendAreaChart(
        PlotArea plotArea, XLChartType chartType,
        IXLChartSeriesCollection seriesCollection, uint catAxisId, uint valAxisId, uint indexOffset)
    {
        var areaChart = new AreaChart
        {
            Grouping = new Grouping { Val = GetAreaGrouping(chartType) }
        };
        foreach (var s in seriesCollection)
        {
            var series = new AreaChartSeries
            {
                Index = new C.Index { Val = s.Index + indexOffset },
                Order = new Order { Val = s.Order + indexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendCatAndVal(series, s);
            areaChart.Append(series);
        }
        areaChart.Append(new AxisId { Val = catAxisId });
        areaChart.Append(new AxisId { Val = valAxisId });
        plotArea.Append(areaChart);
    }

    // ── Bubble ──

    private static PlotArea BuildBubblePlotArea(XLChart xlChart)
    {
        // Bubble charts use XValues + YValues + BubbleSize, and two ValueAxis (like scatter).
        // CategoryReferences = X values, ValueReferences = Y values.
        // For simplicity, bubble size defaults to the Y values if no separate size data.
        const uint xAxisId = 1u;
        const uint yAxisId = 2u;

        var bubbleChart = new BubbleChart();
        foreach (var s in xlChart.Series)
        {
            var series = new BubbleChartSeries
            {
                Index = new C.Index { Val = s.Index },
                Order = new Order { Val = s.Order },
                SeriesText = BuildSeriesText(s)
            };
            if (s.CategoryReferences != null)
            {
                series.Append(new XValues(
                    new NumberReference { Formula = new C.Formula(s.CategoryReferences) }
                ));
            }
            series.Append(new YValues(
                new NumberReference { Formula = new C.Formula(s.ValueReferences) }
            ));
            series.Append(new BubbleSize(
                new NumberReference { Formula = new C.Formula(s.ValueReferences) }
            ));
            bubbleChart.Append(series);
        }
        bubbleChart.Append(new AxisId { Val = xAxisId });
        bubbleChart.Append(new AxisId { Val = yAxisId });

        var plotArea = new PlotArea(
            new Layout(),
            bubbleChart,
            new ValueAxis(
                new AxisId { Val = xAxisId },
                new Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new Delete { Val = false },
                new AxisPosition { Val = AxisPositionValues.Bottom },
                new CrossingAxis { Val = yAxisId }
            ),
            new ValueAxis(
                new AxisId { Val = yAxisId },
                new Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new Delete { Val = false },
                new AxisPosition { Val = AxisPositionValues.Left },
                new CrossingAxis { Val = xAxisId }
            )
        );
        return plotArea;
    }

    // ── Bar/Column ──

    private static void AppendBarChart(
        PlotArea plotArea, XLChartType chartType,
        IXLChartSeriesCollection seriesCollection, uint catAxisId, uint valAxisId, uint indexOffset)
    {
        var bp = new BarParams(chartType);
        var barChart = new BarChart
        {
            BarDirection = new BarDirection { Val = bp.Direction },
            BarGrouping = new BarGrouping { Val = bp.Grouping }
        };
        foreach (var s in seriesCollection)
        {
            var series = new BarChartSeries
            {
                Index = new C.Index { Val = s.Index + indexOffset },
                Order = new Order { Val = s.Order + indexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendCatAndVal(series, s);
            barChart.Append(series);
        }
        barChart.Append(new AxisId { Val = catAxisId });
        barChart.Append(new AxisId { Val = valAxisId });
        plotArea.Append(barChart);
    }

    // ── Bar3D (Cone, Cylinder, Pyramid, Column3D, 3D Bar variants) ──

    private static void AppendBar3DChart(
        PlotArea plotArea, XLChartType chartType,
        IXLChartSeriesCollection seriesCollection, uint catAxisId, uint valAxisId, uint indexOffset)
    {
        var bp = new BarParams(chartType);
        var bar3DChart = new Bar3DChart
        {
            BarDirection = new BarDirection { Val = bp.Direction },
            BarGrouping = new BarGrouping { Val = bp.Grouping }
        };
        foreach (var s in seriesCollection)
        {
            var series = new BarChartSeries
            {
                Index = new C.Index { Val = s.Index + indexOffset },
                Order = new Order { Val = s.Order + indexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendCatAndVal(series, s);
            bar3DChart.Append(series);
        }
        bar3DChart.Append(new Shape { Val = GetBar3DShape(chartType) });
        bar3DChart.Append(new AxisId { Val = catAxisId });
        bar3DChart.Append(new AxisId { Val = valAxisId });
        plotArea.Append(bar3DChart);
    }

    // ── Line ──

    private static void AppendLineChart(
        PlotArea plotArea, XLChartType chartType,
        IXLChartSeriesCollection seriesCollection, uint catAxisId, uint valAxisId, uint indexOffset)
    {
        var lineChart = new LineChart
        {
            Grouping = new Grouping { Val = GetLineGrouping(chartType) }
        };
        foreach (var s in seriesCollection)
        {
            var series = new LineChartSeries
            {
                Index = new C.Index { Val = s.Index + indexOffset },
                Order = new Order { Val = s.Order + indexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendCatAndVal(series, s);
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

    // ── Radar ──

    private static void AppendRadarChart(
        PlotArea plotArea, XLChartType chartType,
        IXLChartSeriesCollection seriesCollection, uint catAxisId, uint valAxisId, uint indexOffset)
    {
        var radarChart = new RadarChart
        {
            RadarStyle = new RadarStyle
            {
                Val = chartType == XLChartType.RadarFilled ? RadarStyleValues.Filled : RadarStyleValues.Marker
            }
        };
        foreach (var s in seriesCollection)
        {
            var series = new RadarChartSeries
            {
                Index = new C.Index { Val = s.Index + indexOffset },
                Order = new Order { Val = s.Order + indexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendCatAndVal(series, s);
            radarChart.Append(series);
        }
        radarChart.Append(new AxisId { Val = catAxisId });
        radarChart.Append(new AxisId { Val = valAxisId });
        plotArea.Append(radarChart);
    }

    // ── Scatter ──

    private static void AppendScatterChart(
        PlotArea plotArea, XLChartType chartType,
        IXLChartSeriesCollection seriesCollection, uint xAxisId, uint yAxisId, uint indexOffset)
    {
        var scatterChart = new ScatterChart
        {
            ScatterStyle = new ScatterStyle { Val = GetScatterStyle(chartType) }
        };
        foreach (var s in seriesCollection)
        {
            var series = new ScatterChartSeries
            {
                Index = new C.Index { Val = s.Index + indexOffset },
                Order = new Order { Val = s.Order + indexOffset },
                SeriesText = BuildSeriesText(s)
            };
            // Scatter uses XValues + YValues, not CategoryAxisData + Values
            if (s.CategoryReferences != null)
            {
                series.Append(new XValues(
                    new NumberReference { Formula = new C.Formula(s.CategoryReferences) }
                ));
            }
            series.Append(new YValues(
                new NumberReference { Formula = new C.Formula(s.ValueReferences) }
            ));
            scatterChart.Append(series);
        }
        scatterChart.Append(new AxisId { Val = xAxisId });
        scatterChart.Append(new AxisId { Val = yAxisId });
        plotArea.Append(scatterChart);
    }

    // ── Stock ──

    private static void AppendStockChart(
        PlotArea plotArea, IXLChartSeriesCollection seriesCollection,
        uint catAxisId, uint valAxisId, uint indexOffset)
    {
        var stockChart = new StockChart();
        foreach (var s in seriesCollection)
        {
            var series = new LineChartSeries
            {
                Index = new C.Index { Val = s.Index + indexOffset },
                Order = new Order { Val = s.Order + indexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendCatAndVal(series, s);
            stockChart.Append(series);
        }
        stockChart.Append(new AxisId { Val = catAxisId });
        stockChart.Append(new AxisId { Val = valAxisId });
        plotArea.Append(stockChart);
    }

    // ── Surface ──

    private static void AppendSurfaceChart(
        PlotArea plotArea, XLChartType chartType,
        IXLChartSeriesCollection seriesCollection, uint catAxisId, uint valAxisId, uint indexOffset)
    {
        const uint serAxisId = 3u;
        var wireframe = chartType is XLChartType.SurfaceWireframe or XLChartType.SurfaceContourWireframe;

        var surfaceChart = new SurfaceChart();
        if (wireframe)
            surfaceChart.Append(new Wireframe { Val = true });

        foreach (var s in seriesCollection)
        {
            var series = new SurfaceChartSeries
            {
                Index = new C.Index { Val = s.Index + indexOffset },
                Order = new Order { Val = s.Order + indexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendCatAndVal(series, s);
            surfaceChart.Append(series);
        }
        surfaceChart.Append(new AxisId { Val = catAxisId });
        surfaceChart.Append(new AxisId { Val = valAxisId });
        surfaceChart.Append(new AxisId { Val = serAxisId });
        plotArea.Append(surfaceChart);
    }

    // ── Shared helpers ──────────────────────────────────────────────────

    private static void AppendCatAndVal(OpenXmlCompositeElement series, IXLChartSeries s)
    {
        if (s.CategoryReferences != null)
        {
            series.Append(new CategoryAxisData(
                new StringReference { Formula = new C.Formula(s.CategoryReferences) }
            ));
        }
        series.Append(new C.Values(
            new NumberReference { Formula = new C.Formula(s.ValueReferences) }
        ));
    }

    private static SeriesText BuildSeriesText(IXLChartSeries s) =>
        new(new StringReference(
            new StringCache(
                new PointCount { Val = 1 },
                new StringPoint(new NumericValue(s.Name)) { Index = 0 }
            )
        ));

    private static void AppendAnchor(Xdr.WorksheetDrawing worksheetDrawing, XLChart xlChart, A.GraphicData graphicData)
    {
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
                new A.Graphic(graphicData)
            ),
            new Xdr.ClientData()
        );
        worksheetDrawing.Append(anchor);
    }

    /// <summary>
    /// Appends a TwoCellAnchor for an extended chart, wrapping the GraphicFrame in mc:AlternateContent
    /// as required by Excel for Office 2016+ chart types.
    /// </summary>
    private static void AppendExtendedAnchor(Xdr.WorksheetDrawing worksheetDrawing, XLChart xlChart, string chartRelId)
    {
        var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
        var nvpId = nvps.Any()
            ? (UInt32Value)nvps.Max(p => p.Id!.Value) + 1
            : 1U;

        var chartName = string.IsNullOrEmpty(xlChart.Name) ? $"Chart {nvpId}" : xlChart.Name;
        var fromPos = xlChart.Position;
        var toPos = xlChart.SecondPosition;

        var fromCol = fromPos.Column.ToString();
        var fromRow = fromPos.Row.ToString();
        var fromColOff = ((long)(fromPos.ColumnOffset * 9525)).ToString();
        var fromRowOff = ((long)(fromPos.RowOffset * 9525)).ToString();
        var toCol = toPos.Column.ToString();
        var toRow = toPos.Row.ToString();
        var toColOff = ((long)(toPos.ColumnOffset * 9525)).ToString();
        var toRowOff = ((long)(toPos.RowOffset * 9525)).ToString();
        var guid = System.Guid.NewGuid().ToString().ToUpperInvariant();

        // Build the entire TwoCellAnchor as raw XML to ensure namespace declarations
        // are exactly where Excel expects them (not hoisted to the root element).
        var anchorXml = $@"<xdr:twoCellAnchor xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><xdr:from><xdr:col>{fromCol}</xdr:col><xdr:colOff>{fromColOff}</xdr:colOff><xdr:row>{fromRow}</xdr:row><xdr:rowOff>{fromRowOff}</xdr:rowOff></xdr:from><xdr:to><xdr:col>{toCol}</xdr:col><xdr:colOff>{toColOff}</xdr:colOff><xdr:row>{toRow}</xdr:row><xdr:rowOff>{toRowOff}</xdr:rowOff></xdr:to><mc:AlternateContent xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006""><mc:Choice xmlns:cx1=""http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"" Requires=""cx1""><xdr:graphicFrame macro=""""><xdr:nvGraphicFramePr><xdr:cNvPr id=""{nvpId}"" name=""{chartName}""><a:extLst><a:ext uri=""{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}""><a16:creationId xmlns:a16=""http://schemas.microsoft.com/office/drawing/2014/main"" id=""{{{guid}}}""/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""0"" cy=""0""/></xdr:xfrm><a:graphic><a:graphicData uri=""http://schemas.microsoft.com/office/drawing/2014/chartex""><cx:chart xmlns:cx=""http://schemas.microsoft.com/office/drawing/2014/chartex"" r:id=""{chartRelId}""/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback><xdr:sp macro="""" textlink=""""><xdr:nvSpPr><xdr:cNvPr id=""0"" name=""""/><xdr:cNvSpPr><a:spLocks noTextEdit=""1""/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""4572000"" cy=""2743200""/></a:xfrm><a:prstGeom prst=""rect""><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val=""white""/></a:solidFill><a:ln w=""1""><a:solidFill><a:prstClr val=""green""/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow=""clip"" horzOverflow=""clip""/><a:lstStyle/><a:p><a:r><a:rPr lang=""en-US"" sz=""1100""/><a:t>This chart isn't available in your version of Excel.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>";

        var anchor = new Xdr.TwoCellAnchor(anchorXml);
        worksheetDrawing.Append(anchor);
    }

    private static DrawingsPart EnsureDrawingsPart(WorksheetPart worksheetPart, SaveContext context)
    {
        var drawingsPart = worksheetPart.DrawingsPart ??
                           worksheetPart.AddNewPart<DrawingsPart>(context.RelIdGenerator.GetNext(RelType.Workbook));
        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
        return drawingsPart;
    }

    private static void EnsureDrawingElement(
        Worksheet worksheet, XLWorksheetContentManager cm,
        WorksheetPart worksheetPart, DrawingsPart drawingsPart)
    {
        if (!worksheet.OfType<Drawing>().Any())
        {
            var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
            var drawingRef = new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) };
            drawingRef.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            if (tableParts != null)
                worksheet.InsertBefore(drawingRef, tableParts);
            else
                worksheet.AppendChild(drawingRef);
            cm.SetElement(XLWorksheetContents.Drawing, worksheet.Elements<Drawing>().First());
        }
    }

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

    // ── Type classification ─────────────────────────────────────────────

    private static bool IsPieType(XLChartType ct) =>
        ct is XLChartType.Pie or XLChartType.PieExploded
            or XLChartType.Pie3D or XLChartType.PieExploded3D
            or XLChartType.PieToPie or XLChartType.PieToBar;

    private static bool IsDoughnutType(XLChartType ct) =>
        ct is XLChartType.Doughnut or XLChartType.DoughnutExploded;

    private static bool IsAreaType(XLChartType ct) =>
        ct is XLChartType.Area or XLChartType.Area3D
            or XLChartType.AreaStacked or XLChartType.AreaStacked100Percent
            or XLChartType.AreaStacked100Percent3D or XLChartType.AreaStacked3D;

    private static bool IsBubbleType(XLChartType ct) =>
        ct is XLChartType.Bubble or XLChartType.Bubble3D;

    private static bool IsBar3DType(XLChartType ct) =>
        ct is XLChartType.BarClustered3D or XLChartType.BarStacked3D or XLChartType.BarStacked100Percent3D
            or XLChartType.Column3D or XLChartType.ColumnClustered3D
            or XLChartType.ColumnStacked3D or XLChartType.ColumnStacked100Percent3D
            or XLChartType.Cone or XLChartType.ConeClustered
            or XLChartType.ConeHorizontalClustered or XLChartType.ConeHorizontalStacked
            or XLChartType.ConeHorizontalStacked100Percent
            or XLChartType.ConeStacked or XLChartType.ConeStacked100Percent
            or XLChartType.Cylinder or XLChartType.CylinderClustered
            or XLChartType.CylinderHorizontalClustered or XLChartType.CylinderHorizontalStacked
            or XLChartType.CylinderHorizontalStacked100Percent
            or XLChartType.CylinderStacked or XLChartType.CylinderStacked100Percent
            or XLChartType.Pyramid or XLChartType.PyramidClustered
            or XLChartType.PyramidHorizontalClustered or XLChartType.PyramidHorizontalStacked
            or XLChartType.PyramidHorizontalStacked100Percent
            or XLChartType.PyramidStacked or XLChartType.PyramidStacked100Percent;

    private static bool IsLineType(XLChartType ct) =>
        ct is XLChartType.Line or XLChartType.Line3D
            or XLChartType.LineStacked or XLChartType.LineStacked100Percent
            or XLChartType.LineWithMarkers or XLChartType.LineWithMarkersStacked
            or XLChartType.LineWithMarkersStacked100Percent;

    private static bool IsRadarType(XLChartType ct) =>
        ct is XLChartType.Radar or XLChartType.RadarFilled or XLChartType.RadarWithMarkers;

    private static bool IsScatterType(XLChartType ct) =>
        ct is XLChartType.XYScatterMarkers or XLChartType.XYScatterSmoothLinesNoMarkers
            or XLChartType.XYScatterSmoothLinesWithMarkers
            or XLChartType.XYScatterStraightLinesNoMarkers
            or XLChartType.XYScatterStraightLinesWithMarkers;

    private static bool IsStockType(XLChartType ct) =>
        ct is XLChartType.StockHighLowClose or XLChartType.StockOpenHighLowClose
            or XLChartType.StockVolumeHighLowClose or XLChartType.StockVolumeOpenHighLowClose;

    private static bool IsSurfaceType(XLChartType ct) =>
        ct is XLChartType.Surface or XLChartType.SurfaceContour
            or XLChartType.SurfaceContourWireframe or XLChartType.SurfaceWireframe;

    internal static bool IsExtendedType(XLChartType ct) =>
        ct is XLChartType.BoxWhisker or XLChartType.Funnel
            or XLChartType.Sunburst or XLChartType.Treemap
            or XLChartType.Waterfall;

    // ── Mapping helpers ─────────────────────────────────────────────────

    private static GroupingValues GetLineGrouping(XLChartType ct) => ct switch
    {
        XLChartType.LineStacked or XLChartType.LineWithMarkersStacked => GroupingValues.Stacked,
        XLChartType.LineStacked100Percent or XLChartType.LineWithMarkersStacked100Percent => GroupingValues.PercentStacked,
        _ => GroupingValues.Standard
    };

    private static GroupingValues GetAreaGrouping(XLChartType ct) => ct switch
    {
        XLChartType.AreaStacked or XLChartType.AreaStacked3D => GroupingValues.Stacked,
        XLChartType.AreaStacked100Percent or XLChartType.AreaStacked100Percent3D => GroupingValues.PercentStacked,
        _ => GroupingValues.Standard
    };

    private static ShapeValues GetBar3DShape(XLChartType ct) => ct switch
    {
        XLChartType.Cone or XLChartType.ConeClustered
            or XLChartType.ConeHorizontalClustered or XLChartType.ConeHorizontalStacked
            or XLChartType.ConeHorizontalStacked100Percent
            or XLChartType.ConeStacked or XLChartType.ConeStacked100Percent
            => ShapeValues.Cone,
        XLChartType.Cylinder or XLChartType.CylinderClustered
            or XLChartType.CylinderHorizontalClustered or XLChartType.CylinderHorizontalStacked
            or XLChartType.CylinderHorizontalStacked100Percent
            or XLChartType.CylinderStacked or XLChartType.CylinderStacked100Percent
            => ShapeValues.Cylinder,
        XLChartType.Pyramid or XLChartType.PyramidClustered
            or XLChartType.PyramidHorizontalClustered or XLChartType.PyramidHorizontalStacked
            or XLChartType.PyramidHorizontalStacked100Percent
            or XLChartType.PyramidStacked or XLChartType.PyramidStacked100Percent
            => ShapeValues.Pyramid,
        _ => ShapeValues.Box
    };

    private static ScatterStyleValues GetScatterStyle(XLChartType ct) => ct switch
    {
        XLChartType.XYScatterMarkers => ScatterStyleValues.LineMarker,
        XLChartType.XYScatterSmoothLinesNoMarkers => ScatterStyleValues.SmoothMarker,
        XLChartType.XYScatterSmoothLinesWithMarkers => ScatterStyleValues.SmoothMarker,
        XLChartType.XYScatterStraightLinesNoMarkers => ScatterStyleValues.LineMarker,
        XLChartType.XYScatterStraightLinesWithMarkers => ScatterStyleValues.LineMarker,
        _ => ScatterStyleValues.LineMarker
    };

    private readonly struct BarParams
    {
        public BarDirectionValues Direction { get; }
        public BarGroupingValues Grouping { get; }

        public BarParams(XLChartType ct)
        {
            Direction = IsHorizontal(ct) ? BarDirectionValues.Bar : BarDirectionValues.Column;
            Grouping = GetGrouping(ct);
        }

        private static bool IsHorizontal(XLChartType ct) =>
            ct is XLChartType.BarClustered or XLChartType.BarClustered3D
                or XLChartType.BarStacked or XLChartType.BarStacked100Percent
                or XLChartType.BarStacked100Percent3D or XLChartType.BarStacked3D
                or XLChartType.ConeHorizontalClustered or XLChartType.ConeHorizontalStacked
                or XLChartType.ConeHorizontalStacked100Percent
                or XLChartType.CylinderHorizontalClustered or XLChartType.CylinderHorizontalStacked
                or XLChartType.CylinderHorizontalStacked100Percent
                or XLChartType.PyramidHorizontalClustered or XLChartType.PyramidHorizontalStacked
                or XLChartType.PyramidHorizontalStacked100Percent;

        private static BarGroupingValues GetGrouping(XLChartType ct) => ct switch
        {
            XLChartType.BarClustered or XLChartType.BarClustered3D
                or XLChartType.ColumnClustered or XLChartType.ColumnClustered3D
                or XLChartType.ConeClustered or XLChartType.ConeHorizontalClustered
                or XLChartType.CylinderClustered or XLChartType.CylinderHorizontalClustered
                or XLChartType.PyramidClustered or XLChartType.PyramidHorizontalClustered
                => BarGroupingValues.Clustered,
            XLChartType.BarStacked or XLChartType.BarStacked3D
                or XLChartType.ColumnStacked or XLChartType.ColumnStacked3D
                or XLChartType.ConeStacked or XLChartType.ConeHorizontalStacked
                or XLChartType.CylinderStacked or XLChartType.CylinderHorizontalStacked
                or XLChartType.PyramidStacked or XLChartType.PyramidHorizontalStacked
                => BarGroupingValues.Stacked,
            XLChartType.BarStacked100Percent or XLChartType.BarStacked100Percent3D
                or XLChartType.ColumnStacked100Percent or XLChartType.ColumnStacked100Percent3D
                or XLChartType.ConeStacked100Percent or XLChartType.ConeHorizontalStacked100Percent
                or XLChartType.CylinderStacked100Percent or XLChartType.CylinderHorizontalStacked100Percent
                or XLChartType.PyramidStacked100Percent or XLChartType.PyramidHorizontalStacked100Percent
                => BarGroupingValues.PercentStacked,
            _ => BarGroupingValues.Standard
        };
    }
}
