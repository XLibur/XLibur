using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ContentManagers;
using XLibur.Excel.IO.Charts;
using static XLibur.Excel.IO.OpenXmlConst;
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
            {
                // Charts that already exist in the package are not regenerated — that is what keeps
                // the parts of their XML that XLibur does not model intact. Only the properties the
                // caller actually changed are patched into the existing part.
                ChartPatcher.PatchChart(worksheetPart, xlChart);
                continue;
            }

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
        // Remembered so that a later save can patch this part instead of regenerating it.
        xlChart.RelId = chartRelId;

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
        var package = PackageExtensions.GetPackage(worksheetPart.OpenXmlPackage);
#pragma warning restore OOXML0001

        var packagePart = package.CreatePart(
            partUri,
            "application/vnd.ms-office.chartex+xml",
            CompressionOption.Normal);

        // Write chart XML
        var chartSpace = BuildExtendedChartSpace(xlChart);
        using (var stream = packagePart.GetStream(FileMode.Create, FileAccess.Write))
        {
            chartSpace.Save(stream);
        }

        // Create relationship from DrawingsPart to the chart part using relative path
        // Excel requires relative target URIs for extended chart relationships
        var relativeTarget = new Uri("../charts/chartEx" + _extChartCounter + ".xml", UriKind.Relative);
        var drawingsPackagePart = package.GetPart(drawingsPart.Uri);
        drawingsPackagePart.Relationships.Create(
            relativeTarget,
            TargetMode.Internal,
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
        IPackage package,
        IPackagePart chartPart,
        int chartIndex)
#pragma warning restore OOXML0001
    {
        var colorsUri = new Uri($"/xl/charts/colors{chartIndex}.xml", UriKind.Relative);
        var styleUri = new Uri($"/xl/charts/style{chartIndex}.xml", UriKind.Relative);

        // Create color style part
        var colorsPart = package.CreatePart(colorsUri,
            "application/vnd.ms-office.chartcolorstyle+xml",
            CompressionOption.Normal);
        using (var stream = colorsPart.GetStream(FileMode.Create, FileAccess.Write))
        {
            var asm = typeof(ChartWriter).Assembly;
            using var resStream = asm.GetManifestResourceStream("XLibur.Excel.IO.ChartExDefaultColors.xml")!;
            resStream.CopyTo(stream);
        }

        // Create chart style part
        var stylePart = package.CreatePart(styleUri,
            "application/vnd.ms-office.chartstyle+xml",
            CompressionOption.Normal);
        using (var stream = stylePart.GetStream(FileMode.Create, FileAccess.Write))
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
            TargetMode.Internal,
            "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
            "rId1");
        chartPart.Relationships.Create(
            colorsRelTarget,
            TargetMode.Internal,
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
            xml = Regex.Replace(
                xml,
                $@"\s*xmlns:{prefix}=""[^""]*""",
                "",
                RegexOptions.None,
                TimeSpan.FromSeconds(1));
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
        ChartTitleXml.ApplyExtended(cxChart, xlChart);
        cxChart.AppendChild(plotArea);

        var chartSpace = new Cx.ChartSpace();
        chartSpace.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
        chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        chartSpace.AddNamespaceDeclaration("r", RelationshipsNs);
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
            UniqueId = "{" + Guid.NewGuid() + "}"
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

        if (!string.IsNullOrWhiteSpace(s.CategoryReferences))
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


    // ── Standard ChartSpace building ────────────────────────────────────

    private static C.ChartSpace BuildChartSpace(XLChart xlChart)
    {
        var chart = new C.Chart();

        ChartTitleXml.Apply(chart, xlChart);

        chart.Append(BuildPlotArea(xlChart));

        ChartLegendXml.Apply(chart, xlChart.LegendInternal);

        chart.Append(new C.PlotVisibleOnly { Val = true });

        return new C.ChartSpace(chart);
    }

    private const uint CatAxisId = 1u;
    private const uint ValAxisId = 2u;
    private const uint SerAxisId = 3u;
    private const uint SecondaryCatAxisId = 4u;
    private const uint SecondaryValAxisId = 5u;

    /// <summary>
    /// One chart group to emit into the plot area: a chart type, the series plotted with it, and
    /// whether those series hang off the secondary value axis.
    /// </summary>
    private readonly struct PlotGroup(
        XLChart chart, XLChartType chartType, List<XLChartSeries> series,
        bool secondaryAxis, uint indexOffset)
    {
        internal XLChart Chart { get; } = chart;
        internal XLChartType ChartType { get; } = chartType;
        internal List<XLChartSeries> Series { get; } = series;
        internal bool SecondaryAxis { get; } = secondaryAxis;

        /// <summary>Added to each series index so that combo charts do not reuse an index.</summary>
        internal uint IndexOffset { get; } = indexOffset;

        internal uint CatAxisIdOfGroup => SecondaryAxis ? SecondaryCatAxisId : CatAxisId;
        internal uint ValAxisIdOfGroup => SecondaryAxis ? SecondaryValAxisId : ValAxisId;
    }

    private static C.PlotArea BuildPlotArea(XLChart xlChart)
    {
        if (IsPieType(xlChart.ChartType) || IsDoughnutType(xlChart.ChartType))
            return BuildNoAxesPlotArea(xlChart);

        if (IsBubbleType(xlChart.ChartType))
            return BuildBubblePlotArea(xlChart);

        var plotArea = new C.PlotArea();
        plotArea.Append(new C.Layout());

        var groups = BuildPlotGroups(xlChart);
        foreach (var group in groups)
            AppendChartElement(plotArea, group);

        // Every chart group has to precede every axis in CT_PlotArea.
        AppendAxes(plotArea, xlChart, groups.Exists(g => g.SecondaryAxis));

        return plotArea;
    }

    /// <summary>
    /// Splits the chart's series into the groups the plot area needs: one per chart type, and one
    /// more for each chart type that has series bound to the secondary value axis.
    /// </summary>
    private static List<PlotGroup> BuildPlotGroups(XLChart xlChart)
    {
        var groups = new List<PlotGroup>(4);
        AddPlotGroups(groups, xlChart, xlChart.ChartType, xlChart.SeriesInternal.Items, 0);

        if (xlChart.SecondaryChartType.HasValue && xlChart.SecondarySeries.Count > 0)
        {
            AddPlotGroups(groups, xlChart, xlChart.SecondaryChartType.Value,
                xlChart.SecondarySeriesInternal.Items, (uint)xlChart.Series.Count);
        }

        return groups;
    }

    private static void AddPlotGroups(
        List<PlotGroup> groups, XLChart xlChart, XLChartType chartType,
        IReadOnlyList<XLChartSeries> series, uint indexOffset)
    {
        if (series.Count == 0)
            return;

        if (!SupportsSecondaryAxis(chartType))
        {
            groups.Add(new PlotGroup(xlChart, chartType, [.. series], secondaryAxis: false, indexOffset));
            return;
        }

        var primary = new List<XLChartSeries>(series.Count);
        var secondary = new List<XLChartSeries>();
        foreach (var s in series)
            (s.UseSecondaryAxis ? secondary : primary).Add(s);

        if (primary.Count > 0)
            groups.Add(new PlotGroup(xlChart, chartType, primary, secondaryAxis: false, indexOffset));
        if (secondary.Count > 0)
            groups.Add(new PlotGroup(xlChart, chartType, secondary, secondaryAxis: true, indexOffset));
    }

    /// <summary>
    /// Whether <see cref="IXLChartSeries.UseSecondaryAxis"/> means anything for a chart type. The
    /// types with two value axes (scatter, bubble), no value axis (pie, doughnut) or a series axis
    /// (surface) have no secondary value axis to bind to.
    /// </summary>
    private static bool SupportsSecondaryAxis(XLChartType ct) =>
        !IsScatterType(ct) && !IsBubbleType(ct) && !IsSurfaceType(ct)
        && !IsPieType(ct) && !IsDoughnutType(ct);

    private static void AppendAxes(C.PlotArea plotArea, XLChart xlChart, bool hasSecondaryAxis)
    {
        var primaryChartType = xlChart.ChartType;
        var horizontal = xlChart.CategoryAxisInternal;
        var vertical = xlChart.ValueAxisInternal;

        if (IsScatterType(primaryChartType))
        {
            // Scatter uses two ValueAxis (X and Y)
            plotArea.Append(BuildValueAxis(CatAxisId, ValAxisId, C.AxisPositionValues.Bottom, horizontal));
            plotArea.Append(BuildValueAxis(ValAxisId, CatAxisId, C.AxisPositionValues.Left, vertical));
        }
        else if (IsSurfaceType(primaryChartType))
        {
            plotArea.Append(BuildCategoryAxis(CatAxisId, ValAxisId, horizontal));
            plotArea.Append(BuildValueAxis(ValAxisId, CatAxisId, C.AxisPositionValues.Left, vertical));
            plotArea.Append(new C.SeriesAxis(
                new C.AxisId { Val = SerAxisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                new C.CrossingAxis { Val = ValAxisId }
            ));
        }
        else
        {
            plotArea.Append(BuildCategoryAxis(CatAxisId, ValAxisId, horizontal));
            plotArea.Append(BuildValueAxis(ValAxisId, CatAxisId, C.AxisPositionValues.Left, vertical));
        }

        if (!hasSecondaryAxis)
            return;

        // The secondary group needs its own axis pair. Excel hides the extra category axis and puts
        // the extra value axis on the right, crossing the category axis at its maximum.
        plotArea.Append(BuildCategoryAxis(SecondaryCatAxisId, SecondaryValAxisId,
            model: null, deleted: true));
        plotArea.Append(BuildValueAxis(SecondaryValAxisId, SecondaryCatAxisId,
            C.AxisPositionValues.Right, xlChart.SecondaryValueAxisInternal, crossesMaximum: true));
    }

    /// <summary>
    /// Builds a <c>c:catAx</c>. Its children follow the CT_CatAx order: identity and position first,
    /// then the optional gridlines, title and number format, then the crossing axis.
    /// </summary>
    /// <param name="axisId">The identifier this axis is known by.</param>
    /// <param name="crossingAxisId">The identifier of the axis this one crosses.</param>
    /// <param name="model">
    /// The axis model to take the optional properties from, or <c>null</c> for the hidden helper axis
    /// of a secondary group, which has no public counterpart.
    /// </param>
    /// <param name="deleted">Whether the axis is hidden regardless of what the model says.</param>
    private static C.CategoryAxis BuildCategoryAxis(
        uint axisId, uint crossingAxisId, XLChartAxis? model, bool deleted = false)
    {
        var axis = new C.CategoryAxis();
        axis.Append(new C.AxisId { Val = axisId });
        axis.Append(model == null
            ? new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax })
            : ChartFormatting.BuildScaling(model));
        axis.Append(new C.Delete { Val = deleted || model is { Visible: false } });
        axis.Append(new C.AxisPosition { Val = C.AxisPositionValues.Bottom });

        if (model != null)
            ChartFormatting.AppendAxisBody(axis, model);

        axis.Append(new C.CrossingAxis { Val = crossingAxisId });
        return axis;
    }

    private static C.ValueAxis BuildValueAxis(
        uint axisId, uint crossingAxisId, C.AxisPositionValues position, XLChartAxis model,
        bool crossesMaximum = false)
    {
        var axis = new C.ValueAxis();
        axis.Append(new C.AxisId { Val = axisId });
        axis.Append(ChartFormatting.BuildScaling(model));
        axis.Append(new C.Delete { Val = !model.Visible });
        axis.Append(new C.AxisPosition { Val = position });

        ChartFormatting.AppendAxisBody(axis, model);

        axis.Append(new C.CrossingAxis { Val = crossingAxisId });
        if (crossesMaximum)
            axis.Append(new C.Crosses { Val = C.CrossesValues.Maximum });

        ChartFormatting.AppendAxisUnits(axis, model);
        return axis;
    }

    private static void AppendChartElement(C.PlotArea plotArea, PlotGroup group)
    {
        var chartType = group.ChartType;

        if (IsAreaType(chartType))
            AppendAreaChart(plotArea, group);
        else if (IsLineType(chartType))
            AppendLineChart(plotArea, group);
        else if (IsRadarType(chartType))
            AppendRadarChart(plotArea, group);
        else if (IsScatterType(chartType))
            AppendScatterChart(plotArea, group);
        else if (IsStockType(chartType))
            AppendStockChart(plotArea, group);
        else if (IsSurfaceType(chartType))
            AppendSurfaceChart(plotArea, group);
        else if (IsBar3DType(chartType))
            AppendBar3DChart(plotArea, group);
        else
            AppendBarChart(plotArea, group);
    }

    // ── Pie / Doughnut (no axes) ──

    private static C.PlotArea BuildNoAxesPlotArea(XLChart xlChart)
    {
        var isDoughnut = IsDoughnutType(xlChart.ChartType);
        OpenXmlCompositeElement chartElement = isDoughnut
            ? new C.DoughnutChart()
            : new C.PieChart();

        foreach (var s in xlChart.SeriesInternal.Items)
        {
            var series = new C.PieChartSeries
            {
                Index = new C.Index { Val = s.Index },
                Order = new C.Order { Val = s.Order },
                SeriesText = BuildSeriesText(s)
            };
            AppendShapeProperties(series, s);
            AppendSeriesDataLabels(series, s, xlChart.ChartType);
            AppendCatAndVal(series, s);
            chartElement.Append(series);
        }

        AppendGroupDataLabels(chartElement, xlChart, xlChart.ChartType);

        // c:holeSize is required by CT_DoughnutChart; 75% is the size Excel gives a new doughnut.
        if (isDoughnut)
            chartElement.Append(new C.HoleSize { Val = 75 });

        return new C.PlotArea(new C.Layout(), chartElement);
    }

    // ── Area ──

    private static void AppendAreaChart(C.PlotArea plotArea, PlotGroup group)
    {
        var areaChart = new C.AreaChart
        {
            Grouping = new C.Grouping { Val = GetAreaGrouping(group.ChartType) }
        };
        foreach (var s in group.Series)
        {
            var series = new C.AreaChartSeries
            {
                Index = new C.Index { Val = s.Index + group.IndexOffset },
                Order = new C.Order { Val = s.Order + group.IndexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendShapeProperties(series, s);
            AppendSeriesDataLabels(series, s, group.ChartType);
            AppendCatAndVal(series, s);
            areaChart.Append(series);
        }
        AppendGroupDataLabels(areaChart, group.Chart, group.ChartType);
        AppendAxisIds(areaChart, group);
        plotArea.Append(areaChart);
    }

    // ── Bubble ──

    private static C.PlotArea BuildBubblePlotArea(XLChart xlChart)
    {
        // Bubble charts use XValues + YValues + BubbleSize, and two ValueAxis (like scatter).
        // CategoryReferences = X values, ValueReferences = Y values.
        // For simplicity, bubble size defaults to the Y values if no separate size data.
        const uint xAxisId = 1u;
        const uint yAxisId = 2u;

        var bubbleChart = new C.BubbleChart();
        foreach (var s in xlChart.SeriesInternal.Items)
        {
            var series = new C.BubbleChartSeries
            {
                Index = new C.Index { Val = s.Index },
                Order = new C.Order { Val = s.Order },
                SeriesText = BuildSeriesText(s)
            };
            AppendShapeProperties(series, s);
            AppendSeriesDataLabels(series, s, xlChart.ChartType);
            AppendXAndY(series, s);
            series.Append(new C.BubbleSize(
                new C.NumberReference { Formula = new C.Formula(s.ValueReferences) }
            ));
            bubbleChart.Append(series);
        }
        AppendGroupDataLabels(bubbleChart, xlChart, xlChart.ChartType);
        bubbleChart.Append(new C.AxisId { Val = xAxisId });
        bubbleChart.Append(new C.AxisId { Val = yAxisId });

        var plotArea = new C.PlotArea(
            new C.Layout(),
            bubbleChart,
            BuildValueAxis(xAxisId, yAxisId, C.AxisPositionValues.Bottom, xlChart.CategoryAxisInternal),
            BuildValueAxis(yAxisId, xAxisId, C.AxisPositionValues.Left, xlChart.ValueAxisInternal)
        );
        return plotArea;
    }

    // ── Bar/Column ──

    private static void AppendBarChart(C.PlotArea plotArea, PlotGroup group)
    {
        var bp = new BarParams(group.ChartType);
        var barChart = new C.BarChart
        {
            BarDirection = new C.BarDirection { Val = bp.Direction },
            BarGrouping = new C.BarGrouping { Val = bp.Grouping }
        };
        foreach (var s in group.Series)
        {
            var series = new C.BarChartSeries
            {
                Index = new C.Index { Val = s.Index + group.IndexOffset },
                Order = new C.Order { Val = s.Order + group.IndexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendShapeProperties(series, s);
            AppendSeriesDataLabels(series, s, group.ChartType);
            AppendCatAndVal(series, s);
            barChart.Append(series);
        }
        AppendGroupDataLabels(barChart, group.Chart, group.ChartType);
        AppendAxisIds(barChart, group);
        plotArea.Append(barChart);
    }

    // ── Bar3D (Cone, Cylinder, Pyramid, Column3D, 3D Bar variants) ──

    private static void AppendBar3DChart(C.PlotArea plotArea, PlotGroup group)
    {
        var bp = new BarParams(group.ChartType);
        var bar3DChart = new C.Bar3DChart
        {
            BarDirection = new C.BarDirection { Val = bp.Direction },
            BarGrouping = new C.BarGrouping { Val = bp.Grouping }
        };
        foreach (var s in group.Series)
        {
            var series = new C.BarChartSeries
            {
                Index = new C.Index { Val = s.Index + group.IndexOffset },
                Order = new C.Order { Val = s.Order + group.IndexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendShapeProperties(series, s);
            AppendSeriesDataLabels(series, s, group.ChartType);
            AppendCatAndVal(series, s);
            bar3DChart.Append(series);
        }
        AppendGroupDataLabels(bar3DChart, group.Chart, group.ChartType);
        bar3DChart.Append(new C.Shape { Val = GetBar3DShape(group.ChartType) });
        AppendAxisIds(bar3DChart, group);
        plotArea.Append(bar3DChart);
    }

    // ── Line ──

    private static void AppendLineChart(C.PlotArea plotArea, PlotGroup group)
    {
        var lineChart = new C.LineChart
        {
            Grouping = new C.Grouping { Val = GetLineGrouping(group.ChartType) }
        };
        var markersByChartType = group.ChartType is XLChartType.LineWithMarkers
            or XLChartType.LineWithMarkersStacked
            or XLChartType.LineWithMarkersStacked100Percent;

        foreach (var s in group.Series)
        {
            var series = new C.LineChartSeries
            {
                Index = new C.Index { Val = s.Index + group.IndexOffset },
                Order = new C.Order { Val = s.Order + group.IndexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendShapeProperties(series, s);
            AppendMarker(series, s, markersByChartType);
            AppendSeriesDataLabels(series, s, group.ChartType);
            AppendCatAndVal(series, s);
            AppendSmooth(series, s, smoothByChartType: false);
            lineChart.Append(series);
        }
        AppendGroupDataLabels(lineChart, group.Chart, group.ChartType);
        AppendAxisIds(lineChart, group);
        plotArea.Append(lineChart);
    }

    // ── Radar ──

    private static void AppendRadarChart(C.PlotArea plotArea, PlotGroup group)
    {
        var radarChart = new C.RadarChart
        {
            RadarStyle = new C.RadarStyle
            {
                Val = group.ChartType == XLChartType.RadarFilled
                    ? C.RadarStyleValues.Filled
                    : C.RadarStyleValues.Marker
            }
        };
        foreach (var s in group.Series)
        {
            var series = new C.RadarChartSeries
            {
                Index = new C.Index { Val = s.Index + group.IndexOffset },
                Order = new C.Order { Val = s.Order + group.IndexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendShapeProperties(series, s);
            AppendMarker(series, s, autoSymbol: false);
            AppendSeriesDataLabels(series, s, group.ChartType);
            AppendCatAndVal(series, s);
            radarChart.Append(series);
        }
        AppendGroupDataLabels(radarChart, group.Chart, group.ChartType);
        AppendAxisIds(radarChart, group);
        plotArea.Append(radarChart);
    }

    // ── Scatter ──

    private static void AppendScatterChart(C.PlotArea plotArea, PlotGroup group)
    {
        var scatterChart = new C.ScatterChart
        {
            ScatterStyle = new C.ScatterStyle { Val = GetScatterStyle(group.ChartType) }
        };
        var smoothByChartType = group.ChartType is XLChartType.XYScatterSmoothLinesNoMarkers
            or XLChartType.XYScatterSmoothLinesWithMarkers;

        foreach (var s in group.Series)
        {
            var series = new C.ScatterChartSeries
            {
                Index = new C.Index { Val = s.Index + group.IndexOffset },
                Order = new C.Order { Val = s.Order + group.IndexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendShapeProperties(series, s);
            AppendMarker(series, s, autoSymbol: false);
            AppendSeriesDataLabels(series, s, group.ChartType);
            // Scatter uses XValues + YValues, not CategoryAxisData + Values
            AppendXAndY(series, s);
            AppendSmooth(series, s, smoothByChartType);
            scatterChart.Append(series);
        }
        AppendGroupDataLabels(scatterChart, group.Chart, group.ChartType);
        AppendAxisIds(scatterChart, group);
        plotArea.Append(scatterChart);
    }

    // ── Stock ──

    private static void AppendStockChart(C.PlotArea plotArea, PlotGroup group)
    {
        var stockChart = new C.StockChart();
        foreach (var s in group.Series)
        {
            var series = new C.LineChartSeries
            {
                Index = new C.Index { Val = s.Index + group.IndexOffset },
                Order = new C.Order { Val = s.Order + group.IndexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendShapeProperties(series, s);
            AppendMarker(series, s, autoSymbol: false);
            AppendSeriesDataLabels(series, s, group.ChartType);
            AppendCatAndVal(series, s);
            AppendSmooth(series, s, smoothByChartType: false);
            stockChart.Append(series);
        }
        AppendGroupDataLabels(stockChart, group.Chart, group.ChartType);
        AppendAxisIds(stockChart, group);
        plotArea.Append(stockChart);
    }

    // ── Surface ──

    private static void AppendSurfaceChart(C.PlotArea plotArea, PlotGroup group)
    {
        var wireframe = group.ChartType is XLChartType.SurfaceWireframe
            or XLChartType.SurfaceContourWireframe;

        var surfaceChart = new C.SurfaceChart();
        if (wireframe)
            surfaceChart.Append(new C.Wireframe { Val = true });

        foreach (var s in group.Series)
        {
            var series = new C.SurfaceChartSeries
            {
                Index = new C.Index { Val = s.Index + group.IndexOffset },
                Order = new C.Order { Val = s.Order + group.IndexOffset },
                SeriesText = BuildSeriesText(s)
            };
            AppendShapeProperties(series, s);
            AppendCatAndVal(series, s);
            surfaceChart.Append(series);
        }
        AppendAxisIds(surfaceChart, group);
        surfaceChart.Append(new C.AxisId { Val = SerAxisId });
        plotArea.Append(surfaceChart);
    }

    // ── Shared helpers ──────────────────────────────────────────────────

    private static void AppendAxisIds(OpenXmlCompositeElement chartElement, PlotGroup group)
    {
        chartElement.Append(new C.AxisId { Val = group.CatAxisIdOfGroup });
        chartElement.Append(new C.AxisId { Val = group.ValAxisIdOfGroup });
    }

    private static void AppendCatAndVal(OpenXmlCompositeElement series, XLChartSeries s)
    {
        if (!string.IsNullOrWhiteSpace(s.CategoryReferences))
        {
            series.Append(new C.CategoryAxisData(
                new C.StringReference { Formula = new C.Formula(s.CategoryReferences) }
            ));
        }
        series.Append(new C.Values(
            new C.NumberReference { Formula = new C.Formula(s.ValueReferences) }
        ));
    }

    private static void AppendXAndY(OpenXmlCompositeElement series, XLChartSeries s)
    {
        if (!string.IsNullOrWhiteSpace(s.CategoryReferences))
        {
            series.Append(new C.XValues(
                new C.NumberReference { Formula = new C.Formula(s.CategoryReferences) }
            ));
        }
        series.Append(new C.YValues(
            new C.NumberReference { Formula = new C.Formula(s.ValueReferences) }
        ));
    }

    private static void AppendShapeProperties(OpenXmlCompositeElement series, XLChartSeries s)
    {
        var shapeProperties = ChartFormatting.BuildSeriesShapeProperties(s);
        if (shapeProperties != null)
            series.Append(shapeProperties);
    }

    private static void AppendMarker(OpenXmlCompositeElement series, XLChartSeries s, bool autoSymbol)
    {
        var marker = ChartFormatting.BuildMarker(s, autoSymbol);
        if (marker != null)
            series.Append(marker);
    }

    private static void AppendSmooth(OpenXmlCompositeElement series, XLChartSeries s, bool smoothByChartType)
    {
        var smooth = ChartFormatting.BuildSmooth(s, smoothByChartType);
        if (smooth != null)
            series.Append(smooth);
    }

    /// <summary>
    /// Appends the series' own <c>c:dLbls</c>. Must be called after <c>c:marker</c> and before
    /// <c>c:cat</c>/<c>c:val</c>.
    /// </summary>
    private static void AppendSeriesDataLabels(
        OpenXmlCompositeElement series, XLChartSeries s, XLChartType chartType)
    {
        var dataLabels = ChartFormatting.BuildDataLabels(s.DataLabelsInternal, chartType);
        if (dataLabels != null)
            series.Append(dataLabels);
    }

    /// <summary>
    /// Appends the chart-wide <c>c:dLbls</c> to a chart group. Must be called after every
    /// <c>c:ser</c> and before the group's remaining children.
    /// </summary>
    private static void AppendGroupDataLabels(
        OpenXmlCompositeElement chartElement, XLChart xlChart, XLChartType chartType)
    {
        var dataLabels = ChartFormatting.BuildDataLabels(xlChart.DataLabelsInternal, chartType);
        if (dataLabels != null)
            chartElement.Append(dataLabels);
    }

    /// <summary>
    /// Writes the series name as the literal <c>&lt;c:tx&gt;&lt;c:v&gt;</c> form. The alternative,
    /// <c>c:strRef</c>, is for names that come from a cell and requires a <c>c:f</c> formula, which
    /// a literal name does not have.
    /// </summary>
    private static C.SeriesText BuildSeriesText(XLChartSeries s) =>
        new(new C.NumericValue(s.Name));

    /// <summary>EMU per pixel, the unit the drawing markers and extents are stored in.</summary>
    private const double EmuPerPixel = 9525;

    private static void AppendAnchor(Xdr.WorksheetDrawing worksheetDrawing, XLChart xlChart, A.GraphicData graphicData)
    {
        var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
        var nvpId = nvps.Any()
            ? (UInt32Value)nvps.Max(p => p.Id!.Value) + 1
            : 1U;

        var graphicFrame = new Xdr.GraphicFrame(
            new Xdr.NonVisualGraphicFrameProperties(
                new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = xlChart.Name },
                new Xdr.NonVisualGraphicFrameDrawingProperties()
            ),
            new Xdr.Transform(
                new A.Offset { X = 0, Y = 0 },
                new A.Extents { Cx = 0, Cy = 0 }
            ),
            new A.Graphic(graphicData)
        );

        OpenXmlCompositeElement anchor = xlChart.Anchor switch
        {
            XLDrawingAnchor.MoveWithCells => new Xdr.OneCellAnchor(
                BuildFromMarker(xlChart.Position),
                BuildExtent(xlChart),
                graphicFrame,
                new Xdr.ClientData()),

            XLDrawingAnchor.Absolute => new Xdr.AbsoluteAnchor(
                new Xdr.Position
                {
                    X = ToEmu(xlChart.Left),
                    Y = ToEmu(xlChart.Top)
                },
                BuildExtent(xlChart),
                graphicFrame,
                new Xdr.ClientData()),

            _ => new Xdr.TwoCellAnchor(
                BuildFromMarker(xlChart.Position),
                BuildToMarker(xlChart.SecondPosition),
                graphicFrame,
                new Xdr.ClientData())
        };

        worksheetDrawing.Append(anchor);
    }

    private static Xdr.FromMarker BuildFromMarker(IXLDrawingPosition position) => new()
    {
        ColumnId = new Xdr.ColumnId(position.Column.ToString()),
        RowId = new Xdr.RowId(position.Row.ToString()),
        ColumnOffset = new Xdr.ColumnOffset(ToEmu(position.ColumnOffset).ToString()),
        RowOffset = new Xdr.RowOffset(ToEmu(position.RowOffset).ToString())
    };

    private static Xdr.ToMarker BuildToMarker(IXLDrawingPosition position) => new()
    {
        ColumnId = new Xdr.ColumnId(position.Column.ToString()),
        RowId = new Xdr.RowId(position.Row.ToString()),
        ColumnOffset = new Xdr.ColumnOffset(ToEmu(position.ColumnOffset).ToString()),
        RowOffset = new Xdr.RowOffset(ToEmu(position.RowOffset).ToString())
    };

    private static Xdr.Extent BuildExtent(XLChart xlChart) => new()
    {
        Cx = ToEmu(xlChart.Width),
        Cy = ToEmu(xlChart.Height)
    };

    private static long ToEmu(double pixels) => (long)(pixels * EmuPerPixel);

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
        var guid = Guid.NewGuid().ToString().ToUpperInvariant();

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
            drawingRef.AddNamespaceDeclaration("r", RelationshipsNs);
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
                nd.Value.Equals(RelationshipsNs)))
            worksheetDrawing.AddNamespaceDeclaration("r", RelationshipsNs);
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

    private static C.GroupingValues GetLineGrouping(XLChartType ct) => ct switch
    {
        XLChartType.LineStacked or XLChartType.LineWithMarkersStacked => C.GroupingValues.Stacked,
        XLChartType.LineStacked100Percent or XLChartType.LineWithMarkersStacked100Percent => C.GroupingValues.PercentStacked,
        _ => C.GroupingValues.Standard
    };

    private static C.GroupingValues GetAreaGrouping(XLChartType ct) => ct switch
    {
        XLChartType.AreaStacked or XLChartType.AreaStacked3D => C.GroupingValues.Stacked,
        XLChartType.AreaStacked100Percent or XLChartType.AreaStacked100Percent3D => C.GroupingValues.PercentStacked,
        _ => C.GroupingValues.Standard
    };

    private static C.ShapeValues GetBar3DShape(XLChartType ct) => ct switch
    {
        XLChartType.Cone or XLChartType.ConeClustered
            or XLChartType.ConeHorizontalClustered or XLChartType.ConeHorizontalStacked
            or XLChartType.ConeHorizontalStacked100Percent
            or XLChartType.ConeStacked or XLChartType.ConeStacked100Percent
            => C.ShapeValues.Cone,
        XLChartType.Cylinder or XLChartType.CylinderClustered
            or XLChartType.CylinderHorizontalClustered or XLChartType.CylinderHorizontalStacked
            or XLChartType.CylinderHorizontalStacked100Percent
            or XLChartType.CylinderStacked or XLChartType.CylinderStacked100Percent
            => C.ShapeValues.Cylinder,
        XLChartType.Pyramid or XLChartType.PyramidClustered
            or XLChartType.PyramidHorizontalClustered or XLChartType.PyramidHorizontalStacked
            or XLChartType.PyramidHorizontalStacked100Percent
            or XLChartType.PyramidStacked or XLChartType.PyramidStacked100Percent
            => C.ShapeValues.Pyramid,
        _ => C.ShapeValues.Box
    };

    private static C.ScatterStyleValues GetScatterStyle(XLChartType ct) => ct switch
    {
        XLChartType.XYScatterMarkers => C.ScatterStyleValues.LineMarker,
        XLChartType.XYScatterSmoothLinesNoMarkers => C.ScatterStyleValues.SmoothMarker,
        XLChartType.XYScatterSmoothLinesWithMarkers => C.ScatterStyleValues.SmoothMarker,
        XLChartType.XYScatterStraightLinesNoMarkers => C.ScatterStyleValues.LineMarker,
        XLChartType.XYScatterStraightLinesWithMarkers => C.ScatterStyleValues.LineMarker,
        _ => C.ScatterStyleValues.LineMarker
    };

    private readonly struct BarParams
    {
        public C.BarDirectionValues Direction { get; }
        public C.BarGroupingValues Grouping { get; }

        public BarParams(XLChartType ct)
        {
            Direction = IsHorizontal(ct) ? C.BarDirectionValues.Bar : C.BarDirectionValues.Column;
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

        private static C.BarGroupingValues GetGrouping(XLChartType ct) => ct switch
        {
            XLChartType.BarClustered or XLChartType.BarClustered3D
                or XLChartType.ColumnClustered or XLChartType.ColumnClustered3D
                or XLChartType.ConeClustered or XLChartType.ConeHorizontalClustered
                or XLChartType.CylinderClustered or XLChartType.CylinderHorizontalClustered
                or XLChartType.PyramidClustered or XLChartType.PyramidHorizontalClustered
                => C.BarGroupingValues.Clustered,
            XLChartType.BarStacked or XLChartType.BarStacked3D
                or XLChartType.ColumnStacked or XLChartType.ColumnStacked3D
                or XLChartType.ConeStacked or XLChartType.ConeHorizontalStacked
                or XLChartType.CylinderStacked or XLChartType.CylinderHorizontalStacked
                or XLChartType.PyramidStacked or XLChartType.PyramidHorizontalStacked
                => C.BarGroupingValues.Stacked,
            XLChartType.BarStacked100Percent or XLChartType.BarStacked100Percent3D
                or XLChartType.ColumnStacked100Percent or XLChartType.ColumnStacked100Percent3D
                or XLChartType.ConeStacked100Percent or XLChartType.ConeHorizontalStacked100Percent
                or XLChartType.CylinderStacked100Percent or XLChartType.CylinderHorizontalStacked100Percent
                or XLChartType.PyramidStacked100Percent or XLChartType.PyramidHorizontalStacked100Percent
                => C.BarGroupingValues.PercentStacked,
            _ => C.BarGroupingValues.Standard
        };
    }
}
